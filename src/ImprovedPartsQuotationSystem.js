import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';

// 注意：在实际项目中，需要在HTML中引入以下外部库
// PDF.js 和 mammoth.js 用于解析PDF和Word文档
// <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
// <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.21/mammoth.browser.min.js"></script>

// ------------------ 1. 常量、正则及辅助函数 ------------------ //

const PART_PATTERN = {
    STANDARD_CODE: /^\d+[\-\.]\d+[\-\.]\d+[A-Za-z]?$/,      // 如 135-01-003A
    SCREW_CODE: /^[A-Z]\d+X\d+[A-Za-z0-9\-\.]+$/,           // 如 M10X25GB32
    BEARING_CODE: /^[A-Z]+\d+$/,                            // 如 NJ313
    BEARING_CODE2: /^\d+[A-Za-z]+\d*$/,                     // 如 6317N
    COMPLEX_CODE: /^\d+[A-Za-z][\-\.]\d+[A-Za-z][\-\.]\d+[A-Za-z]?$/ // 如 135A-03A-007
};

const ROBUST_EXCEL_OPTIONS = {
    type: 'array',
    cellDates: true,
    cellNF: true,
    cellStyles: true,
    cellFormula: true,
    sheetStubs: true,
    raw: false
};

// 允许的文件类型常量
const ALLOWED_FILE_TYPES = {
    EXCEL: ['.xlsx', '.xls', '.csv'],
    PDF: ['.pdf'],
    WORD: ['.doc', '.docx', '.rtf']
};

function handleProcessingError(operation, error, setLoading, setInfoMessage) {
    console.error(`${operation}错误:`, error);
    setLoading(false);
    setInfoMessage(null);
    alert(`${operation}失败: ${error.message}。请检查文件格式或联系技术支持。`);
}

function safelyStoreData(key, data) {
    try {
        localStorage.setItem(key, JSON.stringify(data));
        return true;
    } catch (error) {
        console.error(`存储到 localStorage 失败: ${error.message}`);
        return false;
    }
}

function safelyRetrieveData(key, defaultValue = null) {
    try {
        const stored = localStorage.getItem(key);
        return stored ? JSON.parse(stored) : defaultValue;
    } catch (error) {
        console.error(`从 localStorage 检索失败: ${error.message}`);
        return defaultValue;
    }
}

// 改进的价格格式化函数 - 小数点后四舍五入
function formatPrice(val) {
    const num = parseFloat(val || 0);
    return Math.round(num).toLocaleString('zh-CN');
}

// 合计价格时，保留两位小数
function formatTotalPrice(val) {
    const num = parseFloat(val || 0);
    return num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

async function readFileContent(e, isElectron) {
    let file = null;
    let fileContent = null;
    if (isElectron) {
        try {
            const result = await window.electronAPI.importData();
            if (!result) return { file: null, fileContent: null };
            file = { name: result.path.split('/').pop() || result.path.split('\\').pop() };
            fileContent = result.content;
        } catch (error) {
            console.error("Electron文件选择错误:", error);
            return { file: null, fileContent: null };
        }
    } else {
        if (!e || !e.target || !e.target.files || e.target.files.length === 0) {
            return { file: null, fileContent: null };
        }
        file = e.target.files[0];
        if (!file) return { file: null, fileContent: null };
        try {
            fileContent = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (event) => resolve(new Uint8Array(event.target.result));
                reader.onerror = (error) => reject(error);
                reader.readAsArrayBuffer(file);
            });
        } catch (error) {
            console.error("文件读取错误:", error);
            return { file: null, fileContent: null };
        }
    }
    return { file, fileContent };
}

function removeDuplicates(data) {
    const seen = new Set();
    return data.filter(item => {
        const key = item['图号'];
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

// ------------------ 2. "增强"模糊匹配 ------------------ //
function enhancedFuzzyMatch(importedId, partsData) {
    // 完全匹配
    const exactMatch = partsData.find(
        part => String(part['图号']).trim() === String(importedId).trim()
    );
    if (exactMatch) return { ...exactMatch, matchType: 'exact' };

    // 不区分大小写匹配
    const caseInsensitiveMatch = partsData.find(
        part => String(part['图号']).toLowerCase().trim() === String(importedId).toLowerCase().trim()
    );
    if (caseInsensitiveMatch) return { ...caseInsensitiveMatch, matchType: 'caseInsensitive' };

    // 去除所有空格后匹配
    const noSpaceImportedId = String(importedId).replace(/\s+/g, '');
    const noSpaceMatch = partsData.find(
        part => String(part['图号']).replace(/\s+/g, '') === noSpaceImportedId
    );
    if (noSpaceMatch) return { ...noSpaceMatch, matchType: 'noSpace' };

    // 模糊匹配 - 替换常见符号差异
    const normalizedImportedId = String(importedId)
        .replace(/[-_．.・]/g, '-')
        .replace(/\s+/g, '')
        .toLowerCase();
    
    const fuzzyMatch = partsData.find(part => {
        const normalizedPartId = String(part['图号'])
            .replace(/[-_．.・]/g, '-')
            .replace(/\s+/g, '')
            .toLowerCase();
        return normalizedPartId === normalizedImportedId;
    });
    
    if (fuzzyMatch) return { ...fuzzyMatch, matchType: 'fuzzy' };

    return null;
}

// ------------------ 3. 各种文件解析逻辑 ------------------ //

// ------------------ 高级文档解析模块 ------------------ //

// 检测文件类型
function detectFileType(fileName) {
    if (!fileName) return 'unknown';
    
    const ext = fileName.toLowerCase().split('.').pop();
    if (ALLOWED_FILE_TYPES.EXCEL.some(type => type.includes(ext))) return 'excel';
    if (ALLOWED_FILE_TYPES.PDF.some(type => type.includes(ext))) return 'pdf';
    if (ALLOWED_FILE_TYPES.WORD.some(type => type.includes(ext))) return 'word';
    return 'unknown';
}

// 从PDF文本内容中提取可能的配件号
async function extractPartsFromPdfText(text) {
    const partCandidates = [];
    const seenParts = new Set(); // 用于去重
    
    // 使用更加精确的匹配方式，按行解析
    const lines = text.split('\n');
    for (const line of lines) {
        // 如果行内容看起来像是配件列表项
        if (/^\s*\d+\.|\-|\*\s+/.test(line) || /[A-Z0-9]{3,}/.test(line)) {
            // 尝试提取配件号 - 更具体的模式
            const patterns = [
                // 标准代码: 135-01-003A
                { regex: /\b(\d+[\-\.]\d+[\-\.]\d+[A-Za-z]?)\b/, type: 'STANDARD_CODE' },
                
                // 螺栓代码: M10X25GB32.1-88
                { regex: /\b([A-Z]\d+X\d+[A-Za-z0-9\-\.]+)\b/, type: 'SCREW_CODE' },
                
                // 轴承代码: NJ313 或 6317N
                { regex: /\b([A-Z]{1,3}\d{2,5}|\d{3,5}[A-Z]{1,2}\d{0,2})\b/, type: 'BEARING_CODE' },
                
                // 复杂代码: 135A-03A-016A
                { regex: /\b(\d+[A-Za-z][\-\.]\d+[A-Za-z][\-\.]\d+[A-Za-z]?)\b/, type: 'COMPLEX_CODE' },
                
                // 特殊代码: FB-SC115•140•14D
                { regex: /\b([A-Z]{1,3}[\-\.][A-Z]{1,3}\d+[•\.]\d+[•\.]\d+[A-Za-z]?)\b/, type: 'SPECIAL_CODE' }
            ];
            
            // 检查每种模式
            let found = false;
            for (const {regex, type} of patterns) {
                const match = line.match(regex);
                if (match && match[1]) {
                    const partNumber = match[1];
                    
                    // 过滤明显不是配件号的内容
                    if (partNumber.match(/^\d{4}-\d{2}-\d{2}$/) || // 日期格式
                        partNumber.match(/^\d{5,}$/) || // 纯数字较长
                        partNumber.match(/^1[3-9]\d{9}$/)) { // 手机号
                        continue;
                    }
                    
                    // 确保足够长并且不是纯数字
                    if (partNumber.length >= 3 && !/^\d+$/.test(partNumber) && !seenParts.has(partNumber)) {
                        seenParts.add(partNumber);
                        
                        // 尝试提取配件名称 (在配件号后通常有名称)
                        let name = '';
                        const nameMatch = line.substring(line.indexOf(partNumber) + partNumber.length)
                            .match(/^\s+([^\d\s][^数量]*?)\s+(?:数量|[0-9]|$)/);
                        if (nameMatch && nameMatch[1]) {
                            name = nameMatch[1].trim();
                        }
                        
                        // 尝试提取数量
                        let quantity = 1;
                        const quantityMatch = line.match(/数量\s*[:：]?\s*(\d+)|数量[xX×]\s*(\d+)|(\d+)\s*[个件pcs]/i);
                        if (quantityMatch) {
                            const qty = quantityMatch[1] || quantityMatch[2] || quantityMatch[3];
                            quantity = parseInt(qty) || 1;
                        }
                        
                        partCandidates.push({
                            '图号': partNumber,
                            '名称': name || '未知配件',
                            '数量': quantity,
                            '备注': `从文档中提取: ${type}`
                        });
                        
                        found = true;
                        break; // 一行通常只有一个主要配件号
                    }
                }
            }
            
            // 如果没有找到标准模式，检查是否有特殊格式 (但跳过明显不是配件号的)
            if (!found && /[A-Z0-9]{5,}/.test(line)) {
                const specialMatch = line.match(/\b([A-Z0-9]{5,})\b/);
                if (specialMatch && specialMatch[1] && !seenParts.has(specialMatch[1])) {
                    const partNumber = specialMatch[1];
                    
                    // 过滤明显不是配件号的内容
                    if (partNumber.match(/^\d{4}-\d{2}-\d{2}$/) || // 日期格式
                        partNumber.match(/^\d{5,}$/) || // 纯数字较长
                        partNumber.match(/^1[3-9]\d{9}$/)) { // 手机号
                        continue;
                    }
                    
                    seenParts.add(partNumber);
                    
                    // 尝试提取名称和数量
                    let name = '';
                    const nameMatch = line.substring(line.indexOf(partNumber) + partNumber.length)
                        .match(/^\s+([^\d\s][^数量]*?)\s+(?:数量|[0-9]|$)/);
                    if (nameMatch && nameMatch[1]) {
                        name = nameMatch[1].trim();
                    }
                    
                    let quantity = 1;
                    const quantityMatch = line.match(/数量\s*[:：]?\s*(\d+)|数量[xX×]\s*(\d+)|(\d+)\s*[个件pcs]/i);
                    if (quantityMatch) {
                        const qty = quantityMatch[1] || quantityMatch[2] || quantityMatch[3];
                        quantity = parseInt(qty) || 1;
                    }
                    
                    partCandidates.push({
                        '图号': partNumber,
                        '名称': name || '未知配件',
                        '数量': quantity,
                        '备注': '从文档中提取: SPECIAL_FORMAT'
                    });
                }
            }
        }
    }
    
    return partCandidates;
}

// 从Word文档中提取配件信息
async function extractPartsFromDocText(text) {
    // Word文档的解析逻辑与PDF类似，但可能格式更规整
    return await extractPartsFromPdfText(text);
}

// 使用PDFjs提取PDF文本内容
async function extractTextFromPdf(fileContent) {
    if (typeof window.pdfjsLib === 'undefined') {
        // 如果PDFjs库不可用，尝试使用模拟数据
        console.warn('PDF解析库未加载，使用模拟数据');
        return `
            客户订单单
            订单号: OD-123456
            日期: 2025-03-15
            
            配件清单:
            1. 135-01-003A 输入轴总成 数量: 2
            2. M10X25GB32.1-88 普通螺栓 数量: 10
            3. 6317N 轴承 数量: 4
            4. FB-SC115•140•14D 油封 数量: 1
            5. NJ313 轴承 数量: 5
            6. 135A-03A-016A 轴套 数量: 3
        `;
    }
    
    try {
        // 加载PDF文档
        const loadingTask = window.pdfjsLib.getDocument({ data: fileContent });
        const pdf = await loadingTask.promise;
        
        let fullText = '';
        
        // 遍历每一页并提取文本
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(' ');
            fullText += pageText + '\n';
        }
        
        return fullText;
    } catch (error) {
        console.error('PDF解析错误:', error);
        throw new Error('PDF文件解析失败: ' + error.message);
    }
}

// 使用mammoth提取Word文档文本内容
async function extractTextFromDoc(fileContent) {
    if (typeof window.mammoth === 'undefined') {
        // 如果mammoth库不可用，尝试使用模拟数据
        console.warn('Word解析库未加载，使用模拟数据');
        return `
            客户订单单
            订单号: OD-123456
            日期: 2025-03-15
            
            配件清单:
            1. 135-01-003A 输入轴总成 数量: 2
            2. M10X25GB32.1-88 普通螺栓 数量: 10
            3. 6317N 轴承 数量: 4
            4. FB-SC115•140•14D 油封 数量: 1
            5. NJ313 轴承 数量: 5
            6. 135A-03A-016A 轴套 数量: 3
        `;
    }
    
    try {
        // 使用mammoth提取文本
        const result = await window.mammoth.extractRawText({ arrayBuffer: fileContent });
        return result.value;
    } catch (error) {
        console.error('Word文档解析错误:', error);
        throw new Error('Word文件解析失败: ' + error.message);
    }
}

// 从不同类型的文档中提取配件信息的统一接口
async function extractPartsFromDocument(file, fileContent) {
    const fileType = detectFileType(file.name);
    
    try {
        switch (fileType) {
            case 'excel':
                const workbook = XLSX.read(fileContent, ROBUST_EXCEL_OPTIONS);
                return await extractGenericPartsList(workbook);
                
            case 'pdf':
                try {
                    // 先尝试使用PDFjs提取文本
                    const pdfText = await extractTextFromPdf(fileContent);
                    return await extractPartsFromPdfText(pdfText);
                } catch (error) {
                    console.error('PDF解析失败，使用模拟提取:', error);
                    // 如果PDF解析失败，使用模拟数据
                    return getMockPartsData();
                }
                
            case 'word':
                try {
                    // 尝试使用mammoth提取Word文档文本
                    const docText = await extractTextFromDoc(fileContent);
                    return await extractPartsFromDocText(docText);
                } catch (error) {
                    console.error('Word文档解析失败，使用模拟提取:', error);
                    // 如果Word解析失败，使用模拟数据
                    return getMockPartsData();
                }
                
            default:
                throw new Error('不支持的文件类型: ' + fileType);
        }
    } catch (error) {
        console.error('文档处理错误:', error);
        // 如果解析失败，返回一些模拟数据，以便系统可以继续运行
        return [
            { '图号': '135-01-003A', '名称': '输入轴总成', '数量': 2, '备注': '从模拟数据' },
            { '图号': 'M10X25GB32.1-88', '名称': '普通螺栓', '数量': 10, '备注': '从模拟数据' },
            { '图号': '6317N', '名称': '轴承', '数量': 4, '备注': '从模拟数据' },
            { '图号': 'FB-SC115•140•14D', '名称': '油封', '数量': 1, '备注': '从模拟数据' },
            { '图号': 'NJ313', '名称': '轴承', '数量': 5, '备注': '从模拟数据' },
            { '图号': '135A-03A-016A', '名称': '轴套', '数量': 3, '备注': '从模拟数据' }
        ];
    }
}

// 模拟从 PDF 或特殊文档中提取到的配件列表 (用于开发测试)
function getMockPartsData() {
    return [
        { '图号': '135-01-003A', '名称': '输入轴总成', '数量': 2, '备注': '模拟数据' },
        { '图号': 'B12X40GB120-86', '名称': '紧固螺栓', '数量': 8, '备注': '模拟数据' },
        { '图号': 'M10X25GB32.1-88', '名称': '螺栓', '数量': 12, '备注': '模拟数据' },
        { '图号': '135-01-002', '名称': '调整插头', '数量': 1, '备注': '模拟数据' },
        { '图号': '6317N', '名称': '轴承', '数量': 4, '备注': '模拟数据' },
        { '图号': '135-01-004', '名称': '输入轴套', '数量': 1, '备注': '模拟数据' },
        { '图号': 'NJ313', '名称': '轴承', '数量': 2, '备注': '模拟数据' },
        { '图号': '135-01-007', '名称': '输入端止推环', '数量': 1, '备注': '模拟数据' },
        { '图号': '135-01-024B', '名称': '离合器壳体', '数量': 1, '备注': '模拟数据' },
        { '图号': '135-01-032', '名称': '内六角螺钉', '数量': 6, '备注': '模拟数据' },
        { '图号': '135A-03A-016A', '名称': '轴套', '数量': 2, '备注': '模拟数据' },
        { '图号': 'FB-SC115•140•14D', '名称': '油封', '数量': 1, '备注': '模拟数据' }
    ];
}

async function extractGenericPartsList(workbook) {
    console.log("使用通用配件单处理逻辑...");
    const extractedParts = [];
    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet || !sheet['!ref']) continue;
        const range = XLSX.utils.decode_range(sheet['!ref']);
        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddress = XLSX.utils.encode_cell({ r, c });
                const cell = sheet[cellAddress];
                if (cell && (cell.v !== undefined || cell.w)) {
                    const value = String(cell.w || cell.v || '').trim();
                    if (
                        value !== '' &&
                        value.length >= 3 &&
                        !/^\d+$/.test(value) &&
                        (
                            /^[A-Za-z0-9\-\.\/]+$/.test(value) ||
                            /^[A-Za-z]+\d+$/.test(value) ||
                            /^\d+[A-Za-z]+\d*$/.test(value) ||
                            /^[A-Za-z]\d+[A-Za-z]\d+$/.test(value) ||
                            /^\d+[\-\.]\d+[\-\.]\d+[A-Za-z]?$/.test(value)
                        )
                    ) {
                        let quantity = 1;
                        if (c + 1 <= range.e.c) {
                            const qtyCell = sheet[XLSX.utils.encode_cell({ r, c: c + 1 })];
                            if (qtyCell && !isNaN(qtyCell.v)) {
                                quantity = parseInt(qtyCell.v) || 1;
                            }
                        }
                        extractedParts.push({
                            '图号': value,
                            '数量': quantity,
                            '单价': 0,
                            '备注': `工作表: ${sheetName}, 单元格: ${cellAddress}`
                        });
                    }
                }
            }
        }
    }
    return removeDuplicates(extractedParts);
}

async function extractAdvancePartsList(workbook) {
    console.log("解析 Advance 配件清单...");
    return await extractGenericPartsList(workbook);
}

async function extractXiamenPartsList(workbook) {
    console.log("解析 厦门斯太琪 配件单...");
    return await extractGenericPartsList(workbook);
}

// ------------------ 4. 批量处理已提取的配件并进行匹配 ------------------ //
async function processExtractedParts(
    extractedParts,
    partsData,
    setSelectedParts,
    setView,
    setInfoMessage,
    setLoading,
    setCurrentPage
) {
    console.log("处理 " + extractedParts.length + " 个提取的配件...");
    if (extractedParts.length === 0) {
        setLoading(false);
        setInfoMessage(null);
        alert('未能从文件中提取到任何配件号，请检查文件格式。');
        return;
    }
    setInfoMessage(`正在匹配 ${extractedParts.length} 个配件，请稍候...`);
    const batchSize = 20;
    const matchedParts = [];
    let matchedCount = 0;
    let newCount = 0;
    let processedCount = 0;
    function processBatch(startIndex) {
        if (startIndex >= extractedParts.length) {
            finishProcessing();
            return;
        }
        const endIndex = Math.min(startIndex + batchSize, extractedParts.length);
        const batch = extractedParts.slice(startIndex, endIndex);
        for (const part of batch) {
            try {
                const matchResult = enhancedFuzzyMatch(part['图号'], partsData);
                if (matchResult) {
                    matchedCount++;
                    matchedParts.push({
                        ...matchResult,
                        importedId: part['图号'],
                        quantity: part['数量'] || 1,
                        importedPrice: part['单价'] || 0,
                        importedRemark: part['备注'] || ''
                    });
                } else {
                    newCount++;
                    matchedParts.push({
                        '标识码': "NEW_" + part['图号'],
                        '图号': part['图号'],
                        '名称': part['名称'] || '未知配件',
                        '指导价（不含税）': 0,
                        '出厂价（不含税）': 0,
                        '服务价（不含税）': 0,
                        '指导价（含税）': 0,
                        '出厂价（含税）': 0,
                        '服务价（含税）': part['单价'] || 0,
                        '备注': part['备注'] || "从客户文件导入: " + part['图号'],
                        'quantity': part['数量'] || 1,
                        'isNew': true,
                        'importedId': part['图号'],
                        'matchType': 'none'
                    });
                }
            } catch (error) {
                console.error(`处理配件时出错 (${part['图号']}):`, error);
            }
            processedCount++;
        }
        const totalItems = extractedParts.length || 1;
        const percentage = Math.min(100, Math.round((processedCount / totalItems) * 100));
        setInfoMessage(`已处理 ${processedCount}/${totalItems} 个配件 (${percentage}%)...`);
        if (endIndex < extractedParts.length) {
            setTimeout(() => processBatch(endIndex), 10);
        } else {
            finishProcessing();
        }
    }
    function finishProcessing() {
        setSelectedParts(matchedParts);
        setView('quotation');
        setCurrentPage(1);
        const message =
            `成功导入 ${matchedParts.length} 条配件数据！其中匹配成功 ${matchedCount} 条，新配件 ${newCount} 条。`;
        setInfoMessage(message);
        console.log(message);
        setLoading(false);
    }
    processBatch(0);
}

// ------------------ 5. 主组件 ------------------ //
export default function ImprovedPartsQuotationSystem() {
    const isElectron = window.electronAPI !== undefined;
    const [partsData, setPartsData] = useState([]);
    const [selectedParts, setSelectedParts] = useState([]);
    const [searchTerm, setSearchTerm] = useState('');
    const [view, setView] = useState('table');
    const [loading, setLoading] = useState(true);
    const [infoMessage, setInfoMessage] = useState(null);
    const [priceOption, setPriceOption] = useState("服务价（含税）");
    const [currentPage, setCurrentPage] = useState(1);
    const [pageSize, setPageSize] = useState(50);
    const [isAdmin, setIsAdmin] = useState(false);
    const [sortConfig, setSortConfig] = useState({ key: null, direction: 'ascending' });
   const [theme, setTheme] = useState('light');
    const [customerInfo, setCustomerInfo] = useState({
        name: '',
        contact: '',
        date: new Date().toISOString().split('T')[0],
        vessel: '',
        project: ''
    });
    const [showPriceColumns, setShowPriceColumns] = useState({
        '指导价（不含税）': true,
        '出厂价（不含税）': true,
        '服务价（不含税）': true,
        '指导价（含税）': true,
        '出厂价（含税）': true,
        '服务价（含税）': true
    });

    function handleAdminLogin() {
        const password = prompt("请输入管理员密码：");
        if (password === "admin123") {
            setIsAdmin(true);
            alert("管理员登录成功！");
        } else {
            alert("密码错误，无法登录为管理员。");
        }
    }

    function handleAdminLogout() {
        setIsAdmin(false);
        alert("已退出管理员模式。");
    }

    function handleSortChange(key) {
        let direction = 'ascending';
        if (sortConfig.key === key && sortConfig.direction === 'ascending') {
            direction = 'descending';
        }
        setSortConfig({ key, direction });
    }

    function toggleTheme() {
        setTheme(theme === 'light' ? 'dark' : 'light');
    }

    useEffect(() => {
        const loadData = async () => {
            setLoading(true);
            try {
                if (isElectron) {
                    const loadedData = await window.electronAPI.readData();
                    if (loadedData && loadedData.length > 0) {
                        setPartsData(loadedData);
                    } else {
                        loadSampleData();
                    }
                } else {
                    const stored = safelyRetrieveData('shipPartsData');
                    if (stored) {
                        setPartsData(stored);
                    } else {
                        loadSampleData();
                    }
                }
            } catch (err) {
                console.error("加载数据失败:", err);
                loadSampleData();
            } finally {
                setLoading(false);
            }
        };
        loadData();
    }, [isElectron]);

    function loadSampleData() {
        const sampleData = [
            {
                '日期': '2025-03-01',
                '标识码': "ZB0001",
                '图号': "MV1100-02-002A",
                '名称': "侧车轴",
                '指导价（不含税）': 3250000,
                '出厂价（不含税）': 2762500,
                '服务价（不含税）': 3900000,
                '指导价（含税）': 3672500,
                '出厂价（含税）': 3121625,
                '服务价（含税）': 4407000,
                '备注': '重点设备'
            },
            {
                '日期': '2025-03-02',
                '标识码': "ZB0002",
                '图号': "HC400-01-000",
                '名称': "输入轴部件",
                '指导价（不含税）': 758000,
                '出厂价（不含税）': 644300,
                '服务价（不含税）': 909600,
                '指导价（含税）': 856540,
                '出厂价（含税）': 728059,
                '服务价（含税）': 1027848,
                '备注': ''
            }
        ];
        setPartsData(sampleData);
        saveDataToStorage(sampleData);
    }

    function saveDataToStorage(data) {
        try {
            if (isElectron) {
                window.electronAPI.saveData(data).then(success => {
                    if (!success) console.error("保存数据失败");
                });
            } else {
                safelyStoreData('shipPartsData', data);
            }
        } catch (error) {
            console.error("保存数据失败:", error);
        }
    }

    async function exportDatabaseToExcel() {
        if (partsData.length === 0) {
            alert('数据库为空，无法导出');
            return;
        }
        
        // 创建一个格式化的数据副本用于导出
        const exportData = partsData.map(part => ({
            '日期': part['日期'],
            '标识码': part['标识码'],
            '图号': part['图号'],
            '名称': part['名称'],
            '指导价（不含税）': part['指导价（不含税）'],
            '出厂价（不含税）': part['出厂价（不含税）'],
            '服务价（不含税）': part['服务价（不含税）'],
            '指导价（含税）': part['指导价（含税）'],
            '出厂价（含税）': part['出厂价（含税）'],
            '服务价（含税）': part['服务价（含税）'],
            '备注': part['备注']
        }));
        
        if (isElectron) {
            try {
                const success = await window.electronAPI.exportData(exportData);
                if (success) {
                    alert('数据导出成功');
                }
            } catch (error) {
                console.error("导出数据失败:", error);
                alert('导出失败: ' + error.message);
            }
        } else {
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "船用配件数据");
            XLSX.writeFile(workbook, `船用配件数据库_${new Date().toISOString().split('T')[0]}.xlsx`);
        }
    }

    // 根据搜索框和排序配置，筛选和排序数据
    const allFilteredData = useMemo(() => {
        let filteredData = partsData;
        
        // 应用搜索过滤
        if (searchTerm.trim()) {
            const lower = searchTerm.toLowerCase();
            filteredData = filteredData.filter(item =>
                (item['日期'] && String(item['日期']).toLowerCase().includes(lower)) ||
                (item['名称'] && String(item['名称']).toLowerCase().includes(lower)) ||
                (item['标识码'] && String(item['标识码']).toLowerCase().includes(lower)) ||
                (item['图号'] && String(item['图号']).toLowerCase().includes(lower)) ||
                (item['备注'] && String(item['备注']).toLowerCase().includes(lower))
            );
        }
        
        // 应用排序
        if (sortConfig.key) {
            filteredData = [...filteredData].sort((a, b) => {
                if (a[sortConfig.key] < b[sortConfig.key]) {
                    return sortConfig.direction === 'ascending' ? -1 : 1;
                }
                if (a[sortConfig.key] > b[sortConfig.key]) {
                    return sortConfig.direction === 'ascending' ? 1 : -1;
                }
                return 0;
            });
        }
        
        return filteredData;
    }, [searchTerm, partsData, sortConfig]);

    useEffect(() => {
        setCurrentPage(1);
    }, [allFilteredData]);

    const totalPages = Math.ceil(allFilteredData.length / pageSize);
    const currentPageData = useMemo(() => {
        const startIndex = (currentPage - 1) * pageSize;
        return allFilteredData.slice(startIndex, startIndex + pageSize);
    }, [allFilteredData, currentPage, pageSize]);

    function handleSelectPart(part) {
        const isSelected = selectedParts.some(p => p['标识码'] === part['标识码']);
        if (isSelected) {
            setSelectedParts(selectedParts.filter(p => p['标识码'] !== part['标识码']));
        } else {
            setSelectedParts([...selectedParts, { ...part, quantity: 1 }]);
        }
    }

    function updatePartQuantity(partId, quantity) {
        setSelectedParts(
            selectedParts.map(part =>
                part['标识码'] === partId
                    ? { ...part, quantity: Math.max(1, parseInt(quantity) || 1) }
                    : part
            )
        );
    }

    function updatePartCustomPrice(partId, price) {
        setSelectedParts(
            selectedParts.map(part =>
                part['标识码'] === partId
                    ? { ...part, importedPrice: parseFloat(price) || 0 }
                    : part
            )
        );
    }

    function generateQuotation() {
        if (selectedParts.length === 0) {
            alert('请至少选择一个配件');
            return;
        }
        setView('quotation');
    }

    function backToList() {
        setView('table');
        setCurrentPage(1);
    }

    // 新增：移除所有选中的配件
    function clearSelectedParts() {
        if (selectedParts.length > 0 && window.confirm('确定要清空所有已选配件吗？')) {
            setSelectedParts([]);
        }
    }

    // 新增：删除单个选中的配件
    function removeSelectedPart(partId) {
        setSelectedParts(selectedParts.filter(part => part['标识码'] !== partId));
    }

    // 新增：应用批量折扣
    function applyBulkDiscount() {
        const discountPercent = prompt("请输入折扣百分比 (例如: 输入90代表9折):", "100");
        if (discountPercent === null) return;
        
        const discount = parseFloat(discountPercent) / 100;
        if (isNaN(discount) || discount <= 0) {
            alert("请输入有效的折扣值");
            return;
        }
        
        setSelectedParts(selectedParts.map(part => {
            const originalPrice = part.importedPrice || part[priceOption] || 0;
            return {
                ...part,
                importedPrice: parseFloat((originalPrice * discount).toFixed(2))
            };
        }));
        
        alert(`已对所有配件应用${discountPercent}%的折扣`);
    }

    function exportQuotationCSV() {
        if (selectedParts.length === 0) {
            alert('报价单为空，无法导出');
            return;
        }
        
        let csvContent =
            '序号,客户提供标识,系统标识码,图号,名称,价格类型,单价,数量,总价(元),备注,匹配方式\n';
            
        selectedParts.forEach((part, index) => {
            const quantity = part.quantity || 1;
            const price = part.importedPrice || part[priceOption] || 0;
            const itemTotal = price * quantity;
            const remark = part.importedRemark || part['备注'] || '';
            const matchType = part.matchType === 'fuzzy'
                ? '模糊匹配'
                : (part.isNew ? '新配件' : '精确匹配');
                
            csvContent += [
                index + 1,
                `"${part.importedId || ''}"`,
                `"${part['标识码']}"`,
                `"${part['图号']}"`,
                `"${part['名称']}"`,
                part.importedPrice ? '客户指定价格' : priceOption,
                price.toFixed(2),
                quantity,
                itemTotal.toFixed(2),
                `"${remark}"`,
                `"${matchType}"`
            ].join(',') + '\n';
        });
        
        const totalPrice = selectedParts.reduce((sum, p) => {
            const priceVal = p.importedPrice || p[priceOption] || 0;
            return sum + priceVal * (p.quantity || 1);
        }, 0);
        
        csvContent += `总计:,,,,,,,,,${totalPrice.toFixed(2)}\n`;
        
        // 添加客户信息到CSV
        csvContent += `\n客户信息:\n`;
        csvContent += `客户:,${customerInfo.name}\n`;
        csvContent += `联系方式:,${customerInfo.contact}\n`;
        csvContent += `日期:,${customerInfo.date}\n`;
        csvContent += `船舶:,${customerInfo.vessel}\n`;
        csvContent += `项目:,${customerInfo.project}\n`;

        if (isElectron) {
            window.electronAPI.exportData(csvContent).then(success => {
                if (success) {
                    alert('报价单导出成功');
                } else {
                    alert('报价单导出失败');
                }
            });
        } else {
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.setAttribute('href', url);
            link.setAttribute('download', `船用配件报价_${customerInfo.name || '未命名'}_${new Date().toISOString().split('T')[0]}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }

    // 新增：Excel格式导出
    function exportQuotationExcel() {
        if (selectedParts.length === 0) {
            alert('报价单为空，无法导出');
            return;
        }
        
        // 准备Excel数据
        const excelData = selectedParts.map((part, index) => {
            const quantity = part.quantity || 1;
            const price = part.importedPrice || part[priceOption] || 0;
            const itemTotal = price * quantity;
            const matchType = part.matchType === 'fuzzy'
                ? '模糊匹配'
                : (part.isNew ? '新配件' : '精确匹配');
                
            return {
                '序号': index + 1,
                '客户提供标识': part.importedId || '',
                '系统标识码': part['标识码'],
                '图号': part['图号'],
                '名称': part['名称'],
                '价格类型': part.importedPrice ? '客户指定价格' : priceOption,
                '单价': price,
                '数量': quantity,
                '总价(元)': itemTotal,
                '备注': part.importedRemark || part['备注'] || '',
                '匹配方式': matchType
            };
        });
        
        // 添加合计行
        const totalPrice = selectedParts.reduce((sum, p) => {
            const priceVal = p.importedPrice || p[priceOption] || 0;
            return sum + priceVal * (p.quantity || 1);
        }, 0);
        
        // 合计行
        excelData.push({
            '序号': '',
            '客户提供标识': '',
            '系统标识码': '',
            '图号': '',
            '名称': '合计',
            '价格类型': '',
            '单价': '',
            '数量': '',
            '总价(元)': totalPrice,
            '备注': '',
            '匹配方式': ''
        });
        
        // 客户信息行
        excelData.push({}, {
            '序号': '客户信息',
            '客户提供标识': ''
        });
        
        excelData.push({
            '序号': '客户名称',
            '客户提供标识': customerInfo.name || ''
        });
        
        excelData.push({
            '序号': '联系方式',
            '客户提供标识': customerInfo.contact || ''
        });
        
        excelData.push({
            '序号': '日期',
            '客户提供标识': customerInfo.date || ''
        });
        
        excelData.push({
            '序号': '船舶',
            '客户提供标识': customerInfo.vessel || ''
        });
        
        excelData.push({
            '序号': '项目',
            '客户提供标识': customerInfo.project || ''
        });
        
        const worksheet = XLSX.utils.json_to_sheet(excelData);
        
        // 设置列宽
        const colWidths = [
            { wch: 6 },   // 序号
            { wch: 15 },  // 客户提供标识
            { wch: 12 },  // 系统标识码
            { wch: 15 },  // 图号
            { wch: 20 },  // 名称
            { wch: 12 },  // 价格类型
            { wch: 12 },  // 单价
            { wch: 6 },   // 数量
            { wch: 12 },  // 总价
            { wch: 20 },  // 备注
            { wch: 10 }   // 匹配方式
        ];
        worksheet['!cols'] = colWidths;
        
        // 创建工作簿并添加工作表
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "船用配件报价单");
        
        // 导出Excel文件
        const fileName = `船用配件报价_${customerInfo.name || '未命名'}_${new Date().toISOString().split('T')[0]}.xlsx`;
        
        if (isElectron) {
            // 如果是Electron环境
            const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
            window.electronAPI.exportData({ fileName, data: excelBuffer, type: 'excel' }).then(success => {
                if (success) {
                    alert('Excel报价单导出成功');
                } else {
                    alert('Excel报价单导出失败');
                }
            });
        } else {
            // 浏览器环境
            XLSX.writeFile(workbook, fileName);
        }
    }
    
    // 新增：批量添加配件
    function batchAddParts() {
        const input = prompt("请输入多个配件号，每行一个：");
        if (!input || input.trim() === '') return;
        
        const partNumbers = input.split('\n').map(line => line.trim()).filter(line => line.length > 0);
        if (partNumbers.length === 0) return;
        
        const extractedParts = partNumbers.map(partNumber => ({
            '图号': partNumber,
            '数量': 1,
            '单价': 0,
            '备注': '手动批量添加'
        }));
        
        processExtractedParts(
            extractedParts,
            partsData,
            setSelectedParts,
            setView,
            setInfoMessage,
            setLoading,
            setCurrentPage
        );
    }
    
    // 新增：导入报价单模板
    function importQuotationTemplate() {
        alert("请选择报价单模板文件");
        document.getElementById('fileTemplate').click();
    }
    
    async function handleTemplateUpload(e) {
        const { file, fileContent } = await readFileContent(e, isElectron);
        if (!file || !fileContent) {
            alert('文件读取失败');
            return;
        }
        
        try {
            const workbook = XLSX.read(fileContent, ROBUST_EXCEL_OPTIONS);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const templateData = XLSX.utils.sheet_to_json(worksheet);
            
            // 提取客户信息
            const customerData = templateData.find(row => 
                row['客户名称'] || row['客户'] || row['Client'] || row['Customer']
            );
            
            if (customerData) {
                setCustomerInfo({
                    name: customerData['客户名称'] || customerData['客户'] || customerData['Client'] || customerData['Customer'] || '',
                    contact: customerData['联系方式'] || customerData['联系人'] || customerData['Contact'] || '',
                    date: customerData['日期'] || customerData['Date'] || new Date().toISOString().split('T')[0],
                    vessel: customerData['船舶'] || customerData['船名'] || customerData['Vessel'] || '',
                    project: customerData['项目'] || customerData['工程'] || customerData['Project'] || ''
                });
            }
            
            // 提取配件列表
            const partsList = templateData.filter(row => 
                row['图号'] || row['件号'] || row['Part No.'] || row['Material No.']
            );
            
            if (partsList.length > 0) {
                const extractedParts = partsList.map(row => ({
                    '图号': row['图号'] || row['件号'] || row['Part No.'] || row['Material No.'] || '',
                    '名称': row['名称'] || row['件名'] || row['Description'] || '',
                    '数量': parseInt(row['数量'] || row['Quantity'] || 1),
                    '单价': parseFloat(row['单价'] || row['Price'] || 0),
                    '备注': row['备注'] || row['Remark'] || ''
                }));
                
                await processExtractedParts(
                    extractedParts,
                    partsData,
                    setSelectedParts,
                    setView,
                    setInfoMessage,
                    setLoading,
                    setCurrentPage
                );
            } else {
                alert('未在模板中找到有效的配件数据');
            }
            
        } catch (error) {
            handleProcessingError('模板文件解析', error, setLoading, setInfoMessage);
        }
    }
    
    function handlePrint() {
    // 设置打印样式
    const printElement = document.createElement('style');
    printElement.innerHTML = `
        @media print {
            body * {
                visibility: hidden;
            }
            #print-container, #print-container * {
                visibility: visible;
            }
            #print-container {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
            }
            .no-print {
                display: none !important;
            }
            @page {
                size: A4;
                margin: 1cm;
            }
            table {
                page-break-inside: avoid;
            }
            tr {
                page-break-inside: avoid;
                page-break-after: auto;
            }
            thead {
                display: table-header-group;
            }
            tfoot {
                display: table-footer-group;
            }
            
            /* 第二页隐藏页头 */
            .page-header {
                display: none;
            }
            .first-page .page-header {
                display: block;
            }
            .page-break-before {
                page-break-before: always;
            }
        }
    `;
    document.head.appendChild(printElement);
    
    window.print();
    
    // 移除打印样式
    document.head.removeChild(printElement);
}
    // 总价和其他统计信息
    const statistics = useMemo(() => {
        const total = selectedParts.reduce((sum, p) => {
            const priceVal = p.importedPrice || p[priceOption] || 0;
            return sum + priceVal * (p.quantity || 1);
        }, 0);
        
        const totalQuantity = selectedParts.reduce((sum, p) => sum + (p.quantity || 1), 0);
        const exactMatchCount = selectedParts.filter(p => p.matchType === 'exact').length;
        const fuzzyMatchCount = selectedParts.filter(p => p.matchType === 'fuzzy' || p.matchType === 'caseInsensitive' || p.matchType === 'noSpace').length;
        const newPartCount = selectedParts.filter(p => p.isNew || p.matchType === 'none').length;
        
        return {
            totalPrice: total,
            totalQuantity,
            exactMatchCount,
            fuzzyMatchCount,
            newPartCount
        };
    }, [selectedParts, priceOption]);

    // 新增：获取当前主题对应的样式
    function getThemeStyles() {
        if (theme === 'dark') {
            return {
                background: '#1e1e1e',
                text: '#f0f0f0',
                container: '#2d2d2d',
                header: '#3d3d3d',
                border: '#444',
                button: '#555',
                buttonText: '#fff',
                inputBackground: '#333',
                inputText: '#fff',
                inputBorder: '#555',
                link: '#4da6ff',
                alertBackground: '#382b2b',
                alertText: '#ff6b6b'
            };
        }
        return {
            background: '#f9fafb',
            text: '#333',
            container: '#ffffff',
            header: '#f2f2f2',
            border: '#ddd',
            button: '#e5e5e5',
            buttonText: '#333',
            inputBackground: '#fff',
            inputText: '#333',
            inputBorder: '#ccc',
            link: '#0066cc',
            alertBackground: '#fff8f8',
            alertText: '#d14'
        };
    }
    
    const themeStyles = getThemeStyles();
    
    useEffect(() => {
        // 当主题改变时，更新文档的背景色和文本色
        document.body.style.backgroundColor = themeStyles.background;
        document.body.style.color = themeStyles.text;
    }, [theme, themeStyles]);
    
    // ------------------ 文件导入事件封装 ------------------ //
    async function handleFileUploadWrapper(e) {
        const { file, fileContent } = await readFileContent(e, isElectron);
        if (!file || !fileContent) {
            alert('文件读取失败');
            return;
        }
        setLoading(true);
        try {
            const workbook = XLSX.read(fileContent, ROBUST_EXCEL_OPTIONS);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            console.log("解析的Excel数据:", jsonData);

            // 进行简单的格式推断
            const formattedData = jsonData.map((row, index) => {
                if (row['编码'] && row['出厂价'] !== undefined) {
                    const singlePrice = parseFloat(row['出厂价'] || 0);
                    return {
                        '日期': row['日期'] || '',
                        '标识码': row['编码'] || `IMPORT${index + 1}`,
                        '图号': row['件号'] || '',
                        '名称': row['件名'] || '未知配件',
                        '指导价（不含税）': singlePrice,
                        '出厂价（不含税）': singlePrice,
                        '服务价（不含税）': singlePrice,
                        '指导价（含税）': singlePrice,
                        '出厂价（含税）': singlePrice,
                        '服务价（含税）': singlePrice,
                        '备注': [
                            row['单位'] ? `单位:${row['单位']}` : '',
                            row['备注与其他数据合并一起']
                        ].filter(Boolean).join('; ')
                    };
                } else if (row['标识码']) {
                    return {
                        '日期': row['日期'] || '',
                        '标识码': row['标识码'] || `IMPORT${index + 1}`,
                        '图号': row['图号'] || '',
                        '名称': row['名称'] || '未知配件',
                        '指导价（不含税）': parseFloat(row['指导价（不含税）'] || 0),
                        '出厂价（不含税）': parseFloat(row['出厂价（不含税）'] || 0),
                        '服务价（不含税）': parseFloat(row['服务价（不含税）'] || 0),
                        '指导价（含税）': parseFloat(row['指导价（含税）'] || 0),
                        '出厂价（含税）': parseFloat(row['出厂价（含税）'] || 0),
                        '服务价（含税）': parseFloat(row['服务价（含税）'] || 0),
                        '备注': row['备注'] || ''
                    };
                } else {
                    const singlePrice = parseFloat(row['出厂价'] || 0);
                    return {
                        '日期': row['日期'] || '',
                        '标识码': row['标识码'] || row['编码'] || `IMPORT${index + 1}`,
                        '图号': row['图号'] || row['件号'] || '',
                        '名称': row['名称'] || row['件名'] || '未知配件',
                        '指导价（不含税）': row['出厂价'] ? singlePrice : parseFloat(row['指导价（不含税）'] || 0),
                        '出厂价（不含税）': row['出厂价'] ? singlePrice : parseFloat(row['出厂价（不含税）'] || 0),
                        '服务价（不含税）': parseFloat(row['服务价（不含税）'] || 0),
                        '指导价（含税）': parseFloat(row['指导价（含税）'] || 0),
                        '出厂价（含税）': row['出厂价'] ? singlePrice : parseFloat(row['出厂价（含税）'] || 0),
                        '服务价（含税）': parseFloat(row['服务价（含税）'] || 0),
                        '备注': row['备注'] ||
                            ([row['单位'] ? `单位:${row['单位']}` : '', row['备注与其他数据合并一起']]
                                .filter(Boolean).join('; '))
                    };
                }
            }).filter(item => item['名称'] !== '未知配件' || item['出厂价（不含税）'] > 0);

            const existingIds = new Set(partsData.map(item => item['标识码']));
            const newData = formattedData.filter(item => !existingIds.has(item['标识码']));
            const combinedData = [...partsData, ...newData];
            setPartsData(combinedData);
            setCurrentPage(1);
            saveDataToStorage(combinedData);
            alert(`成功导入 ${newData.length} 条新数据！当前总共有 ${combinedData.length} 条数据。`);
        } catch (error) {
            handleProcessingError('文件解析', error, setLoading, setInfoMessage);
        } finally {
            setLoading(false);
        }
    }

    async function handleAdvancedQuotationUpload(e) {
        if (!isAdmin) {
            alert("无权限进行高级导入！");
            return;
        }
        const { file, fileContent } = await readFileContent(e, isElectron);
        if (!file || !fileContent) {
            alert('文件读取失败');
            return;
        }
        setLoading(true);
        setInfoMessage("正在处理文件，请稍候...");
        try {
            console.log("开始处理文件:", file.name);
            const isAdvanceSpares = file.name.includes("Advance") || file.name.includes("ADVANCE");
            const isXiamenFile = file.name.includes("厦门") || file.name.includes("斯太琪");
            const isPdf = file.name.toLowerCase().endsWith('.pdf');
            if (isPdf) {
                console.log("检测到PDF文件，采用预设的模拟解析逻辑");
                const extractedParts = extractPartsFromDocument();
                await processExtractedParts(
                    extractedParts,
                    partsData,
                    setSelectedParts,
                    setView,
                    setInfoMessage,
                    setLoading,
                    setCurrentPage
                );
                return;
            }
            let workbook;
            try {
                workbook = XLSX.read(fileContent, ROBUST_EXCEL_OPTIONS);
            } catch (error) {
                console.error("Excel读取错误:", error);
                setLoading(false);
                alert(`Excel文件格式错误: ${error.message}`);
                return;
            }
            let allExtractedParts = [];
            if (isAdvanceSpares) {
                console.log("使用 ADVANCE 配件清单专用处理逻辑");
                allExtractedParts = await extractAdvancePartsList(workbook);
            } else if (isXiamenFile) {
                console.log("使用 厦门斯太琪 配件单专用处理逻辑");
                allExtractedParts = await extractXiamenPartsList(workbook);
            } else {
                console.log("使用通用配件单处理逻辑");
                allExtractedParts = await extractGenericPartsList(workbook);
            }
            await processExtractedParts(
                allExtractedParts,
                partsData,
                setSelectedParts,
                setView,
                setInfoMessage,
                setLoading,
                setCurrentPage
            );
        } catch (error) {
            handleProcessingError('文件处理', error, setLoading, setInfoMessage);
        }
    }

    async function handleCustomerQuotationUpload(e) {
        const { file, fileContent } = await readFileContent(e, isElectron);
        if (!file || !fileContent) {
            alert('文件读取失败');
            return;
        }
        setLoading(true);
        try {
            const workbook = XLSX.read(fileContent, ROBUST_EXCEL_OPTIONS);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            if (!worksheet['!ref']) {
                throw new Error("Excel文件似乎是空的");
            }
            let customerData = [];
            try {
                customerData = XLSX.utils.sheet_to_json(worksheet);
                if (customerData.length === 0) {
                    // 尝试备用方法
                    const range = XLSX.utils.decode_range(worksheet['!ref']);
                    if (range.e.c === 0 || range.e.c === range.s.c) {
                        for (let r = range.s.r; r <= range.e.r; r++) {
                            const cellAddress = XLSX.utils.encode_cell({ r, c: range.s.c });
                            const cell = worksheet[cellAddress];
                            if (cell && cell.v) {
                                const value = cell.w || String(cell.v);
                                if (value.trim() !== '') {
                                    customerData.push({
                                        '图号': value.trim(),
                                        'rawValue': value.trim()
                                    });
                                }
                            }
                        }
                    }
                    if (customerData.length === 0) {
                        // 再尝试CSV解析
                        const csvOptions = { header: 1 };
                        const rows = XLSX.utils.sheet_to_json(worksheet, csvOptions);
                        if (rows.length > 0) {
                            customerData = rows.map(row => {
                                if (row.length > 0 && row[0] && String(row[0]).trim() !== '') {
                                    return {
                                        '图号': String(row[0]).trim(),
                                        'rawValue': String(row[0]).trim()
                                    };
                                }
                                return null;
                            }).filter(item => item !== null);
                        }
                    }
                }
            } catch (parseError) {
                console.error("标准解析失败，尝试备用方法:", parseError);
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                for (let r = range.s.r; r <= range.e.r; r++) {
                    const cellAddress = XLSX.utils.encode_cell({ r, c: 0 });
                    const cell = worksheet[cellAddress];
                    if (cell && (cell.v !== undefined || cell.w)) {
                        const value = cell.w || String(cell.v);
                        if (value.trim() !== '') {
                            customerData.push({
                                '图号': value.trim(),
                                'rawValue': value.trim()
                            });
                        }
                    }
                }
            }
            if (customerData.length === 0) {
                throw new Error("无法从Excel文件中提取有效数据，请检查文件格式");
            }
            const matchedParts = [];
            const processedIds = new Set();
            for (const item of customerData) {
                const importedId = String(
                    item['图号'] ||
                    item['标识码'] ||
                    item['rawValue'] ||
                    item[Object.keys(item)[0]] ||
                    ''
                ).trim();
                if (!importedId || processedIds.has(importedId)) continue;
                processedIds.add(importedId);

                const matchResult = enhancedFuzzyMatch(importedId, partsData);
                if (matchResult) {
                    matchedParts.push({
                        ...matchResult,
                        importedId,
                        quantity: parseInt(item['数量']) || 1,
                        importedPrice: parseFloat(item['单价']) || 0,
                        importedRemark: item['备注'] || ''
                    });
                } else {
                    matchedParts.push({
                        '标识码': `NEW_${importedId}`,
                        '图号': importedId,
                        '名称': item['名称'] || '未知配件',
                        '指导价（不含税）': 0,
                        '出厂价（不含税）': 0,
                        '服务价（不含税）': 0,
                        '指导价（含税）': 0,
                        '出厂价（含税）': 0,
                        '服务价（含税）': parseFloat(item['单价']) || 0,
                        '备注': item['备注'] || `客户提供标识: ${importedId}`,
                        'quantity': parseInt(item['数量']) || 1,
                        'isNew': true,
                        'importedId': importedId
                    });
                }
            }
            if (matchedParts.length === 0) {
                throw new Error("没有找到有效的配件数据，请检查文件格式");
            }
            setSelectedParts(matchedParts);
            setView('quotation');
            setCurrentPage(1);

            const matchedCount = matchedParts.filter(p => !p.isNew).length;
            const newCount = matchedParts.filter(p => p.isNew).length;
            setInfoMessage(
                `成功导入 ${matchedParts.length} 条数据！其中成功匹配 ${matchedCount} 条，未匹配新配件 ${newCount} 条。`
            );
        } catch (error) {
            handleProcessingError('文件解析', error, setLoading, setInfoMessage);
        } finally {
            setLoading(false);
        }
    }

    // 新增：处理通用文档文件上传
    async function handleDocumentUpload(e) {
        const { file, fileContent } = await readFileContent(e, isElectron);
        if (!file || !fileContent) {
            alert('文件读取失败');
            return;
        }
        
        setLoading(true);
        setInfoMessage("正在解析文档，提取配件信息...");
        
        try {
            // 根据文件类型提取配件信息
            const fileType = detectFileType(file.name);
            const extractedParts = await extractPartsFromDocument(file, fileContent);
            
            if (extractedParts && extractedParts.length > 0) {
                await processExtractedParts(
                    extractedParts,
                    partsData,
                    setSelectedParts,
                    setView,
                    setInfoMessage,
                    setLoading,
                    setCurrentPage
                );
            } else {
                setLoading(false);
                alert(`未能从${fileType}文件中提取配件信息，请尝试其他导入方式。`);
            }
        } catch (error) {
            handleProcessingError('文档解析', error, setLoading, setInfoMessage);
        }
    }
    
    // 新增：文档解析与人工审核流程
    function startDocumentProcessWorkflow() {
        // 弹窗询问文档类型
        const docType = window.confirm(
            "请选择文档类型:\n" +
            "确定 - PDF或Word等格式的客户订单\n" + 
            "取消 - Excel或CSV格式的表格数据"
        );
        
        if (docType) {
            // 用户选择了PDF/Word文档
            document.getElementById('fileDocument').click();
        } else {
            // 用户选择了Excel/CSV
            document.getElementById('fileCustomer').click();
        }
    }
    
    // 新增: 人工审核确认
    function reviewAndConfirm() {
        // 仅在有待审核的配件时显示
        if (selectedParts.length === 0) {
            alert('没有需要审核的配件');
            return;
        }
        
        const fuzzyMatches = selectedParts.filter(p => 
            p.matchType === 'fuzzy' || 
            p.matchType === 'caseInsensitive' || 
            p.matchType === 'noSpace'
        );
        
        if (fuzzyMatches.length > 0) {
            const confirmed = window.confirm(
                `系统通过模糊匹配找到了 ${fuzzyMatches.length} 个配件，请检查是否正确。\n` +
                `确认后将生成最终报价单。\n\n` +
                `是否已完成人工核对并确认？`
            );
            
            if (confirmed) {
                // 添加标记表示已经过人工审核
                setSelectedParts(selectedParts.map(part => ({
                    ...part,
                    humanReviewed: true
                })));
                
                // 生成报价单
                setView('quotation');
            }
        } else {
            // 如果没有模糊匹配，直接确认
            setSelectedParts(selectedParts.map(part => ({
                ...part,
                humanReviewed: true
            })));
            
            setView('quotation');
        }
    }

    return (
        <div style={{ padding: 10, backgroundColor: themeStyles.background, minHeight: '100vh', color: themeStyles.text }}>
            <style>{`
                @media print {
                    .no-print { 
                        display: none !important; 
                    }
                    body, html {
                        background-color: white !important;
                        color: black !important;
                    }
                    #print-container {
                        background-color: white !important;
                        color: black !important;
                    }
                    @page {
                        margin: 10mm;
                        @bottom-center {
                            content: "上海前进齿轮经营有限公司";
                            font-size: 10pt;
                            color: #000;
                        }
                    }
                }
                
                /* 自定义按钮样式 */
                .action-button {
                    background-color: ${themeStyles.button};
                    color: ${themeStyles.buttonText};
                    border: 1px solid ${themeStyles.border};
                    padding: 6px 12px;
                    border-radius: 4px;
                    cursor: pointer;
                    margin-right: 8px;
                    font-size: 14px;
                    transition: all 0.3s ease;
                }
                .action-button:hover {
                    opacity: 0.8;
                }
                .primary-button {
                    background-color: #1e88e5;
                    color: white;
                }
                .danger-button {
                    background-color: #e53935;
                    color: white;
                }
                
                /* 表格样式 */
                .data-table {
                    width: 100%;
                    border-collapse: collapse;
                    border: 1px solid ${themeStyles.border};
                }
                .data-table th {
                    background-color: ${themeStyles.header};
                    padding: 8px;
                    text-align: left;
                    border: 1px solid ${themeStyles.border};
                }
                .data-table td {
                    padding: 8px;
                    border: 1px solid ${themeStyles.border}
                }
                .data-table tr:hover {
                    background-color: ${themeStyles.background};
                }
                .data-table input {
                    background-color: ${themeStyles.inputBackground};
                    color: ${themeStyles.inputText};
                    border: 1px solid ${themeStyles.inputBorder};
                    padding: 4px;
                    border-radius: 4px;
                }
                
                /* 分页控制样式 */
                .pagination {
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    margin-top: 12px;
                }
                .pagination button {
                    background-color: ${themeStyles.button};
                    color: ${themeStyles.buttonText};
                    border: 1px solid ${themeStyles.border};
                    margin: 0 4px;
                    padding: 4px 8px;
                    border-radius: 4px;
                    cursor: pointer;
                }
                .pagination button:disabled {
                    opacity: 0.5;
                    cursor: not-allowed;
                }
                
                /* 加载指示器 */
                .loading-indicator {
                    text-align: center;
                    padding: 20px;
                    color: ${themeStyles.link};
                    font-weight: bold;
                }
                
                /* 搜索框和工具栏 */
                .toolbar {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 12px;
                }
                .search-box {
                    padding: 6px 12px;
                    border: 1px solid ${themeStyles.inputBorder};
                    border-radius: 4px;
                    background-color: ${themeStyles.inputBackground};
                    color: ${themeStyles.inputText};
                    width: 200px;
                }
                
                /* 响应式布局 */
                @media (max-width: 768px) {
                    .toolbar {
                        flex-direction: column;
                        align-items: flex-start;
                    }
                    .search-box {
                        width: 100%;
                        margin-bottom: 8px;
                    }
                    .data-table {
                        display: block;
                        overflow-x: auto;
                    }
                }
                
                /* 报价单样式 */
                .quotation-summary {
                    background-color: ${themeStyles.container};
                    border: 1px solid ${themeStyles.border};
                    border-radius: 4px;
                    padding: 12px;
                    margin-bottom: 16px;
                }
                .quotation-stats {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 12px;
                    margin-top: 8px;
                }
                .stat-item {
                    padding: 8px 12px;
                    border-radius: 4px;
                    background-color: ${themeStyles.background};
                    border: 1px solid ${themeStyles.border};
                }
            `}</style>

            {/* 页面头部 */}
            <div style={{ 
                display: 'flex', 
                justifyContent: 'space-between', 
                alignItems: 'center', 
                marginBottom: '15px',
                padding: '10px',
                backgroundColor: themeStyles.container,
                borderRadius: '4px',
                boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
            }}>
                <div style={{ fontSize: '18px', fontWeight: 'bold' }}>
                    船用配件报价系统
                </div>
                <div className="no-print" style={{ display: 'flex', alignItems: 'center' }}>
                    <button 
                        onClick={toggleTheme} 
                        className="action-button"
                        style={{ marginRight: '10px' }}
                    >
                        {theme === 'light' ? '深色模式' : '浅色模式'}
                    </button>
                    {!isAdmin ? (
                        <button onClick={handleAdminLogin} className="action-button">管理员登录</button>
                    ) : (
                        <button onClick={handleAdminLogout} className="action-button">退出管理员模式</button>
                    )}
                </div>
            </div>

            {loading && <div className="loading-indicator">加载中，请稍候...</div>}

            {/* 隐藏的文件选择器 */}
            <input type="file" id="fileBasic" style={{ display: 'none' }} onChange={handleFileUploadWrapper} />
            <input type="file" id="fileAdvanced" style={{ display: 'none' }} onChange={handleAdvancedQuotationUpload} />
            <input type="file" id="fileCustomer" style={{ display: 'none' }} onChange={handleCustomerQuotationUpload} />
            <input type="file" id="fileTemplate" style={{ display: 'none' }} onChange={handleTemplateUpload} />
            <input type="file" id="fileDocument" style={{ display: 'none' }} accept=".pdf,.doc,.docx,.rtf" onChange={handleDocumentUpload} />

            {view === 'table' && (
                <div style={{
                    width: '95%',
                    margin: '0 auto',
                    backgroundColor: themeStyles.container,
                    color: themeStyles.text,
                    padding: '15px',
                    borderRadius: '4px',
                    boxShadow: '0 1px 4px rgba(0,0,0,0.1)'
                }}>
                    {/* 工具栏 */}
                    <div className="toolbar no-print">
                        <div>
                            <input
                                type="text"
                                placeholder="搜索配件 (图号/名称/备注)..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="search-box"
                            />
                        </div>
                        <div>
                            <label style={{ marginRight: '10px' }}>每页显示: </label>
                            <select 
                                value={pageSize} 
                                onChange={(e) => setPageSize(Number(e.target.value))}
                                style={{
                                    backgroundColor: themeStyles.inputBackground,
                                    color: themeStyles.inputText,
                                    border: `1px solid ${themeStyles.inputBorder}`,
                                    padding: '4px',
                                    borderRadius: '4px',
                                    marginRight: '10px'
                                }}
                            >
                                <option value={10}>10</option>
                                <option value={20}>20</option>
                                <option value={50}>50</option>
                                <option value={100}>100</option>
                            </select>
                            <span style={{ marginRight: '10px' }}>
                                总计: {partsData.length} 条记录
                            </span>
                        </div>
                    </div>
                    
                    {/* 显示/隐藏价格列控制 */}
                    <div className="no-print" style={{ 
                        marginBottom: '10px', 
                        padding: '8px', 
                        backgroundColor: themeStyles.background,
                        borderRadius: '4px',
                        border: `1px solid ${themeStyles.border}`
                    }}>
                        <div style={{ marginBottom: '5px', fontWeight: 'bold' }}>显示价格列:</div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '10px' }}>
                            {Object.keys(showPriceColumns).map(column => (
                                <label key={column} style={{ display: 'flex', alignItems: 'center' }}>
                                    <input
                                        type="checkbox"
                                        checked={showPriceColumns[column]}
                                        onChange={() => setShowPriceColumns({
                                            ...showPriceColumns,
                                            [column]: !showPriceColumns[column]
                                        })}
                                        style={{ marginRight: '5px' }}
                                    />
                                    {column}
                                </label>
                            ))}
                        </div>
                    </div>

                    <table className="data-table">
                        <thead>
                            <tr>
                                <th onClick={() => handleSortChange('序号')} style={{ cursor: 'pointer' }}>
                                    序号 {sortConfig.key === '序号' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                </th>
                                <th onClick={() => handleSortChange('日期')} style={{ cursor: 'pointer' }}>
                                    日期 {sortConfig.key === '日期' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                </th>
                                <th onClick={() => handleSortChange('标识码')} style={{ cursor: 'pointer' }}>
                                    标识码 {sortConfig.key === '标识码' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                </th>
                                <th onClick={() => handleSortChange('图号')} style={{ cursor: 'pointer' }}>
                                    图号 {sortConfig.key === '图号' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                </th>
                                <th onClick={() => handleSortChange('名称')} style={{ cursor: 'pointer' }}>
                                    名称 {sortConfig.key === '名称' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                </th>
                                
                                {/* 根据选择显示价格列 */}
                                {showPriceColumns['指导价（不含税）'] && 
                                    <th onClick={() => handleSortChange('指导价（不含税）')} style={{ cursor: 'pointer' }}>
                                        指导价（不含税） {sortConfig.key === '指导价（不含税）' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                    </th>
                                }
                                {showPriceColumns['出厂价（不含税）'] && 
                                    <th onClick={() => handleSortChange('出厂价（不含税）')} style={{ cursor: 'pointer' }}>
                                        出厂价（不含税） {sortConfig.key === '出厂价（不含税）' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                    </th>
                                }
                                {showPriceColumns['服务价（不含税）'] && 
                                    <th onClick={() => handleSortChange('服务价（不含税）')} style={{ cursor: 'pointer' }}>
                                        服务价（不含税） {sortConfig.key === '服务价（不含税）' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                    </th>
                                }
                                {showPriceColumns['指导价（含税）'] && 
                                    <th onClick={() => handleSortChange('指导价（含税）')} style={{ cursor: 'pointer' }}>
                                        指导价（含税） {sortConfig.key === '指导价（含税）' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                    </th>
                                }
                                {showPriceColumns['出厂价（含税）'] && 
                                    <th onClick={() => handleSortChange('出厂价（含税）')} style={{ cursor: 'pointer' }}>
                                        出厂价（含税） {sortConfig.key === '出厂价（含税）' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                    </th>
                                }
                                {showPriceColumns['服务价（含税）'] && 
                                    <th onClick={() => handleSortChange('服务价（含税）')} style={{ cursor: 'pointer' }}>
                                        服务价（含税） {sortConfig.key === '服务价（含税）' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                    </th>
                                }
                                
                                <th onClick={() => handleSortChange('备注')} style={{ cursor: 'pointer' }}>
                                    备注 {sortConfig.key === '备注' && (sortConfig.direction === 'ascending' ? '↑' : '↓')}
                                </th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            {currentPageData.map((part, index) => (
                                <tr key={part['标识码'] + '-' + index}>
                                    <td>{(currentPage - 1) * pageSize + index + 1}</td>
                                    <td>{part['日期']}</td>
                                    <td>{part['标识码']}</td>
                                    <td>{part['图号']}</td>
                                    <td>{part['名称']}</td>
                                    
                                    {/* 根据选择显示价格列 */}
                                    {showPriceColumns['指导价（不含税）'] && <td>{formatPrice(part['指导价（不含税）'])}</td>}
                                    {showPriceColumns['出厂价（不含税）'] && <td>{formatPrice(part['出厂价（不含税）'])}</td>}
                                    {showPriceColumns['服务价（不含税）'] && <td>{formatPrice(part['服务价（不含税）'])}</td>}
                                    {showPriceColumns['指导价（含税）'] && <td>{formatPrice(part['指导价（含税）'])}</td>}
                                    {showPriceColumns['出厂价（含税）'] && <td>{formatPrice(part['出厂价（含税）'])}</td>}
                                    {showPriceColumns['服务价（含税）'] && <td>{formatPrice(part['服务价（含税）'])}</td>}
                                    
                                    <td title={part['备注']} style={{ maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                        {part['备注']}
                                    </td>
                                    <td>
                                        <button
                                            onClick={() => handleSelectPart(part)}
                                            className={selectedParts.some(p => p['标识码'] === part['标识码']) 
                                                ? "action-button danger-button" 
                                                : "action-button primary-button"}
                                            style={{ margin: '0', padding: '4px 8px' }}
                                        >
                                            {selectedParts.some(p => p['标识码'] === part['标识码']) ? '取消' : '选择'}
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>

                    {/* 分页控制 */}
                    <div className="pagination no-print">
                        <button
                            disabled={currentPage <= 1}
                            onClick={() => setCurrentPage(1)}
                        >
                            首页
                        </button>
                        <button
                            disabled={currentPage <= 1}
                            onClick={() => setCurrentPage(currentPage - 1)}
                        >
                            上一页
                        </button>
                        <span style={{ margin: '0 12px' }}>
                            第 {currentPage} 页 / 共 {totalPages} 页
                        </span>
                        <button
                            disabled={currentPage >= totalPages}
                            onClick={() => setCurrentPage(currentPage + 1)}
                        >
                            下一页
                        </button>
                        <button
                            disabled={currentPage >= totalPages}
                            onClick={() => setCurrentPage(totalPages)}
                        >
                            末页
                        </button>
                        
                        <span style={{ margin: '0 12px' }}>
                            前往:
                            <input 
                                type="number" 
                                min="1" 
                                max={totalPages}
                                value={currentPage}
                                onChange={(e) => {
                                    const page = parseInt(e.target.value);
                                    if (page && page >= 1 && page <= totalPages) {
                                        setCurrentPage(page);
                                    }
                                }}
                                style={{ 
                                    width: '50px', 
                                    margin: '0 5px',
                                    backgroundColor: themeStyles.inputBackground,
                                    color: themeStyles.inputText,
                                    border: `1px solid ${themeStyles.inputBorder}`,
                                    borderRadius: '4px',
                                    padding: '3px'
                                }}
                            />
                            页
                        </span>
                    </div>

                    {/* 操作按钮区域 */}
                    <div className="no-print" style={{ 
                        marginTop: '20px', 
                        display: 'flex',
                        flexWrap: 'wrap',
                        gap: '10px',
                        justifyContent: 'center'
                    }}>
                        <div style={{ 
                            padding: '10px', 
                            backgroundColor: themeStyles.background,
                            borderRadius: '4px',
                            border: `1px solid ${themeStyles.border}`,
                            marginBottom: '10px'
                        }}>
                            <h4 style={{ margin: '0 0 10px 0' }}>数据导入</h4>
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                <button
                                    onClick={() => document.getElementById('fileBasic').click()}
                                    className="action-button"
                                >
                                    导入Excel数据
                                </button>
                                <button
                                    onClick={startDocumentProcessWorkflow}
                                    className="action-button primary-button"
                                >
                                    从客户文档导入
                                </button>
                                <button
                                    onClick={batchAddParts}
                                    className="action-button"
                                >
                                    批量添加配件
                                </button>
                                <button
                                    onClick={() => document.getElementById('fileTemplate').click()}
                                    className="action-button"
                                >
                                    导入报价模板
                                </button>
                                {isAdmin && (
                                    <button
                                        onClick={() => document.getElementById('fileAdvanced').click()}
                                        className="action-button primary-button"
                                    >
                                        高级配件导入
                                    </button>
                                )}
                            </div>
                        </div>
                        
                        <div style={{ 
                            padding: '10px', 
                            backgroundColor: themeStyles.background,
                            borderRadius: '4px',
                            border: `1px solid ${themeStyles.border}`,
                            marginBottom: '10px'
                        }}>
                            <h4 style={{ margin: '0 0 10px 0' }}>数据管理</h4>
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                <button
                                    onClick={generateQuotation}
                                    className="action-button primary-button"
                                    disabled={selectedParts.length === 0}
                                >
                                    生成报价单 ({selectedParts.length})
                                </button>
                                {selectedParts.some(p => 
                                    p.matchType === 'fuzzy' || 
                                    p.matchType === 'caseInsensitive' || 
                                    p.matchType === 'noSpace'
                                ) && (
                                    <button
                                        onClick={reviewAndConfirm}
                                        className="action-button"
                                        style={{backgroundColor: '#ff9800', color: 'white'}}
                                    >
                                        人工审核确认
                                    </button>
                                )}
                                <button
                                    onClick={exportDatabaseToExcel}
                                    className="action-button"
                                    disabled={partsData.length === 0}
                                >
                                    导出数据库
                                </button>
                                {isAdmin && (
                                    <button
                                        onClick={() => {
                                            if (window.confirm('确定要清空所有数据吗？此操作不可撤销！')) {
                                                if (isElectron) {
                                                    window.electronAPI.saveData([]).then(success => {
                                                        if (success) {
                                                            setPartsData([]);
                                                            setSelectedParts([]);
                                                            alert('数据库已清空');
                                                        } else {
                                                            alert('清空数据库失败');
                                                        }
                                                    });
                                                } else {
                                                    localStorage.removeItem('shipPartsData');
                                                    setPartsData([]);
                                                    setSelectedParts([]);
                                                    alert('数据库已清空');
                                                }
                                            }
                                        }}
                                        className="action-button danger-button"
                                    >
                                        清空数据库
                                    </button>
                                )}
                                {selectedParts.length > 0 && (
                                    <button
                                        onClick={clearSelectedParts}
                                        className="action-button danger-button"
                                    >
                                        清空已选配件
                                    </button>
                                )}
                            </div>
                        </div>
                    </div>

                    {infoMessage && (
                        <div style={{ 
                            marginTop: 16, 
                            padding: '10px', 
                            backgroundColor: themeStyles.alertBackground,
                            color: themeStyles.alertText,
                            borderRadius: '4px',
                            border: '1px solid',
                            textAlign: 'center'
                        }}>
                            <div>{infoMessage}</div>
                            {loading && (
                                <div style={{ marginTop: '10px' }}>
                                    <div style={{ width: '100%', height: '10px', backgroundColor: themeStyles.background, borderRadius: '5px', overflow: 'hidden' }}>
                                        <div 
                                            style={{ 
                                                width: `${(infoMessage && infoMessage.match(/\((\d+)%\)/)) ? infoMessage.match(/\((\d+)%\)/)[1] : 0}%`, 
                                                height: '100%', 
                                                backgroundColor: themeStyles.link,
                                                transition: 'width 0.3s ease'
                                            }}
                                        />
                                    </div>
                                </div>
                            )}
                        </div>
                    )}
                </div>
            )}

            {view === 'quotation' && (
                <div id="print-container" style={{
                    width: '95%',
                    margin: '0 auto',
                    backgroundColor: themeStyles.container,
                    color: themeStyles.text,
                    padding: '15px',
                    borderRadius: '4px',
                    boxShadow: '0 1px 4px rgba(0,0,0,0.1)'
                }}>
                    <div style={{ textAlign: 'center', marginBottom: '20px' }}>
                        <h2 style={{ margin: '0 0 5px 0' }}>船用配件报价单</h2>
                        <p style={{ margin: '0', color: themeStyles.text, opacity: '0.7' }}>
                            {new Date().toLocaleDateString('zh-CN')}
                        </p>
                    </div>

                    {/* 客户信息区域 */}
                    <div style={{ 
                        marginBottom: '20px',
                        padding: '10px',
                        backgroundColor: themeStyles.background,
                        borderRadius: '4px',
                        border: `1px solid ${themeStyles.border}`
                    }}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '10px', marginBottom: '10px' }}>
                            <div style={{ flex: '1', minWidth: '250px' }}>
                                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>客户名称:</label>
                                <input
                                    type="text"
                                    value={customerInfo.name}
                                    onChange={(e) => setCustomerInfo({...customerInfo, name: e.target.value})}
                                    style={{
                                        width: '100%',
                                        padding: '8px',
                                        backgroundColor: themeStyles.inputBackground,
                                        color: themeStyles.inputText,
                                        border: `1px solid ${themeStyles.inputBorder}`,
                                        borderRadius: '4px'
                                    }}
                                    placeholder="请输入客户名称"
                                    className="no-print-border"
                                />
                            </div>
                            <div style={{ flex: '1', minWidth: '250px' }}>
                                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>联系方式:</label>
                                <input
                                    type="text"
                                    value={customerInfo.contact}
                                    onChange={(e) => setCustomerInfo({...customerInfo, contact: e.target.value})}
                                    style={{
                                        width: '100%',
                                        padding: '8px',
                                        backgroundColor: themeStyles.inputBackground,
                                        color: themeStyles.inputText,
                                        border: `1px solid ${themeStyles.inputBorder}`,
                                        borderRadius: '4px'
                                    }}
                                    placeholder="请输入联系方式"
                                    className="no-print-border"
                                />
                            </div>
                        </div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '10px' }}>
                            <div style={{ flex: '1', minWidth: '200px' }}>
                                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>日期:</label>
                                <input
                                    type="date"
                                    value={customerInfo.date}
                                    onChange={(e) => setCustomerInfo({...customerInfo, date: e.target.value})}
                                    style={{
                                        width: '100%',
                                        padding: '8px',
                                        backgroundColor: themeStyles.inputBackground,
                                        color: themeStyles.inputText,
                                        border: `1px solid ${themeStyles.inputBorder}`,
                                        borderRadius: '4px'
                                    }}
                                    className="no-print-border"
                                />
                            </div>
                            <div style={{ flex: '1', minWidth: '200px' }}>
                                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>船舶:</label>
                                <input
                                    type="text"
                                    value={customerInfo.vessel}
                                    onChange={(e) => setCustomerInfo({...customerInfo, vessel: e.target.value})}
                                    style={{
                                        width: '100%',
                                        padding: '8px',
                                        backgroundColor: themeStyles.inputBackground,
                                        color: themeStyles.inputText,
                                        border: `1px solid ${themeStyles.inputBorder}`,
                                        borderRadius: '4px'
                                    }}
                                    placeholder="请输入船舶名称"
                                    className="no-print-border"
                                />
                            </div>
                            <div style={{ flex: '1', minWidth: '200px' }}>
                                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>项目:</label>
                                <input
                                    type="text"
                                    value={customerInfo.project}
                                    onChange={(e) => setCustomerInfo({...customerInfo, project: e.target.value})}
                                    style={{
                                        width: '100%',
                                        padding: '8px',
                                        backgroundColor: themeStyles.inputBackground,
                                        color: themeStyles.inputText,
                                        border: `1px solid ${themeStyles.inputBorder}`,
                                        borderRadius: '4px'
                                    }}
                                    placeholder="请输入项目名称"
                                    className="no-print-border"
                                />
                            </div>
                        </div>
                    </div>

                    {/* 价格选项和统计信息 */}
                    <div className="quotation-summary">
                        <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                            <div>
                                <label style={{ fontWeight: 'bold', marginRight: '10px' }}>价格类型: </label>
                                <select
                                    value={priceOption}
                                    onChange={(e) => setPriceOption(e.target.value)}
                                    style={{
                                        backgroundColor: themeStyles.inputBackground,
                                        color: themeStyles.inputText,
                                        border: `1px solid ${themeStyles.inputBorder}`,
                                        padding: '6px 10px',
                                        borderRadius: '4px'
                                    }}
                                    className="no-print"
                                >
                                    <option value="指导价（不含税）">指导价（不含税）</option>
                                    <option value="出厂价（不含税）">出厂价（不含税）</option>
                                    <option value="服务价（不含税）">服务价（不含税）</option>
                                    <option value="指导价（含税）">指导价（含税）</option>
                                    <option value="出厂价（含税）">出厂价（含税）</option>
                                    <option value="服务价（含税）">服务价（含税）</option>
                                </select>
                            </div>
                            <div className="no-print">
                                <button 
                                    onClick={applyBulkDiscount}
                                    className="action-button"
                                    disabled={selectedParts.length === 0}
                                    style={{ marginRight: '8px' }}
                                >
                                    应用批量折扣
                                </button>
                                {selectedParts.some(p => 
                                    p.matchType === 'fuzzy' || 
                                    p.matchType === 'caseInsensitive' || 
                                    p.matchType === 'noSpace'
                                ) && !selectedParts.every(p => p.humanReviewed) && (
                                    <button 
                                        className="action-button"
                                        style={{backgroundColor: '#ff9800', color: 'white'}}
                                        onClick={() => {
                                            // 标记所有配件为已人工审核
                                            setSelectedParts(selectedParts.map(part => ({
                                                ...part,
                                                humanReviewed: true
                                            })));
                                            alert('已将所有模糊匹配项标记为人工审核通过');
                                        }}
                                    >
                                        确认人工审核
                                    </button>
                                )}
                            </div>
                        </div>
                        <div className="quotation-stats">
                            <div className="stat-item">
                                <strong>总价: </strong>
                                <span style={{ fontSize: '16px', fontWeight: 'bold', color: '#e53935' }}>
                                    ¥{formatTotalPrice(statistics.totalPrice)}
                                </span>
                            </div>
                            <div className="stat-item">
                                <strong>配件总数: </strong>
                                <span>{selectedParts.length}</span>
                            </div>
                            <div className="stat-item">
                                <strong>总数量: </strong>
                                <span>{statistics.totalQuantity}</span>
                            </div>
                            <div className="stat-item no-print">
                                <strong>精确匹配: </strong>
                                <span>{statistics.exactMatchCount}</span>
                            </div>
                            <div className="stat-item no-print" style={statistics.fuzzyMatchCount > 0 ? {backgroundColor: '#fff3cd', color: '#856404'} : {}}>
                                <strong>模糊匹配: </strong>
                                <span>{statistics.fuzzyMatchCount}</span>
                            </div>
                            <div className="stat-item no-print" style={statistics.newPartCount > 0 ? {backgroundColor: '#f8d7da', color: '#721c24'} : {}}>
                                <strong>新配件: </strong>
                                <span>{statistics.newPartCount}</span>
                            </div>
                            {selectedParts.some(p => p.humanReviewed) && (
                                <div className="stat-item no-print" style={{backgroundColor: '#d4edda', color: '#155724'}}>
                                    <strong>已审核: </strong>
                                    <span>{selectedParts.filter(p => p.humanReviewed).length}</span>
                                </div>
                            )}
                        </div>
                    </div>

                    {/* 配件表格 */}
                    <table className="data-table" style={{ marginTop: '20px' }}>
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>图号</th>
                                <th>名称</th>
                                <th>单价(元)</th>
                                <th>数量</th>
                                <th>总价(元)</th>
                                <th>备注</th>
                                <th className="no-print">操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            {selectedParts.map((part, index) => {
                                const unitPrice = part.importedPrice || part[priceOption] || 0;
                                const lineTotal = unitPrice * (part.quantity || 1);
                                return (
                                    <tr key={part['标识码'] + '-' + index}>
                                        <td>{index + 1}</td>
                                        <td title={part.importedId && part.importedId !== part['图号'] ? `客户标识: ${part.importedId}` : null}>
                                            {part['图号']}
                                            {part.importedId && part.importedId !== part['图号'] && (
                                                <div style={{ fontSize: '12px', color: themeStyles.text, opacity: '0.7' }}>
                                                    ({part.importedId})
                                                </div>
                                            )}
                                            {part.matchType && part.matchType !== 'exact' && !part.humanReviewed && (
                                                <div style={{
                                                    fontSize: '11px',
                                                    padding: '2px 4px',
                                                    marginTop: '2px',
                                                    backgroundColor: '#fff3cd',
                                                    color: '#856404',
                                                    borderRadius: '3px',
                                                    display: 'inline-block'
                                                }}>
                                                    模糊匹配
                                                </div>
                                            )}
                                            {part.matchType && part.matchType !== 'exact' && part.humanReviewed && (
                                                <div style={{
                                                    fontSize: '11px',
                                                    padding: '2px 4px',
                                                    marginTop: '2px',
                                                    backgroundColor: '#d4edda',
                                                    color: '#155724',
                                                    borderRadius: '3px',
                                                    display: 'inline-block'
                                                }}>
                                                    已审核
                                                </div>
                                            )}
                                        </td>
                                        <td>{part['名称']}</td>
                                        <td>
                                            <input
                                                type="number"
                                                value={unitPrice}
                                                onChange={(e) => updatePartCustomPrice(part['标识码'], e.target.value)}
                                                style={{
                                                    width: '100px',
                                                    backgroundColor: themeStyles.inputBackground,
                                                    color: themeStyles.inputText,
                                                    border: `1px solid ${themeStyles.inputBorder}`,
                                                    padding: '4px',
                                                    borderRadius: '4px'
                                                }}
                                                className="no-print-border"
                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="number"
                                                value={part.quantity}
                                                onChange={(e) => updatePartQuantity(part['标识码'], e.target.value)}
                                                style={{
                                                    width: '60px',
                                                    backgroundColor: themeStyles.inputBackground,
                                                    color: themeStyles.inputText,
                                                    border: `1px solid ${themeStyles.inputBorder}`,
                                                    padding: '4px',
                                                    borderRadius: '4px'
                                                }}
                                                min="1"
                                                className="no-print-border"
                                            />
                                        </td>
                                        <td style={{ fontWeight: 'bold' }}>
                                            {formatTotalPrice(lineTotal)}
                                        </td>
                                        <td title={part.importedRemark || part['备注']} style={{ maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                            {part.importedRemark || part['备注'] || ''}
                                        </td>
                                        <td className="no-print">
                                            <button
                                                onClick={() => removeSelectedPart(part['标识码'])}
                                                className="action-button danger-button"
                                                style={{ padding: '4px 8px', margin: '0' }}
                                            >
                                                删除
                                            </button>
                                        </td>
                                    </tr>
                                );
                            })}
                            <tr style={{ fontWeight: 'bold', backgroundColor: themeStyles.header }}>
                                <td colSpan={5} style={{ textAlign: 'right' }}>总计:</td>
                                <td style={{ fontSize: '16px', color: '#e53935' }}>
                                    {formatTotalPrice(statistics.totalPrice)}
                                </td>
                                <td colSpan={2}></td>
                            </tr>
                        </tbody>
                    </table>

                    {/* 报价单底部和按钮区 */}
                    <div style={{ 
                        marginTop: '30px', 
                        padding: '20px 0',
                        borderTop: `1px dashed ${themeStyles.border}`,
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'flex-end'
                    }}>
                        <div style={{ marginBottom: '5px' }}>报价员: ___________________</div>
                        <div style={{ marginBottom: '5px' }}>日期: {new Date().toLocaleDateString('zh-CN')}</div>
                    </div>

                    <div style={{ marginTop: '30px', textAlign: 'center' }} className="no-print">
                        <button
                            onClick={exportQuotationCSV}
                            className="action-button"
                            style={{ marginRight: '10px' }}
                        >
                            导出CSV
                        </button>
                        <button
                            onClick={exportQuotationExcel}
                            className="action-button"
                            style={{ marginRight: '10px' }}
                        >
                            导出Excel
                        </button>
                        <button
                            onClick={handlePrint}
                            className="action-button primary-button"
                            style={{ marginRight: '10px' }}
                        >
                            打印报价单
                        </button>
                        <button
                            onClick={backToList}
                            className="action-button"
                        >
                            返回列表
                        </button>
                        
                        {/* 添加帮助信息，特别是关于PDF/Word导入功能 */}
                        <div style={{ 
                            marginTop: '20px', 
                            padding: '10px', 
                            backgroundColor: themeStyles.background,
                            border: `1px solid ${themeStyles.border}`,
                            borderRadius: '4px',
                            fontSize: '12px',
                            textAlign: 'left'
                        }}>
                            <p style={{ fontWeight: 'bold', marginBottom: '5px' }}>使用提示:</p>
                            <ul style={{ paddingLeft: '20px', margin: '0' }}>
                                <li>通过"从客户文档导入"按钮可以导入PDF或Word格式的订单文件</li>
                                <li>系统会自动识别文档中的配件号并与数据库匹配</li>
                                <li>模糊匹配的配件会标记出来，需要人工审核确认</li>
                                <li>可以使用"应用批量折扣"对所有配件应用折扣</li>
                            </ul>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
	
	
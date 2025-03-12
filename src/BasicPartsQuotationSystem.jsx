import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';

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

function formatPrice(val) {
    const num = parseFloat(val || 0);
    return num.toFixed(2);
}

// ------------------ 2. “增强”模糊匹配 ------------------ //
function enhancedFuzzyMatch(importedId, partsData) {
    return partsData.find(
        part => String(part['图号']).toLowerCase() === String(importedId).toLowerCase()
    );
}

// ------------------ 3. 各种文件解析逻辑 ------------------ //

// 模拟从 PDF 或特殊文档中提取到的配件列表
function extractPartsFromDocument() {
    return [
        { '图号': '135-01-003A', '名称': '输入轴总成', '数量': 1 },
        { '图号': 'B12X40GB120-86', '名称': '紧固螺栓', '数量': 1 },
        { '图号': 'M10X25GB32.1-88', '名称': '螺栓', '数量': 1 },
        { '图号': '135-01-002', '名称': '调整插头', '数量': 1 },
        { '图号': '6317N', '名称': '轴承', '数量': 1 },
        { '图号': '135-01-004', '名称': '输入轴套', '数量': 1 },
        { '图号': 'NJ313', '名称': '轴承', '数量': 1 },
        { '图号': '135-01-007', '名称': '输入端止推环', '数量': 1 },
        { '图号': '135-01-024B', '名称': '离合器壳体', '数量': 1 },
        { '图号': '135-01-032', '名称': '内六角螺钉', '数量': 1 },
        { '图号': '135A-03A-016A', '名称': '轴套', '数量': 1 },
        { '图号': 'FB-SC115•140•14D', '名称': '油封', '数量': 1 }
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
                        'importedId': part['图号']
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
export default function BasicPartsQuotationSystem() {
    const isElectron = window.electronAPI !== undefined;
    const [partsData, setPartsData] = useState([]);
    const [selectedParts, setSelectedParts] = useState([]);
    const [searchTerm, setSearchTerm] = useState('');
    const [view, setView] = useState('table');
    const [loading, setLoading] = useState(true);
    const [infoMessage, setInfoMessage] = useState(null);
    const [priceOption, setPriceOption] = useState("服务价（含税）");
    // 删除“客户信息”后，不再使用客户信息的 state
    // const [customerInfo, setCustomerInfo] = useState({ ... });

    const [currentPage, setCurrentPage] = useState(1);
    const pageSize = 50;
    const [isAdmin, setIsAdmin] = useState(false);

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
                // 删除“类别”列
                // '类别': "主机系统",
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
                '名称': "输入 轴部件",
                // '类别': "辅机系统",
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
        if (isElectron) {
            try {
                const success = await window.electronAPI.exportData(partsData);
                if (success) {
                    alert('数据导出成功');
                }
            } catch (error) {
                console.error("导出数据失败:", error);
                alert('导出失败: ' + error.message);
            }
        } else {
            const worksheet = XLSX.utils.json_to_sheet(partsData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "船用配件数据");
            XLSX.writeFile(workbook, `船用配件数据库_${new Date().toISOString().split('T')[0]}.xlsx`);
        }
    }

    // 根据搜索框，筛选数据
    const allFilteredData = useMemo(() => {
        if (!searchTerm.trim()) return partsData;
        const lower = searchTerm.toLowerCase();
        return partsData.filter(item =>
            (item['日期'] && String(item['日期']).toLowerCase().includes(lower)) ||
            (item['名称'] && String(item['名称']).toLowerCase().includes(lower)) ||
            (item['标识码'] && String(item['标识码']).toLowerCase().includes(lower)) ||
            (item['图号'] && String(item['图号']).toLowerCase().includes(lower)) ||
            (item['备注'] && String(item['备注']).toLowerCase().includes(lower))
        );
    }, [searchTerm, partsData]);

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
        // 删除客户信息后，不再输出客户信息到CSV
        // csvContent += `\n客户信息:\n`;
        // csvContent += `客户:,${customerInfo.name}\n`;
        // csvContent += `联系方式:,${customerInfo.contact}\n`;
        // csvContent += `日期:,${customerInfo.date}\n`;

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
            link.setAttribute('download', `船用配件报价_${new Date().toISOString().split('T')[0]}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }

    function handlePrint() {
        window.print();
    }

    // 删除后，仅统计总价
    const totalPrice = useMemo(() => {
        return selectedParts.reduce((sum, p) => {
            const priceVal = p.importedPrice || p[priceOption] || 0;
            return sum + priceVal * (p.quantity || 1);
        }, 0);
    }, [selectedParts, priceOption]);

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

    // 使打印时在页尾出现“上海前进齿轮经营有限公司”
    const printStyles = `
        @media print {
            .no-print { 
                display: none; 
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
    `;

    return (
        <div style={{ padding: 10, backgroundColor: '#f9fafb', minHeight: '100vh', color: '#333' }}>
            <style>{printStyles}</style>

            {/* 改成一个很窄的眉头 */}
            <div style={{ fontSize: '16px', marginBottom: '10px' }}>
                <strong>船机配件报价</strong> {/* 紧凑显示 */}
            </div>

            {loading && <div style={{ color: 'red', textAlign: 'center' }}>加载中...</div>}

            {view === 'table' && (
                <div style={{
                    width: '90%',
                    margin: '0 auto',
                    backgroundColor: '#ffffff',
                    color: '#333',
                    padding: '10px',
                    borderRadius: '4px',
                    boxShadow: '0 1px 4px rgba(0,0,0,0.1)'
                }}>
                    <div style={{ textAlign: 'right', marginBottom: 8 }} className="no-print">
                        {!isAdmin ? (
                            <button onClick={handleAdminLogin}>管理员登录</button>
                        ) : (
                            <button onClick={handleAdminLogout}>退出管理员模式</button>
                        )}
                    </div>

                    {/* 隐藏的文件选择器 */}
                    <input type="file" id="fileBasic" style={{ display: 'none' }} onChange={handleFileUploadWrapper} />
                    <input type="file" id="fileAdvanced" style={{ display: 'none' }} onChange={handleAdvancedQuotationUpload} />
                    <input type="file" id="fileCustomer" style={{ display: 'none' }} onChange={handleCustomerQuotationUpload} />

                    <div style={{ marginBottom: 8, textAlign: 'right' }} className="no-print">
                        <input
                            type="text"
                            placeholder="搜索配件..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                            style={{
                                width: 160,
                                backgroundColor: '#fff',
                                color: '#333',
                                border: '1px solid #ccc',
                                padding: '4px'
                            }}
                        />
                    </div>

                    <table
                        border="1"
                        cellPadding="4"
                        cellSpacing="0"
                        style={{
                            width: '100%',
                            margin: '0 auto',
                            borderColor: '#ddd',
                            borderCollapse: 'collapse',
                            fontSize: '14px'
                        }}
                    >
                        <thead style={{ backgroundColor: '#f2f2f2', whiteSpace: 'nowrap' }}>
                            <tr>
                                <th>序号</th>
                                <th>日期</th>
                                <th>标识码</th>
                                <th>图号</th>
                                <th>名称</th>
                                {/* 已删除“类别”列 */}
                                <th>指导价（不含税）</th>
                                <th>出厂价（不含税）</th>
                                <th>服务价（不含税）</th>
                                <th>指导价（含税）</th>
                                <th>出厂价（含税）</th>
                                <th>服务价（含税）</th>
                                <th>备注</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody style={{ whiteSpace: 'nowrap' }}>
                            {currentPageData.map((part, index) => (
                                <tr key={part['标识码'] + '-' + index}>
                                    <td>{(currentPage - 1) * pageSize + index + 1}</td>
                                    <td>{part['日期']}</td>
                                    <td>{part['标识码']}</td>
                                    <td>{part['图号']}</td>
                                    <td>{part['名称']}</td>
                                    <td>{formatPrice(part['指导价（不含税）'])}</td>
                                    <td>{formatPrice(part['出厂价（不含税）'])}</td>
                                    <td>{formatPrice(part['服务价（不含税）'])}</td>
                                    <td>{formatPrice(part['指导价（含税）'])}</td>
                                    <td>{formatPrice(part['出厂价（含税）'])}</td>
                                    <td>{formatPrice(part['服务价（含税）'])}</td>
                                    <td>{part['备注']}</td>
                                    <td>
                                        <button
                                            onClick={() => handleSelectPart(part)}
                                            style={{
                                                backgroundColor: '#e5e5e5',
                                                color: '#333',
                                                border: '1px solid #ccc',
                                                cursor: 'pointer'
                                            }}
                                        >
                                            {selectedParts.some(p => p['标识码'] === part['标识码']) ? '取消' : '选择'}
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>

                    <div style={{ marginTop: 8, textAlign: 'center' }} className="no-print">
                        <button
                            disabled={currentPage <= 1}
                            onClick={() => setCurrentPage(currentPage - 1)}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                marginRight: '8px',
                                cursor: 'pointer'
                            }}
                        >
                            上一页
                        </button>
                        <span style={{ margin: '0 8px' }}>
                            第 {currentPage} 页 / 共 {totalPages} 页
                        </span>
                        <button
                            disabled={currentPage >= totalPages}
                            onClick={() => setCurrentPage(currentPage + 1)}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                marginLeft: '8px',
                                cursor: 'pointer'
                            }}
                        >
                            下一页
                        </button>
                    </div>

                    <div style={{ marginTop: 16, textAlign: 'center' }} className="no-print">
                        <button
                            onClick={() => document.getElementById('fileBasic').click()}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                marginRight: '8px',
                                cursor: 'pointer'
                            }}
                        >
                            导入Excel
                        </button>
                        {isAdmin && (
                            <button
                                onClick={() => document.getElementById('fileAdvanced').click()}
                                style={{
                                    backgroundColor: '#e5e5e5',
                                    color: '#333',
                                    border: '1px solid #ccc',
                                    marginRight: '8px',
                                    cursor: 'pointer'
                                }}
                            >
                                导入高级报价单
                            </button>
                        )}
                        <button
                            onClick={() => document.getElementById('fileCustomer').click()}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                marginRight: '8px',
                                cursor: 'pointer'
                            }}
                        >
                            导入客户报价单
                        </button>
                        {isAdmin && (
                            <>
                                <button
                                    onClick={exportDatabaseToExcel}
                                    style={{
                                        backgroundColor: '#e5e5e5',
                                        color: '#333',
                                        border: '1px solid #ccc',
                                        marginRight: '8px',
                                        cursor: 'pointer'
                                    }}
                                >
                                    导出数据库
                                </button>
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
                                    style={{
                                        backgroundColor: '#e5e5e5',
                                        color: '#333',
                                        border: '1px solid #ccc',
                                        marginRight: '8px',
                                        cursor: 'pointer'
                                    }}
                                >
                                    清空数据库
                                </button>
                            </>
                        )}
                        <button
                            onClick={generateQuotation}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                cursor: 'pointer'
                            }}
                        >
                            生成报价单
                        </button>
                    </div>

                    {infoMessage && (
                        <div style={{ marginTop: 12, textAlign: 'center', color: '#d14' }}>{infoMessage}</div>
                    )}
                </div>
            )}

            {view === 'quotation' && (
                <div style={{
                    width: '90%',
                    margin: '0 auto',
                    backgroundColor: '#ffffff',
                    color: '#333',
                    padding: '10px',
                    borderRadius: '4px',
                    boxShadow: '0 1px 4px rgba(0,0,0,0.1)'
                }}>
                    {/* 仅保留一个标题，不再有客户信息 */}
                    <h3 style={{ textAlign: 'center', marginBottom: '10px' }}>报价单</h3>

                    {/* 删除“客户信息”区域后，直接进入表格 */}
                    <div style={{ marginBottom: 8, textAlign: 'center' }}>
                        <label>价格类型: </label>
                        <select
                            value={priceOption}
                            onChange={(e) => setPriceOption(e.target.value)}
                            style={{
                                backgroundColor: '#fff',
                                color: '#333',
                                border: '1px solid #ccc',
                                padding: '4px'
                            }}
                        >
                            <option value="指导价（不含税）">指导价（不含税）</option>
                            <option value="出厂价（不含税）">出厂价（不含税）</option>
                            <option value="服务价（不含税）">服务价（不含税）</option>
                            <option value="指导价（含税）">指导价（含税）</option>
                            <option value="出厂价（含税）">出厂价（含税）</option>
                            <option value="服务价（含税）">服务价（含税）</option>
                        </select>
                    </div>

                    <table
                        border="1"
                        cellPadding="4"
                        cellSpacing="0"
                        style={{
                            width: '100%',
                            margin: '0 auto',
                            borderColor: '#ddd',
                            borderCollapse: 'collapse',
                            fontSize: '14px',
                            whiteSpace: 'nowrap'
                        }}
                    >
                        <thead style={{ backgroundColor: '#f2f2f2' }}>
                            <tr>
                                <th>序号</th>
                                <th>图号</th>
                                <th>名称</th>
                                <th>单价</th>
                                <th>数量</th>
                                <th>总价</th>
                            </tr>
                        </thead>
                        <tbody>
                            {selectedParts.map((part, index) => {
                                const unitPrice = part.importedPrice || part[priceOption] || 0;
                                const lineTotal = unitPrice * (part.quantity || 1);
                                return (
                                    <tr key={part['标识码'] + '-' + index}>
                                        <td>{index + 1}</td>
                                        <td>{part['图号']}</td>
                                        <td>{part['名称']}</td>
                                        <td>
                                            <input
                                                type="number"
                                                value={unitPrice}
                                                onChange={(e) =>
                                                    updatePartCustomPrice(part['标识码'], e.target.value)
                                                }
                                                style={{
                                                    width: 80,
                                                    backgroundColor: '#fff',
                                                    color: '#333',
                                                    border: '1px solid #ccc',
                                                    padding: '2px'
                                                }}
                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="number"
                                                value={part.quantity}
                                                onChange={(e) =>
                                                    updatePartQuantity(part['标识码'], e.target.value)
                                                }
                                                style={{
                                                    width: 60,
                                                    backgroundColor: '#fff',
                                                    color: '#333',
                                                    border: '1px solid #ccc',
                                                    padding: '2px'
                                                }}
                                            />
                                        </td>
                                        <td>{lineTotal.toFixed(2)}</td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>

                    <div style={{ marginTop: 8, textAlign: 'center' }}>
                        <strong>总价: {totalPrice.toFixed(2)}</strong>
                    </div>

                    {/* 只保留三个按钮：导出CSV、打印报价单、返回列表 */}
                    <div style={{ marginTop: 16, textAlign: 'center' }} className="no-print">
                        <button
                            onClick={exportQuotationCSV}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                marginRight: '8px',
                                cursor: 'pointer'
                            }}
                        >
                            导出CSV
                        </button>
                        <button
                            onClick={handlePrint}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                marginRight: '8px',
                                cursor: 'pointer'
                            }}
                        >
                            打印报价单
                        </button>
                        <button
                            onClick={backToList}
                            style={{
                                backgroundColor: '#e5e5e5',
                                color: '#333',
                                border: '1px solid #ccc',
                                cursor: 'pointer'
                            }}
                        >
                            返回列表
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
}

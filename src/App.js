import React from 'react';
import './App.css';
import BasicPartsQuotationSystem from './BasicPartsQuotationSystem';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>船用配件报价系统</h1>
      </header>
      <main>
        <BasicPartsQuotationSystem />
      </main>
    </div>
  );
}

export default App;
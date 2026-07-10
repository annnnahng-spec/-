<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>持份項目日報 - 專業數據分析平台</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- React 18 -->
    <script src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
    <script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
    <!-- Babel for JSX -->
    <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
    <!-- SheetJS (Excel Parser) -->
    <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700;900&display=swap');
        
        body { 
            background-color: #fcf8f8; 
            font-family: 'Inter', 'Noto Sans TC', sans-serif;
        }
        /* 中海地產風格：經典朱紅、深紅、白底 */
        .theme-bg-gradient {
            background: linear-gradient(135deg, #991b1b 0%, #7f1d1d 100%);
        }
        .theme-border-highlight {
            border-top: 4px solid #b91c1c;
        }
        .card-shadow {
            box-shadow: 0 4px 20px -2px rgba(153, 27, 27, 0.06), 0 2px 8px -1px rgba(0, 0, 0, 0.04);
        }
        .custom-scrollbar::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: #fecaca;
            border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
            background: #f87171;
        }
    </style>
</head>
<body class="text-slate-800 antialiased">
    <div id="root"></div>

    <script type="text/babel">
        const { useState, useEffect, useRef, useMemo } = React;

        const PROJECT_MAP = {
            'KTM': '6577 啟德海灣',
            'OV': '6575 維港1號',
            'TV': '6603 維港．雙鑽',
            'VV': '6554 維港 ‧ 灣畔',
            'KB': '6552 天瀧',
            'DC': '6576 Double Coast',
            'GM': '柏瓏',
            'CH': '6551 天璽海',
            'PANO': '6553 澐璟',
            'MQ': '6574 Miami Quay',
            'PF': '6591 柏蔚森'
        };

        const colLetterToIndex = (letter) => {
            let column = 0;
            const cleanLetter = letter.trim().toUpperCase();
            for (let i = 0; i < cleanLetter.length; i++) {
                column = column * 26 + (cleanLetter.charCodeAt(i) - 64);
            }
            return column - 1; // Return 0-based index
        };

        const parseExcelDate = (val) => {
            if (val === undefined || val === null || val === '') return null;
            
            // If it's already a standard date-looking string (e.g. "2026-03-01" or "2026/03/01")
            if (typeof val === 'string') {
                const cleanStr = val.trim();
                if (cleanStr.includes('/') || cleanStr.includes('-')) {
                    const parsed = new Date(cleanStr);
                    if (!isNaN(parsed.getTime())) {
                        return parsed.toISOString().split('T')[0];
                    }
                }
                // Skip headers
                if (cleanStr.includes('日期') || cleanStr.includes('Date')) return null;
                return cleanStr;
            }
            
            // If it's Excel numeric date format (e.g. 46082)
            if (typeof val === 'number') {
                try {
                    // Excel dates start on Dec 30, 1899 due to leap year bug in Lotus 1-2-3
                    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
                    if (!isNaN(date.getTime())) {
                        return date.toISOString().split('T')[0];
                    }
                } catch (e) {
                    return null;
                }
            }
            return String(val);
        };

        const App = () => {
            const [rawData, setRawData] = useState([]);
            const [allDates, setAllDates] = useState([]);
            const [selectedDate, setSelectedDate] = useState('');
            const [logs, setLogs] = useState([]);
            const [activeSheet, setActiveSheet] = useState('');
            const [showLogs, setShowLogs] = useState(true);
            const [searchQuery, setSearchQuery] = useState('');

            // Chart references
            const barChartRef = useRef(null);
            const doughnutChartRef = useRef(null);
            const barChartInstance = useRef(null);
            const doughnutChartInstance = useRef(null);

            const addLog = (message, type = 'info') => {
                const timestamp = new Date().toLocaleTimeString();
                setLogs(prev => [{ text: `[${timestamp}] ${message}`, type }, ...prev]);
            };

            const handleFileUpload = (e) => {
                const file = e.target.files[0];
                if (!file) {
                    addLog('未選擇任何檔案', 'error');
                    return;
                }

                addLog(`開始讀取檔案: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`, 'info');
                const reader = new FileReader();
                
                reader.onload = (event) => {
                    try {
                        const binaryData = event.target.result;
                        const workbook = XLSX.read(binaryData, { type: 'binary' });
                        
                        addLog(`成功解析 Excel 活頁簿！含有以下分頁: ${workbook.SheetNames.join(', ')}`, 'success');
                        
                        // 優先尋找名稱包含 "PFMQCHPANO" 的工作表，否則預設取第一張
                        let targetSheetName = workbook.SheetNames.find(
                            name => name.toUpperCase().includes('PFMQCHPANO')
                        );
                        
                        if (!targetSheetName) {
                            targetSheetName = workbook.SheetNames[0];
                            addLog(`未找到指定名稱包含「PFMQCHPANO」的分頁，已自動載入第一個分頁: 「${targetSheetName}」`, 'warning');
                        } else {
                            addLog(`已成功定位至目標分頁: 「${targetSheetName}」`, 'success');
                        }
                        
                        setActiveSheet(targetSheetName);
                        const worksheet = workbook.Sheets[targetSheetName];
                        
                        // 使用 { header: 1 } 以取得二維數組，並啟用 defval 填充空單元格，確保列索引絕對正確
                        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                        
                        addLog(`分頁「${targetSheetName}」讀取到 ${rows.length} 列數據`, 'info');
                        
                        if (rows.length < 2) {
                            addLog('Excel 檔案中的數據列數不足（少於 2 列），無法分析', 'error');
                            return;
                        }

                        // 欄位物理映射設定 (根據您的描述將字母轉換成對應索引)
                        const mapConfig = {
                            date: colLetterToIndex('R'),        // 日期 (R 欄)
                            project: colLetterToIndex('B'),     // 項目 (B 欄)
                            area: colLetterToIndex('J'),        // 實用面積 (J 欄)
                            price: colLetterToIndex('S'),       // 折實價 (S 欄)
                            unitPrice: colLetterToIndex('T'),   // 折實呎價 (T 欄)
                            view: colLetterToIndex('V'),        // 景觀 (V 欄)
                            type: colLetterToIndex('K'),        // 戶型 (K 欄)
                            identity: colLetterToIndex('AI'),    // 身份 (AI 欄)
                            age: colLetterToIndex('AK'),         // 年齡 (AK 欄)
                            region: colLetterToIndex('AL'),      // 居住區域 (AL 欄)
                            agent: colLetterToIndex('AM')        // 用途代理 (AM 欄)
                        };

                        addLog('欄位字母與物理索引對應已設定完畢。開始逐行解析數據...', 'info');

                        const parsedList = [];
                        let skippedRows = 0;

                        // 從第 0 行開始掃描（通常第 0 行可能是標頭，後面程式會自動過濾非資料列）
                        for (let i = 0; i < rows.length; i++) {
                            const row = rows[i];
                            
                            // 讀取原始資料值
                            const rawDateVal = row[mapConfig.date];
                            const rawProjectVal = row[mapConfig.project];
                            
                            // 解析與校正日期
                            const formattedDate = parseExcelDate(rawDateVal);
                            
                            // 過濾機制：若日期為空、或者是欄位標頭字樣，則忽略
                            if (!formattedDate || formattedDate === '日期' || String(rawProjectVal).trim() === '項目') {
                                skippedRows++;
                                continue;
                            }

                            const cleanProjectKey = String(rawProjectVal || '').trim();
                            const finalProjectName = PROJECT_MAP[cleanProjectKey] || cleanProjectKey || '其他/未知項目';

                            // 讀取數值，去除千分號與貨幣符號
                            const cleanNum = (val) => {
                                if (!val) return 0;
                                const parsed = parseFloat(String(val).replace(/[^0-9.-]/g, ''));
                                return isNaN(parsed) ? 0 : parsed;
                            };

                            const area = cleanNum(row[mapConfig.area]);
                            const price = cleanNum(row[mapConfig.price]);
                            const unitPrice = cleanNum(row[mapConfig.unitPrice]);

                            parsedList.push({
                                rawIndex: i + 1,
                                date: formattedDate,
                                projectKey: cleanProjectKey,
                                projectName: finalProjectName,
                                area: area,
                                price: price,
                                unitPrice: unitPrice > 0 ? unitPrice : (area > 0 ? Math.round(price / area) : 0),
                                view: String(row[mapConfig.view] || '').trim(),
                                type: String(row[mapConfig.type] || '').trim(),
                                identity: String(row[mapConfig.identity] || '').trim(),
                                age: String(row[mapConfig.age] || '').trim(),
                                region: String(row[mapConfig.region] || '').trim(),
                                agent: String(row[mapConfig.agent] || '').trim()
                            });
                        }

                        addLog(`解析完畢！過濾無效列(如空行、標題) ${skippedRows} 列，成功提取 ${parsedList.length} 筆成交交易記錄`, 'success');

                        if (parsedList.length === 0) {
                            addLog('警告：未能成功解析出任何有效交易數據！請檢查 Excel 是否為空或格式是否吻合。', 'error');
                            // 提供前兩行的結構預覽協助除錯
                            const sampleRows = rows.slice(0, 3);
                            addLog(`Excel 前 3 列結構預覽 (JSON): ${JSON.stringify(sampleRows, null, 2)}`, 'warning');
                            return;
                        }

                        // 抓取不重複的日期並降序排列
                        const uniqueDates = [...new Set(parsedList.map(item => item.date))].sort().reverse();
                        addLog(`偵測到交易日期共有 ${uniqueDates.length} 天。最新的交易日期為: ${uniqueDates[0]}`, 'success');

                        setRawData(parsedList);
                        setAllDates(uniqueDates);
                        setSelectedDate(uniqueDates[0]);

                    } catch (err) {
                        addLog(`解析過程發生重大錯誤: ${err.message}`, 'error');
                        console.error(err);
                    }
                };

                reader.onerror = (err) => {
                    addLog('檔案讀取失敗', 'error');
                };

                reader.readAsBinaryString(file);
            };

            const dailyMetrics = useMemo(() => {
                if (!selectedDate || rawData.length === 0) return null;
                
                const filteredRows = rawData.filter(r => r.date === selectedDate);
                const totalSales = filteredRows.reduce((sum, r) => sum + r.price, 0);
                const totalArea = filteredRows.reduce((sum, r) => sum + r.area, 0);
                const avgPricePerSqft = totalArea > 0 ? (totalSales / totalArea) : 0;
                
                return {
                    totalSales,
                    totalArea,
                    avgPricePerSqft,
                    rows: filteredRows
                };
            }, [rawData, selectedDate]);

            useEffect(() => {
                if (!dailyMetrics || dailyMetrics.rows.length === 0) return;

                // 1. Bar Chart:成交項目與金額
                if (barChartRef.current) {
                    if (barChartInstance.current) barChartInstance.current.destroy();

                    // 按項目名稱分組彙總當日成交額
                    const projectSummary = {};
                    dailyMetrics.rows.forEach(item => {
                        projectSummary[item.projectName] = (projectSummary[item.projectName] || 0) + item.price;
                    });

                    const labels = Object.keys(projectSummary);
                    const dataValues = Object.values(projectSummary);

                    const ctx = barChartRef.current.getContext('2d');
                    barChartInstance.current = new Chart(ctx, {
                        type: 'bar',
                        data: {
                            labels: labels,
                            datasets: [{
                                label: '當日成交總額 (港元)',
                                data: dataValues,
                                backgroundColor: 'rgba(185, 28, 28, 0.85)', // 中海紅
                                borderColor: 'rgba(153, 27, 27, 1)',
                                borderWidth: 1.5,
                                borderRadius: 6,
                                barPercentage: 0.5
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: { display: false },
                                tooltip: {
                                    callbacks: {
                                        label: function(context) {
                                            return ` 成交金額: $${context.raw.toLocaleString()} 元`;
                                        }
                                    }
                                }
                            },
                            scales: {
                                y: {
                                    beginAtZero: true,
                                    grid: { color: '#f3f4f6' },
                                    ticks: {
                                        callback: function(value) {
                                            return '$' + (value / 10000).toLocaleString() + '萬';
                                        }
                                    }
                                },
                                x: { grid: { display: false } }
                            }
                        }
                    });
                }

                // 2. Doughnut Chart: 戶型佔比分佈
                if (doughnutChartRef.current) {
                    if (doughnutChartInstance.current) doughnutChartInstance.current.destroy();

                    const typeSummary = {};
                    dailyMetrics.rows.forEach(item => {
                        const typeLabel = item.type || '未填寫';
                        typeSummary[typeLabel] = (typeSummary[typeLabel] || 0) + 1;
                    });

                    const labels = Object.keys(typeSummary);
                    const dataValues = Object.values(typeSummary);

                    const ctx = doughnutChartRef.current.getContext('2d');
                    doughnutChartInstance.current = new Chart(ctx, {
                        type: 'doughnut',
                        data: {
                            labels: labels,
                            datasets: [{
                                data: dataValues,
                                backgroundColor: [
                                    '#991b1b', // 深紅
                                    '#dc2626', // 朱紅
                                    '#f87171', // 淺紅
                                    '#fca5a5', // 粉紅
                                    '#1e293b', // 鐵灰
                                    '#475569',
                                    '#94a3b8'
                                ],
                                borderWidth: 2,
                                borderColor: '#ffffff'
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: {
                                    position: 'right',
                                    labels: { boxWidth: 12, font: { size: 12 } }
                                }
                            },
                            cutout: '60%'
                        }
                    });
                }

            }, [dailyMetrics]);

            const filteredTableRows = useMemo(() => {
                if (!dailyMetrics) return [];
                if (!searchQuery.trim()) return dailyMetrics.rows;

                const q = searchQuery.toLowerCase().trim();
                return dailyMetrics.rows.filter(r => 
                    r.projectName.toLowerCase().includes(q) ||
                    r.type.toLowerCase().includes(q) ||
                    r.view.toLowerCase().includes(q) ||
                    r.identity.toLowerCase().includes(q) ||
                    r.region.toLowerCase().includes(q)
                );
            }, [dailyMetrics, searchQuery]);

            return (
                <div className="min-h-screen flex flex-col">
                    {/* Header bar */}
                    <header className="theme-bg-gradient text-white py-6 px-4 md:px-8 shadow-md">
                        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                            <div>
                                <div className="flex items-center gap-3">
                                    <span className="bg-white text-red-900 rounded-lg px-2.5 py-1 text-sm font-black tracking-wider">中海風格</span>
                                    <h1 className="text-2xl md:text-3xl font-bold tracking-tight">持份項目日報系統</h1>
                                </div>
                                <p className="text-red-100/80 text-sm mt-1.5 font-light">
                                    <i className="fa-solid fa-chart-line mr-1.5"></i> 
                                    秒級全自動智能數據解析與多維度指標診斷平台
                                </p>
                            </div>
                            <div className="text-xs md:text-sm text-right text-red-200">
                                預設解析 R欄(日期) / B欄(項目) / J欄(面積) / S欄(價錢)
                            </div>
                        </div>
                    </header>

                    {/* Main Content Area */}
                    <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-8 space-y-6">
                        
                        {}
                        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                            {/* Left: Upload Box */}
                            <div className="lg:col-span-1 bg-white p-6 rounded-2xl border border-red-100 card-shadow theme-border-highlight flex flex-col justify-between">
                                <div>
                                    <h2 className="text-lg font-bold text-red-950 mb-3 flex items-center gap-2">
                                        <i className="fa-solid fa-file-excel text-red-700"></i>
                                        上傳數據來源
                                    </h2>
                                    <p className="text-slate-500 text-xs mb-5 leading-relaxed">
                                        請直接拖入或選擇您要分析的交易報表 Excel 檔案。系統會優先搜尋名為 <strong>PFMQCHPANO</strong> 的分頁，若不存在則默認載入首個分頁。
                                    </p>
                                    
                                    <div className="relative group border-2 border-dashed border-red-200 hover:border-red-500 rounded-xl p-6 transition-colors bg-red-50/20 text-center cursor-pointer">
                                        <input 
                                            type="file" 
                                            accept=".xlsx, .xls"
                                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                            onChange={handleFileUpload} 
                                        />
                                        <div className="space-y-3">
                                            <div className="w-12 h-12 bg-red-50 text-red-600 rounded-full flex items-center justify-center mx-auto group-hover:scale-110 transition-transform">
                                                <i className="fa-solid fa-cloud-arrow-up text-xl"></i>
                                            </div>
                                            <div className="text-sm font-semibold text-red-900">選擇 Excel 檔案 (.xlsx)</div>
                                            <p className="text-xs text-slate-400">或直接將文件拖放至此區域</p>
                                        </div>
                                    </div>
                                </div>

                                {activeSheet && (
                                    <div className="mt-4 pt-4 border-t border-slate-100 flex items-center justify-between text-xs">
                                        <span className="text-slate-500">當前分頁：</span>
                                        <span className="font-bold text-red-800 bg-red-50 px-2 py-1 rounded">{activeSheet}</span>
                                    </div>
                                )}
                            </div>

                            {/* Right: Real-time Debug Logs Console */}
                            <div className="lg:col-span-2 bg-slate-900 text-slate-300 p-5 rounded-2xl shadow-xl flex flex-col h-[280px]">
                                <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-3">
                                    <div className="flex items-center gap-2">
                                        <span className="flex h-2.5 w-2.5 relative">
                                            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                                            <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500"></span>
                                        </span>
                                        <h3 className="text-sm font-bold text-white tracking-wide font-mono">EXCEL PARSER TERMINAL (系統解析終端機)</h3>
                                    </div>
                                    <button 
                                        onClick={() => setLogs([])}
                                        className="text-xs text-slate-500 hover:text-white transition-colors flex items-center gap-1"
                                        title="清除日誌"
                                    >
                                        <i className="fa-solid fa-trash-can"></i> 清除
                                    </button>
                                </div>
                                
                                <div className="flex-1 overflow-y-auto custom-scrollbar font-mono text-xs space-y-1.5 pr-2">
                                    {logs.length === 0 ? (
                                        <div className="text-slate-500 italic h-full flex items-center justify-center">
                                            等待上傳 Excel 檔案以開啟實時日誌監測...
                                        </div>
                                    ) : (
                                        logs.map((log, idx) => (
                                            <div key={idx} className={`leading-relaxed border-l-2 pl-2 py-0.5 ${
                                                log.type === 'success' ? 'text-emerald-400 border-emerald-500' :
                                                log.type === 'warning' ? 'text-amber-400 border-amber-500' :
                                                log.type === 'error' ? 'text-rose-400 border-rose-500 font-bold bg-rose-950/20' :
                                                'text-slate-300 border-slate-600'
                                            }`}>
                                                {log.text}
                                            </div>
                                        ))
                                    )}
                                </div>
                            </div>
                        </div>

                        {}
                        {rawData.length === 0 && (
                            <div className="bg-white border border-dashed border-slate-200 rounded-3xl p-12 text-center max-w-2xl mx-auto my-8">
                                <div className="w-16 h-16 bg-red-50 text-red-700 rounded-full flex items-center justify-center mx-auto mb-4">
                                    <i className="fa-solid fa-receipt text-2xl animate-pulse"></i>
                                </div>
                                <h3 className="text-lg font-bold text-slate-800 mb-2">尚未讀取到分析數據</h3>
                                <p className="text-slate-400 text-sm max-w-md mx-auto leading-relaxed">
                                    請利用上方上傳盒匯入您每日的成交報表。系統將自動從 R欄提取日期、B欄映射項目代號，為您繪製即時的可視化分析儀表板。
                                </p>
                            </div>
                        )}

                        {}
                        {rawData.length > 0 && (
                            <div className="space-y-6">
                                {/* Date Selector & Secondary Actions */}
                                <div className="bg-white p-4 rounded-xl border border-red-100 card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
                                    <div className="flex items-center gap-3">
                                        <span className="w-3 h-3 bg-red-700 rounded-full"></span>
                                        <label className="text-sm font-bold text-red-950">
                                            請選擇成交日期以載入分析：
                                        </label>
                                        <select 
                                            className="px-4 py-2 border border-red-200 rounded-lg bg-white font-medium text-red-950 focus:outline-none focus:ring-2 focus:ring-red-500 text-sm cursor-pointer"
                                            value={selectedDate} 
                                            onChange={(e) => setSelectedDate(e.target.value)}
                                        >
                                            {allDates.map(d => (
                                                <option key={d} value={d}>{d}</option>
                                            ))}
                                        </select>
                                    </div>
                                    <div className="text-xs text-slate-500 font-medium">
                                        <i className="fa-solid fa-database mr-1"></i> 全庫成交紀錄：{rawData.length} 筆
                                    </div>
                                </div>

                                {/* Summary Statistics Cards */}
                                {dailyMetrics && (
                                    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                                        {/* Sales Total */}
                                        <div className="bg-white p-6 rounded-2xl border border-red-100 card-shadow relative overflow-hidden group">
                                            <div className="absolute right-0 bottom-0 opacity-10 text-red-900 group-hover:scale-110 transition-transform">
                                                <i className="fa-solid fa-money-bill-trend-up text-8xl -mr-4 -mb-4"></i>
                                            </div>
                                            <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mb-2">當日總成交金額</p>
                                            <p className="text-3xl font-extrabold text-red-900">
                                                ${dailyMetrics.totalSales.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                                            </p>
                                            <span className="text-[10px] text-slate-400 block mt-2">
                                                當日共成交 {dailyMetrics.rows.length} 個單位
                                            </span>
                                        </div>

                                        {/* Total Area */}
                                        <div className="bg-white p-6 rounded-2xl border border-red-100 card-shadow relative overflow-hidden group">
                                            <div className="absolute right-0 bottom-0 opacity-10 text-red-900 group-hover:scale-110 transition-transform">
                                                <i className="fa-solid fa-chart-area text-8xl -mr-4 -mb-4"></i>
                                            </div>
                                            <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mb-2">當日總實用面積</p>
                                            <p className="text-3xl font-extrabold text-slate-800">
                                                {dailyMetrics.totalArea.toLocaleString()} <span className="text-lg font-normal">平方呎</span>
                                            </p>
                                            <span className="text-[10px] text-slate-400 block mt-2">
                                                所有成交物業實用面積加總
                                            </span>
                                        </div>

                                        {/* Avg Price / Sqft */}
                                        <div className="bg-white p-6 rounded-2xl border border-red-100 card-shadow relative overflow-hidden group">
                                            <div className="absolute right-0 bottom-0 opacity-10 text-red-900 group-hover:scale-110 transition-transform">
                                                <i className="fa-solid fa-calculator text-8xl -mr-4 -mb-4"></i>
                                            </div>
                                            <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mb-2">當日平均折實單價 (呎)</p>
                                            <p className="text-3xl font-extrabold text-red-700">
                                                ${Math.round(dailyMetrics.avgPricePerSqft).toLocaleString()} <span className="text-sm font-normal text-slate-500">/呎</span>
                                            </p>
                                            <span className="text-[10px] text-red-900/60 block mt-2 font-medium bg-red-50 py-0.5 px-2 rounded inline-block">
                                                算法：總成交金額 ÷ 總實用面積
                                            </span>
                                        </div>
                                    </div>
                                )}

                                {/* Charts Group (Two Columns) */}
                                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                                    {/* Bar Chart */}
                                    <div className="bg-white p-6 rounded-2xl border border-red-100 card-shadow">
                                        <h3 className="text-base font-bold text-red-950 mb-1 flex items-center gap-2">
                                            <span className="w-1 h-4 bg-red-700 rounded-full"></span>
                                            當日項目成交金額分佈
                                        </h3>
                                        <p className="text-xs text-slate-400 mb-4">分析各持份項目在當天的具體業績表現</p>
                                        <div className="relative h-[320px]">
                                            <canvas ref={barChartRef}></canvas>
                                        </div>
                                    </div>

                                    {/* Doughnut Chart */}
                                    <div className="bg-white p-6 rounded-2xl border border-red-100 card-shadow">
                                        <h3 className="text-base font-bold text-red-950 mb-1 flex items-center gap-2">
                                            <span className="w-1 h-4 bg-red-700 rounded-full"></span>
                                            當日成交戶型佔比分佈
                                        </h3>
                                        <p className="text-xs text-slate-400 mb-4">掌握當天買家偏好的主力購買戶型</p>
                                        <div className="relative h-[320px] flex items-center justify-center">
                                            <canvas ref={doughnutChartRef}></canvas>
                                        </div>
                                    </div>
                                </div>

                                {}
                                <div className="bg-white rounded-2xl border border-red-100 card-shadow overflow-hidden">
                                    {/* Table Header Controls */}
                                    <div className="p-5 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-50/50">
                                        <div>
                                            <h3 className="text-base font-bold text-red-950 flex items-center gap-2">
                                                <i className="fa-solid fa-list-check text-red-700"></i>
                                                當日成交明細名冊
                                            </h3>
                                            <p className="text-xs text-slate-400 mt-0.5">顯示已選擇成交日期的所有原始記錄</p>
                                        </div>
                                        <div className="relative w-full md:w-80">
                                            <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400 text-xs">
                                                <i className="fa-solid fa-magnifying-glass"></i>
                                            </span>
                                            <input 
                                                type="text" 
                                                placeholder="快速搜尋項目/戶型/景觀/身份..." 
                                                className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500 bg-white"
                                                value={searchQuery}
                                                onChange={(e) => setSearchQuery(e.target.value)}
                                            />
                                        </div>
                                    </div>

                                    {/* Table body */}
                                    <div className="overflow-x-auto custom-scrollbar">
                                        <table className="w-full text-left border-collapse">
                                            <thead>
                                                <tr className="border-b border-slate-100 text-slate-500 text-xs font-bold bg-slate-50/80">
                                                    <th className="py-3 px-4 text-center">列</th>
                                                    <th className="py-3 px-4">項目名稱</th>
                                                    <th className="py-3 px-4">戶型</th>
                                                    <th className="py-3 px-4">景觀</th>
                                                    <th className="py-3 px-4 text-right">實用面積(呎)</th>
                                                    <th className="py-3 px-4 text-right">折實成交價</th>
                                                    <th className="py-3 px-4 text-right">折實呎價</th>
                                                    <th className="py-3 px-4">買家身份</th>
                                                    <th className="py-3 px-4">年齡</th>
                                                    <th className="py-3 px-4">居住區域</th>
                                                    <th className="py-3 px-4">用途代理</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-100 text-xs text-slate-600">
                                                {filteredTableRows.length === 0 ? (
                                                    <tr>
                                                        <td colSpan="11" className="py-8 text-center text-slate-400 italic">
                                                            無符合當前搜尋篩選的交易明細資料。
                                                        </td>
                                                    </tr>
                                                ) : (
                                                    filteredTableRows.map((row, idx) => (
                                                        <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                                                            <td className="py-3 px-4 text-center font-mono text-slate-400">{row.rawIndex}</td>
                                                            <td className="py-3 px-4 font-bold text-slate-900">{row.projectName}</td>
                                                            <td className="py-3 px-4"><span className="bg-slate-100 px-2 py-0.5 rounded text-[10px] text-slate-600 font-semibold">{row.type || '未填寫'}</span></td>
                                                            <td className="py-3 px-4">{row.view || '-'}</td>
                                                            <td className="py-3 px-4 text-right font-mono font-medium">{row.area.toLocaleString()}</td>
                                                            <td className="py-3 px-4 text-right font-mono font-bold text-red-900">${row.price.toLocaleString()}</td>
                                                            <td className="py-3 px-4 text-right font-mono text-emerald-800 font-semibold">${row.unitPrice.toLocaleString()}/呎</td>
                                                            <td className="py-3 px-4">{row.identity || '-'}</td>
                                                            <td className="py-3 px-4">{row.age || '-'}</td>
                                                            <td className="py-3 px-4">{row.region || '-'}</td>
                                                            <td className="py-3 px-4">{row.agent || '-'}</td>
                                                        </tr>
                                                    ))
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        )}
                    </main>

                    {/* Footer bar */}
                    <footer className="border-t border-slate-100 bg-white py-6 mt-12 text-center text-xs text-slate-400">
                        <div className="max-w-7xl mx-auto px-4">
                            <p>© 2026 持份項目日報系統 - China Overseas Style Platform. All rights reserved.</p>
                            <p className="mt-1 text-[10px] text-slate-300">本系統完全在瀏覽器本地安全解析 Excel 數據，您的營業數據絕不會上傳到任何外部伺服器。</p>
                        </div>
                    </footer>
                </div>
            );
        };

        const root = ReactDOM.createRoot(document.getElementById('root'));
        root.render(<App />);
    </script>
</body>
</html>

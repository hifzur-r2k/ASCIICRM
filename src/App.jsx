import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import { Upload, Download, FileText, UploadCloud, LayoutGrid, Search } from 'lucide-react';

export default function App() {
  const [siteData, setSiteData] = useState([]);
  const [dayData, setDayData] = useState([]);
  const [activeTab, setActiveTab] = useState('day');
  const [sheetName, setSheetName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const fileInputRef = useRef(null);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      let rawData = [];

      const currentSheetName = wb.SheetNames[0];
      setSheetName(currentSheetName);
      const sheet = wb.Sheets[currentSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      let activeSection = "Helper"; 
      let dayHeaders = {};

      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;

        const col0 = String(row[0]).trim().toUpperCase();
        const col1 = String(row[1]).trim().toUpperCase();

        if (col0.includes("HELPER") || col1.includes("HELPER") || col0.includes("LABOUR") || col1.includes("LABOUR")) {
          activeSection = "Helper";
        } else if ((col0.includes("MASON") || col1.includes("MASON")) && !col0.includes("HALF") && !col1.includes("HALF")) {
          activeSection = "Mason";
        }
        
        let rowWorkerType = activeSection;

        if (/\bHM\b/.test(col0) || /\bHM\b/.test(col1) || col1.includes("HALF MASON") || col0.includes("HALF MASON")) {
          rowWorkerType = "HalfMason";
        }

        const col0Clean = col0.replace(/[^A-Z]/g, '');
        if (col0Clean === "SN" || col0Clean === "SNO" || col0 === "S.N.") {
          let tempHeaders = {};
          for (let c = 2; c < row.length; c++) {
            const val = String(row[c]).trim();
            if (/^\d+$/.test(val)) {
              tempHeaders[c] = val;
            }
          }
          if (Object.keys(tempHeaders).length > 0) {
            dayHeaders = tempHeaders;
          }
          continue;
        }

        if (/^\d+$/.test(col0) && col1 !== "" && col1 !== "OT" && col1 !== "NAN") {
          const nextRow = rows[r + 1] || [];
          const isOTRow = String(nextRow[1]).trim().toUpperCase() === "OT";

          for (let c in dayHeaders) {
            const dayNum = parseInt(dayHeaders[c]);
            const cellVal = String(row[c]).trim().toUpperCase();
            
            if (cellVal && cellVal !== "NAN") {
              const match = cellVal.match(/^([\d.]+)\s*([A-Z].*)$/i);
              
              if (match) {
                const regDays = parseFloat(match[1]);
                const site = match[2].trim();

                if (regDays > 0 && site !== "NA") {
                  let otHrs = 0;
                  if (isOTRow && nextRow[c]) {
                    const otVal = parseFloat(String(nextRow[c]).trim());
                    if (!isNaN(otVal)) otHrs = otVal;
                  }

                  rawData.push({ day: dayNum, site, type: rowWorkerType, reg: regDays, ot: otHrs });
                }
              }
            }
          }
        }
      }

      const sMap = {};
      rawData.forEach(r => {
        if (!sMap[r.site]) sMap[r.site] = { site: r.site, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0 };
        if (r.type === 'Mason') {
          sMap[r.site].masonReg += r.reg;
          sMap[r.site].masonOT += r.ot;
        } else if (r.type === 'HalfMason') {
          sMap[r.site].halfMasonReg += r.reg;
          sMap[r.site].halfMasonOT += r.ot;
        } else {
          sMap[r.site].helperReg += r.reg;
          sMap[r.site].helperOT += r.ot;
        }
      });
      setSiteData(Object.values(sMap).sort((a, b) => a.site.localeCompare(b.site)));

      const dMap = {};
      rawData.forEach(r => {
        const key = `${r.day}|${r.site}`;
        if (!dMap[key]) dMap[key] = { day: r.day, site: r.site, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0 };
        if (r.type === 'Mason') {
          dMap[key].masonReg += r.reg;
          dMap[key].masonOT += r.ot;
        } else if (r.type === 'HalfMason') {
          dMap[key].halfMasonReg += r.reg;
          dMap[key].halfMasonOT += r.ot;
        } else {
          dMap[key].helperReg += r.reg;
          dMap[key].helperOT += r.ot;
        }
      });
      setDayData(Object.values(dMap).sort((a, b) => a.day - b.day || a.site.localeCompare(b.site)));
    };
    reader.readAsBinaryString(file);
  };

  // 1. Data Selection & Search Filtering
  const activeData = activeTab === 'site' ? siteData : dayData;
  const filteredData = activeData.filter(row => 
    row.site.toLowerCase().includes(searchQuery.toLowerCase())
  );

  // 2. Calculate Totals dynamically based on what is filtered
  const totals = filteredData.reduce((acc, row) => {
    acc.masonReg += row.masonReg || 0;
    acc.masonOT += row.masonOT || 0;
    acc.halfMasonReg += row.halfMasonReg || 0;
    acc.halfMasonOT += row.halfMasonOT || 0;
    acc.helperReg += row.helperReg || 0;
    acc.helperOT += row.helperOT || 0;
    return acc;
  }, { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0 });


  const exportToExcel = () => {
    const formattedData = filteredData.map(row => {
      const base = {};
      if (activeTab === 'day') base["Date (Day)"] = row.day;
      return {
        ...base,
        "Site Code": row.site,
        "Mason Days": row.masonReg,
        "Mason Extra Hours": row.masonOT,
        "Half Mason Days": row.halfMasonReg,
        "Half Mason Extra": row.halfMasonOT,
        "Helper Days": row.helperReg,
        "Helper Extra Hours": row.helperOT
      };
    });

    // Append Total Row to Excel
    const totalRow = { "Site Code": "GRAND TOTAL" };
    if (activeTab === 'day') totalRow["Date (Day)"] = "";
    totalRow["Mason Days"] = totals.masonReg;
    totalRow["Mason Extra Hours"] = totals.masonOT;
    totalRow["Half Mason Days"] = totals.halfMasonReg;
    totalRow["Half Mason Extra"] = totals.halfMasonOT;
    totalRow["Helper Days"] = totals.helperReg;
    totalRow["Helper Extra Hours"] = totals.helperOT;
    formattedData.push(totalRow);

    const ws = XLSX.utils.json_to_sheet(formattedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance Data");
    XLSX.writeFile(wb, `CRM_FIX_${sheetName}_${activeTab === 'day' ? 'DayWise' : 'SiteWise'}.xlsx`);
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    const viewTitle = activeTab === 'day' ? 'Day-Wise & Site-Wise Breakdown' : 'Total Site-Wise Summary';
    
    doc.setFontSize(16);
    doc.setTextColor(37, 99, 235);
    doc.text("CRM_FIX", 14, 15);
    
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Sheet: ${sheetName} | ${viewTitle}`, 14, 22);
    
    const tableColumn = activeTab === 'day' 
      ? ["Day", "Site", "Mason D", "Mason Ex", "HM Days", "HM Ex", "Helper D", "Helper Ex"]
      : ["Site Code", "Mason Days", "Mason Extra", "HM Days", "HM Extra", "Helper Days", "Helper Extra"];

    const tableRows = filteredData.map(row => 
      activeTab === 'day'
        ? [row.day, row.site, row.masonReg, row.masonOT, row.halfMasonReg, row.halfMasonOT, row.helperReg, row.helperOT]
        : [row.site, row.masonReg, row.masonOT, row.halfMasonReg, row.halfMasonOT, row.helperReg, row.helperOT]
    );

    // Append Total Row to PDF Footer
    const footRow = activeTab === 'day'
      ? ["", "TOTAL", totals.masonReg, totals.masonOT, totals.halfMasonReg, totals.halfMasonOT, totals.helperReg, totals.helperOT]
      : ["TOTAL", totals.masonReg, totals.masonOT, totals.halfMasonReg, totals.halfMasonOT, totals.helperReg, totals.helperOT];

    doc.autoTable({
      head: [tableColumn],
      body: tableRows,
      foot: [footRow],
      startY: 28,
      theme: 'grid',
      headStyles: { fillColor: [37, 99, 235] },
      footStyles: { fillColor: [243, 244, 246], textColor: [17, 24, 39], fontStyle: 'bold' }
    });
    
    doc.save(`CRM_FIX_${sheetName}_${activeTab === 'day' ? 'DayWise' : 'SiteWise'}.pdf`);
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] text-gray-800 font-sans selection:bg-blue-100 selection:text-blue-900">
      <input type="file" accept=".xlsx, .xls, .csv" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
      
      {/* --- EMPTY STATE (BEFORE UPLOAD) --- */}
      {activeData.length === 0 ? (
        <div className="flex flex-col items-center justify-center min-h-screen px-4">
          <div className="text-center mb-10 space-y-2">
            <div className="inline-flex items-center justify-center p-3 bg-blue-600 rounded-2xl shadow-lg shadow-blue-600/20 mb-4">
              <LayoutGrid className="w-8 h-8 text-white" />
            </div>
            <h1 className="text-4xl md:text-5xl font-black tracking-tight text-gray-900">
              CRM_FIX
            </h1>
            <p className="text-lg font-medium text-gray-500 uppercase tracking-widest">
              Day Wise - Site wise Evaluation
            </p>
          </div>

          <div 
            onClick={() => fileInputRef.current?.click()}
            className="group w-full max-w-2xl cursor-pointer bg-white rounded-3xl border-2 border-dashed border-gray-300 hover:border-blue-500 hover:bg-blue-50/50 transition-all duration-300 p-12 md:p-20 shadow-sm hover:shadow-xl text-center"
          >
            <div className="flex flex-col items-center justify-center space-y-6">
              <div className="bg-blue-100/50 text-blue-600 p-6 rounded-full group-hover:scale-110 group-hover:bg-blue-100 transition-all duration-300">
                <UploadCloud className="w-12 h-12" />
              </div>
              <div>
                <p className="text-2xl font-bold text-gray-800">Upload Attendance Sheet</p>
                <p className="text-gray-500 mt-2">Upload a single sheet for 100% accurate daily/site parsing.</p>
              </div>
              <span className="mt-4 px-8 py-3 bg-gray-900 text-white font-semibold rounded-xl shadow-md group-hover:bg-blue-600 transition-colors duration-300 flex items-center gap-2">
                <Upload className="w-4 h-4" />
                Select File
              </span>
            </div>
          </div>
        </div>
      ) : (

      /* --- DATA VIEW (AFTER UPLOAD) --- */
        <div className="max-w-7xl mx-auto space-y-6 p-4 md:p-8 pt-8">
          
          <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200 flex flex-col md:flex-row justify-between items-center gap-4">
            <div className="flex items-center gap-4">
              <div className="bg-blue-600 p-2.5 rounded-xl text-white shadow-sm">
                <LayoutGrid className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-2xl font-black tracking-tight text-gray-900 leading-none">
                  CRM_FIX
                </h1>
                <p className="text-xs font-bold uppercase tracking-wider text-gray-400 mt-1">Day Wise - Site wise Evaluation</p>
              </div>
            </div>
            
            <button 
              onClick={() => {
                setSearchQuery("");
                fileInputRef.current?.click();
              }}
              className="w-full md:w-auto justify-center bg-gray-100 hover:bg-gray-200 text-gray-700 px-5 py-2.5 rounded-xl font-semibold flex items-center gap-2 transition-all"
            >
              <Upload className="w-4 h-4" />
              Upload New Sheet
            </button>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-gray-200 overflow-hidden">
            
            <div className="flex flex-col md:flex-row justify-between items-center border-b border-gray-100 bg-gray-50/50 p-3 gap-3">
              
              <div className="flex w-full md:w-auto bg-gray-200/60 p-1 rounded-xl">
                <button 
                  className={`flex-1 md:flex-none px-6 py-2.5 font-bold text-sm rounded-lg transition-all ${activeTab === 'day' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                  onClick={() => setActiveTab('day')}
                >
                  Day-Wise View
                </button>
                <button 
                  className={`flex-1 md:flex-none px-6 py-2.5 font-bold text-sm rounded-lg transition-all ${activeTab === 'site' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                  onClick={() => setActiveTab('site')}
                >
                  Site-Wise View
                </button>
              </div>

              {/* SEARCH BAR */}
              <div className="relative w-full md:w-72">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input
                  type="text"
                  placeholder="Search site code..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-9 pr-4 py-2.5 bg-white border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all shadow-sm"
                />
              </div>

              <div className="flex w-full md:w-auto gap-2">
                <button 
                  onClick={exportToExcel}
                  className="flex-1 md:flex-none justify-center text-sm font-semibold bg-emerald-50 text-emerald-700 hover:bg-emerald-100 hover:text-emerald-800 px-4 py-2.5 rounded-xl flex items-center gap-2 transition-colors border border-emerald-200"
                >
                  <Download className="w-4 h-4" />
                  Excel
                </button>
                <button 
                  onClick={exportToPDF}
                  className="flex-1 md:flex-none justify-center text-sm font-semibold bg-rose-50 text-rose-700 hover:bg-rose-100 hover:text-rose-800 px-4 py-2.5 rounded-xl flex items-center gap-2 transition-colors border border-rose-200"
                >
                  <FileText className="w-4 h-4" />
                  PDF
                </button>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[650px] bg-gray-50 md:bg-white custom-scrollbar relative">
              
              {/* Desktop Table */}
              <table className="w-full text-sm text-left hidden md:table">
                <thead className="bg-gray-50/90 uppercase text-[11px] font-black tracking-wider text-gray-500 sticky top-0 z-10 backdrop-blur-md shadow-sm">
                  <tr>
                    {activeTab === 'day' && <th className="px-6 py-4 border-b border-r border-gray-100">Day</th>}
                    <th className="px-6 py-4 border-b border-r border-gray-100">Site Code</th>
                    <th className="px-6 py-4 border-b border-r border-gray-100 text-blue-800 bg-blue-50/30">Mason Days</th>
                    <th className="px-6 py-4 border-b border-r border-gray-100 text-blue-800 bg-blue-50/30">Mason Extra</th>
                    <th className="px-6 py-4 border-b border-r border-gray-100 text-purple-800 bg-purple-50/30">Half Mason Days</th>
                    <th className="px-6 py-4 border-b border-r border-gray-100 text-purple-800 bg-purple-50/30">Half Mason Extra</th>
                    <th className="px-6 py-4 border-b border-r border-gray-100 text-orange-800 bg-orange-50/30">Helper Days</th>
                    <th className="px-6 py-4 border-b border-gray-100 text-orange-800 bg-orange-50/30">Helper Extra</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.length > 0 ? (
                    filteredData.map((row, idx) => (
                      <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/80 bg-white transition-colors">
                        {activeTab === 'day' && <td className="px-6 py-4 border-r border-gray-50 font-bold text-gray-900">{row.day}</td>}
                        <td className="px-6 py-4 border-r border-gray-50 font-bold text-gray-700">{row.site}</td>
                        <td className="px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.masonReg}</td>
                        <td className="px-6 py-4 border-r border-gray-50 text-gray-500">{row.masonOT}</td>
                        <td className="px-6 py-4 border-r border-gray-50 font-bold text-purple-800">{row.halfMasonReg}</td>
                        <td className="px-6 py-4 border-r border-gray-50 text-purple-600">{row.halfMasonOT}</td>
                        <td className="px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.helperReg}</td>
                        <td className="px-6 py-4 text-gray-500">{row.helperOT}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="8" className="px-6 py-12 text-center text-gray-500 font-medium">
                        No sites found matching "{searchQuery}"
                      </td>
                    </tr>
                  )}
                </tbody>
                
                {/* DESKTOP TOTALS ROW */}
                {filteredData.length > 0 && (
                  <tfoot className="bg-gray-100/90 sticky bottom-0 z-10 backdrop-blur-md shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.05)] border-t-2 border-gray-200">
                    <tr>
                      <td colSpan={activeTab === 'day' ? 2 : 1} className="px-6 py-4 text-right font-black text-gray-900">GRAND TOTAL:</td>
                      <td className="px-6 py-4 font-black text-blue-900 text-base">{totals.masonReg}</td>
                      <td className="px-6 py-4 font-bold text-blue-700">{totals.masonOT}</td>
                      <td className="px-6 py-4 font-black text-purple-900 text-base">{totals.halfMasonReg}</td>
                      <td className="px-6 py-4 font-bold text-purple-700">{totals.halfMasonOT}</td>
                      <td className="px-6 py-4 font-black text-orange-900 text-base">{totals.helperReg}</td>
                      <td className="px-6 py-4 font-bold text-orange-700">{totals.helperOT}</td>
                    </tr>
                  </tfoot>
                )}
              </table>

              {/* Mobile Cards */}
              <div className="block md:hidden p-4 space-y-4">
                {filteredData.length > 0 ? (
                  filteredData.map((row, idx) => (
                    <div key={idx} className="bg-white p-5 rounded-2xl shadow-sm border border-gray-100 space-y-4">
                      <div className="flex justify-between items-center border-b border-gray-50 pb-3">
                        <span className="text-sm font-black text-gray-800 bg-gray-100 px-3 py-1.5 rounded-lg">
                          Site: {row.site}
                        </span>
                        {activeTab === 'day' && (
                          <span className="text-xs font-bold text-blue-700 bg-blue-50 px-3 py-1.5 rounded-lg">
                            Day {row.day}
                          </span>
                        )}
                      </div>
                      
                      <div className="grid grid-cols-2 gap-3 text-xs">
                        <div className="bg-blue-50/50 p-3 rounded-xl border border-blue-100">
                          <p className="text-blue-600/80 font-bold mb-1 uppercase tracking-wider text-[10px]">Mason Days</p>
                          <p className="text-lg font-black text-blue-900">{row.masonReg}</p>
                        </div>
                        <div className="bg-blue-50/50 p-3 rounded-xl border border-blue-100">
                          <p className="text-blue-600/80 font-bold mb-1 uppercase tracking-wider text-[10px]">Mason Extra</p>
                          <p className="text-lg font-black text-blue-900">{row.masonOT} <span className="text-xs font-medium text-blue-400">hrs</span></p>
                        </div>

                        <div className="bg-purple-50/50 p-3 rounded-xl border border-purple-100">
                          <p className="text-purple-600/80 font-bold mb-1 uppercase tracking-wider text-[10px]">HM Days</p>
                          <p className="text-lg font-black text-purple-900">{row.halfMasonReg}</p>
                        </div>
                        <div className="bg-purple-50/50 p-3 rounded-xl border border-purple-100">
                          <p className="text-purple-600/80 font-bold mb-1 uppercase tracking-wider text-[10px]">HM Extra</p>
                          <p className="text-lg font-black text-purple-900">{row.halfMasonOT} <span className="text-xs font-medium text-purple-400">hrs</span></p>
                        </div>

                        <div className="bg-orange-50/50 p-3 rounded-xl border border-orange-100">
                          <p className="text-orange-600/80 font-bold mb-1 uppercase tracking-wider text-[10px]">Helper Days</p>
                          <p className="text-lg font-black text-orange-900">{row.helperReg}</p>
                        </div>
                        <div className="bg-orange-50/50 p-3 rounded-xl border border-orange-100">
                          <p className="text-orange-600/80 font-bold mb-1 uppercase tracking-wider text-[10px]">Helper Extra</p>
                          <p className="text-lg font-black text-orange-900">{row.helperOT} <span className="text-xs font-medium text-orange-400">hrs</span></p>
                        </div>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="text-center py-12 text-gray-500 font-medium bg-white rounded-2xl border border-gray-100">
                    No sites found matching "{searchQuery}"
                  </div>
                )}

                {/* MOBILE TOTALS CARD */}
                {filteredData.length > 0 && (
                  <div className="bg-gray-900 p-6 rounded-2xl shadow-xl mt-6 border border-gray-800">
                    <h3 className="text-white font-black text-center text-lg mb-4 tracking-wider">GRAND TOTAL</h3>
                    <div className="grid grid-cols-2 gap-3 text-xs">
                      <div className="bg-gray-800/80 p-3 rounded-xl">
                        <p className="text-gray-400 font-bold mb-1 uppercase text-[10px]">Total Mason Days</p>
                        <p className="text-xl font-black text-white">{totals.masonReg}</p>
                      </div>
                      <div className="bg-gray-800/80 p-3 rounded-xl">
                        <p className="text-gray-400 font-bold mb-1 uppercase text-[10px]">Total Mason Extra</p>
                        <p className="text-xl font-black text-white">{totals.masonOT} <span className="text-xs font-medium text-gray-500">hrs</span></p>
                      </div>
                      <div className="bg-gray-800/80 p-3 rounded-xl">
                        <p className="text-purple-400 font-bold mb-1 uppercase text-[10px]">Total HM Days</p>
                        <p className="text-xl font-black text-purple-200">{totals.halfMasonReg}</p>
                      </div>
                      <div className="bg-gray-800/80 p-3 rounded-xl">
                        <p className="text-purple-400 font-bold mb-1 uppercase text-[10px]">Total HM Extra</p>
                        <p className="text-xl font-black text-purple-200">{totals.halfMasonOT} <span className="text-xs font-medium text-gray-500">hrs</span></p>
                      </div>
                      <div className="bg-gray-800/80 p-3 rounded-xl">
                        <p className="text-gray-400 font-bold mb-1 uppercase text-[10px]">Total Helper Days</p>
                        <p className="text-xl font-black text-white">{totals.helperReg}</p>
                      </div>
                      <div className="bg-gray-800/80 p-3 rounded-xl">
                        <p className="text-gray-400 font-bold mb-1 uppercase text-[10px]">Total Helper Extra</p>
                        <p className="text-xl font-black text-white">{totals.helperOT} <span className="text-xs font-medium text-gray-500">hrs</span></p>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
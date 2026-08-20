import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import { Upload, HardHat, Download, FileText } from 'lucide-react';

export default function App() {
  const [siteData, setSiteData] = useState([]);
  const [dayData, setDayData] = useState([]);
  const [activeTab, setActiveTab] = useState('day');
  const [sheetName, setSheetName] = useState("");
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

      let workerType = "Helper";
      let dayHeaders = {};

      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;

        const col0 = String(row[0]).trim().toUpperCase();
        const col1 = String(row[1]).trim().toUpperCase();

        if (col0.includes("HELPER") || col1.includes("HELPER") || col0.includes("LABOUR") || col1.includes("LABOUR")) {
          workerType = "Helper";
        }
        if (col0.includes("MASON") || col1.includes("MASON")) {
          workerType = "Mason";
        }

        const col0Clean = col0.replace(/[^A-Z]/g, '');
        if (col0Clean === "SN" || col0Clean === "SNO" || col0 === "S.N.") {
          dayHeaders = {};
          for (let c = 2; c < row.length; c++) {
            const val = String(row[c]).trim();
            if (/^\d+$/.test(val)) {
              dayHeaders[c] = val;
            }
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

                  rawData.push({ day: dayNum, site, type: workerType, reg: regDays, ot: otHrs });
                }
              }
            }
          }
        }
      }

      const sMap = {};
      rawData.forEach(r => {
        if (!sMap[r.site]) sMap[r.site] = { site: r.site, masonReg: 0, masonOT: 0, helperReg: 0, helperOT: 0 };
        if (r.type === 'Mason') {
          sMap[r.site].masonReg += r.reg;
          sMap[r.site].masonOT += r.ot;
        } else {
          sMap[r.site].helperReg += r.reg;
          sMap[r.site].helperOT += r.ot;
        }
      });
      setSiteData(Object.values(sMap).sort((a, b) => a.site.localeCompare(b.site)));

      const dMap = {};
      rawData.forEach(r => {
        const key = `${r.day}|${r.site}`;
        if (!dMap[key]) dMap[key] = { day: r.day, site: r.site, masonReg: 0, masonOT: 0, helperReg: 0, helperOT: 0 };
        if (r.type === 'Mason') {
          dMap[key].masonReg += r.reg;
          dMap[key].masonOT += r.ot;
        } else {
          dMap[key].helperReg += r.reg;
          dMap[key].helperOT += r.ot;
        }
      });
      setDayData(Object.values(dMap).sort((a, b) => a.day - b.day || a.site.localeCompare(b.site)));
    };
    reader.readAsBinaryString(file);
  };

  const exportToExcel = () => {
    const dataToExport = activeTab === 'site' ? siteData : dayData;
    
    const formattedData = dataToExport.map(row => {
      const base = {};
      if (activeTab === 'day') base["Date (Day)"] = row.day;
      return {
        ...base,
        "Site Code": row.site,
        "Mason Days": row.masonReg,
        "Mason Extra Hours": row.masonOT,
        "Helper Days": row.helperReg,
        "Helper Extra Hours": row.helperOT
      };
    });

    const ws = XLSX.utils.json_to_sheet(formattedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance Data");
    XLSX.writeFile(wb, `ContecBuildFlow_${sheetName}_${activeTab === 'day' ? 'DayWise' : 'SiteWise'}.xlsx`);
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    const viewTitle = activeTab === 'day' ? 'Day-Wise & Site-Wise Breakdown' : 'Total Site-Wise Summary';
    
    doc.setFontSize(16);
    doc.setTextColor(37, 99, 235);
    doc.text("ContecBuildFlow", 14, 15);
    
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Contec Solutions | Sheet: ${sheetName} | ${viewTitle}`, 14, 22);
    
    const dataToExport = activeTab === 'site' ? siteData : dayData;
    
    const tableColumn = activeTab === 'day' 
      ? ["Day", "Site Code", "Mason Days", "Mason Extra", "Helper Days", "Helper Extra"]
      : ["Site Code", "Mason Days", "Mason Extra", "Helper Days", "Helper Extra"];

    const tableRows = dataToExport.map(row => 
      activeTab === 'day'
        ? [row.day, row.site, row.masonReg, row.masonOT, row.helperReg, row.helperOT]
        : [row.site, row.masonReg, row.masonOT, row.helperReg, row.helperOT]
    );

    doc.autoTable({
      head: [tableColumn],
      body: tableRows,
      startY: 28,
      theme: 'grid',
      headStyles: { fillColor: [37, 99, 235] }
    });
    
    doc.save(`ContecBuildFlow_${sheetName}_${activeTab === 'day' ? 'DayWise' : 'SiteWise'}.pdf`);
  };

  const displayData = activeTab === 'site' ? siteData : dayData;

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-8 text-gray-800">
      <div className="max-w-5xl mx-auto space-y-6">
        
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <div className="flex items-center gap-2">
              <div className="bg-blue-600 p-2 rounded-lg text-white">
                <HardHat className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-2xl font-black tracking-tight text-gray-900">
                  CRM_FIX
                </h1>
                <p className="text-xs font-semibold uppercase tracking-wider text-gray-400">Day Wise - Site wise Evaluation</p>
              </div>
            </div>
            <p className="text-gray-500 text-sm mt-2">Upload a single sheet for 100% accurate daily/site parsing.</p>
          </div>
          
          <input type="file" accept=".xlsx, .xls, .csv" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="w-full md:w-auto justify-center bg-blue-600 hover:bg-blue-700 text-white px-6 py-2.5 rounded-lg font-medium flex items-center gap-2 shadow-sm transition-all"
          >
            <Upload className="w-5 h-5" />
            Upload Sheet
          </button>
        </div>

        {displayData.length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            
            <div className="flex flex-col sm:flex-row justify-between items-center border-b border-gray-200 bg-gray-50">
              <div className="flex w-full sm:w-auto border-r border-gray-200">
                <button 
                  className={`flex-1 sm:flex-none px-6 py-4 font-semibold text-sm transition-colors ${activeTab === 'day' ? 'text-blue-600 border-b-2 border-blue-600 bg-white' : 'text-gray-500 hover:bg-gray-100'}`}
                  onClick={() => setActiveTab('day')}
                >
                  Day-Wise View
                </button>
                <button 
                  className={`flex-1 sm:flex-none px-6 py-4 font-semibold text-sm transition-colors ${activeTab === 'site' ? 'text-blue-600 border-b-2 border-blue-600 bg-white' : 'text-gray-500 hover:bg-gray-100'}`}
                  onClick={() => setActiveTab('site')}
                >
                  Site-Wise View
                </button>
              </div>

              <div className="p-3 w-full sm:w-auto flex gap-2 justify-end bg-gray-50">
                <button 
                  onClick={exportToExcel}
                  className="flex-1 sm:flex-none justify-center text-sm bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-md flex items-center gap-2 transition-colors shadow-sm"
                >
                  <Download className="w-4 h-4" />
                  Excel
                </button>
                <button 
                  onClick={exportToPDF}
                  className="flex-1 sm:flex-none justify-center text-sm bg-rose-600 hover:bg-rose-700 text-white px-4 py-2 rounded-md flex items-center gap-2 transition-colors shadow-sm"
                >
                  <FileText className="w-4 h-4" />
                  PDF
                </button>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[600px] bg-gray-50 md:bg-white">
              
              {/* Desktop Table View */}
              <table className="w-full text-sm text-left hidden md:table">
                <thead className="bg-gray-100 uppercase text-xs text-gray-700 sticky top-0 z-10">
                  <tr>
                    {activeTab === 'day' && <th className="px-6 py-3 border-r bg-gray-100">Day</th>}
                    <th className="px-6 py-3 border-r bg-gray-100">Site Code</th>
                    <th className="px-6 py-3 border-r bg-blue-50/70 text-blue-900">Mason Days</th>
                    <th className="px-6 py-3 border-r bg-blue-50/70 text-blue-900">Mason Extra</th>
                    <th className="px-6 py-3 border-r bg-orange-50/70 text-orange-900">Helper Days</th>
                    <th className="px-6 py-3 bg-orange-50/70 text-orange-900">Helper Extra</th>
                  </tr>
                </thead>
                <tbody>
                  {displayData.map((row, idx) => (
                    <tr key={idx} className="border-b hover:bg-gray-50 bg-white">
                      {activeTab === 'day' && <td className="px-6 py-3 border-r font-bold text-gray-900">{row.day}</td>}
                      <td className="px-6 py-3 border-r font-bold text-gray-800">{row.site}</td>
                      <td className="px-6 py-3 border-r">{row.masonReg}</td>
                      <td className="px-6 py-3 border-r">{row.masonOT}</td>
                      <td className="px-6 py-3 border-r">{row.helperReg}</td>
                      <td className="px-6 py-3">{row.helperOT}</td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Mobile Card View */}
              <div className="block md:hidden p-4 space-y-4">
                {displayData.map((row, idx) => (
                  <div key={idx} className="bg-white p-4 rounded-xl shadow-sm border border-gray-200 space-y-3">
                    <div className="flex justify-between items-center border-b border-gray-100 pb-2">
                      <span className="text-sm font-black text-gray-800 bg-gray-100 px-3 py-1 rounded-md">
                        Site: {row.site}
                      </span>
                      {activeTab === 'day' && (
                        <span className="text-xs font-bold text-blue-700 bg-blue-100 px-3 py-1 rounded-md">
                          Day {row.day}
                        </span>
                      )}
                    </div>
                    
                    <div className="grid grid-cols-2 gap-3 text-xs">
                      <div className="bg-blue-50/50 p-2.5 rounded-lg border border-blue-100">
                        <p className="text-blue-600/80 font-semibold mb-1 uppercase tracking-wider text-[10px]">Mason Days</p>
                        <p className="text-base font-black text-blue-900">{row.masonReg}</p>
                      </div>
                      <div className="bg-blue-50/50 p-2.5 rounded-lg border border-blue-100">
                        <p className="text-blue-600/80 font-semibold mb-1 uppercase tracking-wider text-[10px]">Mason Extra</p>
                        <p className="text-base font-black text-blue-900">{row.masonOT} <span className="text-xs font-normal">hrs</span></p>
                      </div>
                      <div className="bg-orange-50/50 p-2.5 rounded-lg border border-orange-100">
                        <p className="text-orange-600/80 font-semibold mb-1 uppercase tracking-wider text-[10px]">Helper Days</p>
                        <p className="text-base font-black text-orange-900">{row.helperReg}</p>
                      </div>
                      <div className="bg-orange-50/50 p-2.5 rounded-lg border border-orange-100">
                        <p className="text-orange-600/80 font-semibold mb-1 uppercase tracking-wider text-[10px]">Helper Extra</p>
                        <p className="text-base font-black text-orange-900">{row.helperOT} <span className="text-xs font-normal">hrs</span></p>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

            </div>
          </div>
        )}
      </div>
    </div>
  );
}
import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import { 
  Upload, Download, FileText, UploadCloud, LayoutGrid, Search, 
  Save, Trash2, Database, Clock, Check, BarChart3, Edit2, 
  LogOut, Lock, Plus, X, Layers, IndianRupee, Calendar
} from 'lucide-react';

// FIREBASE IMPORTS
import { collection, addDoc, getDocs, deleteDoc, doc, updateDoc } from "firebase/firestore";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { db, auth } from './firebase'; 

export default function App() {
  // --- AUTHENTICATION STATE ---
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const [siteData, setSiteData] = useState([]);
  const [dayData, setDayData] = useState([]);
  const [activeTab, setActiveTab] = useState('day');
  const [sheetName, setSheetName] = useState("");
  const [sheetMonth, setSheetMonth] = useState(""); // NEW: Stores YYYY-MM
  const [searchQuery, setSearchQuery] = useState("");
  const fileInputRef = useRef(null);

  const [dashboardTab, setDashboardTab] = useState('upload'); 
  const [savedSheets, setSavedSheets] = useState([]);
  const [isSaving, setIsSaving] = useState(false);
  const [hasSavedCurrent, setHasSavedCurrent] = useState(false);
  const [isLoadingRecords, setIsLoadingRecords] = useState(false);
  
  // --- MULTI-SITE & DATE ANALYTICS STATES ---
  const [globalSearch, setGlobalSearch] = useState("");
  const [analyticsMode, setAnalyticsMode] = useState('single'); 
  const [compareSiteInput, setCompareSiteInput] = useState("");
  const [compareSitesList, setCompareSitesList] = useState([]);
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoginError("");
    try {
      await signInWithEmailAndPassword(auth, email, password);
    } catch (err) {
      setLoginError("Invalid email or password. Please try again.");
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
    setSiteData([]);
    setDayData([]);
    setDashboardTab('upload');
  };

  const fetchSavedSheets = async () => {
    if (!user) return;
    setIsLoadingRecords(true);
    try {
      const querySnapshot = await getDocs(collection(db, "attendance_sheets"));
      const sheets = [];
      querySnapshot.forEach((doc) => {
        sheets.push({ id: doc.id, ...doc.data() });
      });
      sheets.sort((a, b) => b.createdAt - a.createdAt);
      setSavedSheets(sheets);
    } catch (error) {
      console.error("Error fetching sheets:", error);
    }
    setIsLoadingRecords(false);
  };

  useEffect(() => {
    if (user) fetchSavedSheets();
  }, [user]);

  const formatCurrency = (amount) => {
    return new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(amount);
  };

  // --- SMART EXTRACTION PARSER (AUTO MONTH DETECT) ---
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const cleanFileName = file.name.replace(/\.[^/.]+$/, "");
    
    // Auto-detect billing period (Month/Year) from filename
    const monthNames = { jan:'01', feb:'02', mar:'03', apr:'04', may:'05', jun:'06', jul:'07', aug:'08', sep:'09', oct:'10', nov:'11', dec:'12' };
    const mMatch = cleanFileName.match(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i);
    const yMatch = cleanFileName.match(/(202\d)/);
    const m = mMatch ? monthNames[mMatch[0].toLowerCase()] : String(new Date().getMonth() + 1).padStart(2, '0');
    const y = yMatch ? yMatch[0] : String(new Date().getFullYear());
    const detectedMonth = `${y}-${m}`;

    const reader = new FileReader();
    
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      let rawData = [];

      setSheetName(cleanFileName); 
      setSheetMonth(detectedMonth);
      const currentSheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(currentSheet, { header: 1, defval: "" });

      let activeSection = "Helper"; 
      let dayHeaders = {};
      let wagesColIdx = -1;
      let otWagesColIdx = -1;

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
          wagesColIdx = -1;
          otWagesColIdx = -1;

          for (let c = 2; c < row.length; c++) {
            const val = String(row[c]).trim();
            if (/^\d+$/.test(val)) {
              tempHeaders[c] = val;
            } else if (val.toUpperCase() === "WAGES") {
              wagesColIdx = c;
            } else if (val.toUpperCase().includes("OT WAGES") || val.toUpperCase() === "OT (IN RS)") {
              otWagesColIdx = c;
            }
          }
          if (Object.keys(tempHeaders).length > 0) { dayHeaders = tempHeaders; }
          continue;
        }

        if (/^\d+$/.test(col0) && col1 !== "" && col1 !== "OT" && col1 !== "NAN") {
          const nextRow = rows[r + 1] || [];
          const isOTRow = String(nextRow[1]).trim().toUpperCase() === "OT";

          let dailyWage = 0;
          let hourlyOTRate = 0;
          
          if (wagesColIdx !== -1 && row[wagesColIdx]) {
            const parsedWage = parseFloat(String(row[wagesColIdx]).replace(/[^0-9.]/g, ''));
            if (!isNaN(parsedWage)) dailyWage = parsedWage;
          }
          if (otWagesColIdx !== -1 && row[otWagesColIdx]) {
            const parsedOT = parseFloat(String(row[otWagesColIdx]).replace(/[^0-9.]/g, ''));
            if (!isNaN(parsedOT)) hourlyOTRate = parsedOT;
          }

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
                  const rowBaseCost = regDays * dailyWage;
                  const rowOTCost = otHrs * hourlyOTRate;
                  rawData.push({ day: dayNum, site, type: rowWorkerType, reg: regDays, ot: otHrs, baseCost: rowBaseCost, otCost: rowOTCost });
                }
              }
            }
          }
        }
      }

      const sMap = {};
      rawData.forEach(r => {
        if (!sMap[r.site]) sMap[r.site] = { site: r.site, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalBaseCost: 0, totalOTCost: 0 };
        sMap[r.site].totalBaseCost += r.baseCost;
        sMap[r.site].totalOTCost += r.otCost;
        if (r.type === 'Mason') { sMap[r.site].masonReg += r.reg; sMap[r.site].masonOT += r.ot; } 
        else if (r.type === 'HalfMason') { sMap[r.site].halfMasonReg += r.reg; sMap[r.site].halfMasonOT += r.ot; } 
        else { sMap[r.site].helperReg += r.reg; sMap[r.site].helperOT += r.ot; }
      });
      setSiteData(Object.values(sMap).sort((a, b) => a.site.localeCompare(b.site)));

      const dMap = {};
      rawData.forEach(r => {
        const key = `${r.day}|${r.site}`;
        if (!dMap[key]) dMap[key] = { day: r.day, site: r.site, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalBaseCost: 0, totalOTCost: 0 };
        dMap[key].totalBaseCost += r.baseCost;
        dMap[key].totalOTCost += r.otCost;
        if (r.type === 'Mason') { dMap[key].masonReg += r.reg; dMap[key].masonOT += r.ot; } 
        else if (r.type === 'HalfMason') { dMap[key].halfMasonReg += r.reg; dMap[key].halfMasonOT += r.ot; } 
        else { dMap[key].helperReg += r.reg; dMap[key].helperOT += r.ot; }
      });
      setDayData(Object.values(dMap).sort((a, b) => a.day - b.day || a.site.localeCompare(b.site)));
      
      setHasSavedCurrent(false);
      setSearchQuery("");
    };
    reader.readAsBinaryString(file);
  };

  const saveToDatabase = async () => {
    if (hasSavedCurrent) return;
    setIsSaving(true);
    try {
      const docRef = await addDoc(collection(db, "attendance_sheets"), {
        sheetName: sheetName,
        sheetMonth: sheetMonth, // Saved for exact date queries
        siteData: siteData,
        dayData: dayData,
        createdAt: Date.now(),
        ownerId: user.uid 
      });
      setHasSavedCurrent(true);
      setSavedSheets([{ id: docRef.id, sheetName, sheetMonth, siteData, dayData, createdAt: Date.now() }, ...savedSheets]);
    } catch (error) {
      console.error("Error saving document: ", error);
      alert("Failed to save to cloud.");
    }
    setIsSaving(false);
  };

  const renameRecord = async (id, currentName, e) => {
    e.stopPropagation(); 
    const newName = window.prompt("Enter a new name for this sheet:", currentName);
    if (!newName || newName.trim() === "" || newName === currentName) return;
    try {
      await updateDoc(doc(db, "attendance_sheets", id), { sheetName: newName.trim() });
      setSavedSheets(savedSheets.map(sheet => sheet.id === id ? { ...sheet, sheetName: newName.trim() } : sheet));
      if (hasSavedCurrent && sheetName === currentName) setSheetName(newName.trim());
    } catch (error) { console.error("Error renaming document: ", error); }
  };

  const deleteRecord = async (id, e) => {
    e.stopPropagation(); 
    if (!window.confirm("Are you sure you want to permanently delete this record?")) return;
    try {
      await deleteDoc(doc(db, "attendance_sheets", id));
      setSavedSheets(savedSheets.filter(sheet => sheet.id !== id));
      if (hasSavedCurrent) { setSiteData([]); setDayData([]); }
    } catch (error) { console.error("Error deleting document: ", error); }
  };

  const loadSavedRecord = (sheet) => {
    setSheetName(sheet.sheetName);
    setSheetMonth(sheet.sheetMonth || "");
    setSiteData(sheet.siteData);
    setDayData(sheet.dayData);
    setHasSavedCurrent(true); 
    setSearchQuery("");
  };

  // --- TRUE DAY-LEVEL DATE FILTERING ENGINE ---
  const getGlobalAnalytics = () => {
    if (!globalSearch.trim()) return null;
    const targetSite = globalSearch.toLowerCase().trim();
    let results = [];
    let totals = { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalSpend: 0 };

    savedSheets.forEach(sheet => {
      const sMonth = sheet.sheetMonth || "2026-08"; // fallback
      let sheetMason = 0, sheetMasonOT = 0, sheetHM = 0, sheetHMOT = 0, sheetHelper = 0, sheetHelperOT = 0, sheetSpend = 0;
      let sheetHasMatchingData = false;

      // Extract specific days from the sheet that match the user's date range
      (sheet.dayData || []).forEach(d => {
        if (d.site.toLowerCase() === targetSite) {
          const recordDateStr = `${sMonth}-${String(d.day).padStart(2, '0')}`;
          
          let inRange = true;
          if (startDate && recordDateStr < startDate) inRange = false;
          if (endDate && recordDateStr > endDate) inRange = false;

          if (inRange) {
            sheetHasMatchingData = true;
            sheetMason += d.masonReg || 0; sheetMasonOT += d.masonOT || 0;
            sheetHM += d.halfMasonReg || 0; sheetHMOT += d.halfMasonOT || 0;
            sheetHelper += d.helperReg || 0; sheetHelperOT += d.helperOT || 0;
            sheetSpend += (d.totalBaseCost || 0) + (d.totalOTCost || 0);
          }
        }
      });

      if (sheetHasMatchingData) {
        results.push({ 
          sheetName: sheet.sheetName, 
          date: sheet.createdAt, 
          masonReg: sheetMason, masonOT: sheetMasonOT,
          halfMasonReg: sheetHM, halfMasonOT: sheetHMOT,
          helperReg: sheetHelper, helperOT: sheetHelperOT,
          totalBaseCost: sheetSpend, totalOTCost: 0 // Aggregated
        });
        totals.masonReg += sheetMason; totals.masonOT += sheetMasonOT;
        totals.halfMasonReg += sheetHM; totals.halfMasonOT += sheetHMOT;
        totals.helperReg += sheetHelper; totals.helperOT += sheetHelperOT;
        totals.totalSpend += sheetSpend;
      }
    });
    return { results, totals };
  };
  const analyticsData = getGlobalAnalytics();

  // --- MULTI-SITE DAY-LEVEL ENGINE ---
  const handleAddCompareSite = (e) => {
    e.preventDefault();
    const code = compareSiteInput.trim();
    if (code && !compareSitesList.includes(code.toUpperCase())) {
      setCompareSitesList([...compareSitesList, code.toUpperCase()]);
    }
    setCompareSiteInput('');
  };

  const removeCompareSite = (siteToRemove) => {
    setCompareSitesList(compareSitesList.filter(s => s !== siteToRemove));
  };

  const getMultiSiteAnalytics = () => {
    if (compareSitesList.length === 0) return null;
    let matrix = [];
    let totals = { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalDays: 0, totalOT: 0, totalSpend: 0 };

    compareSitesList.forEach(siteCode => {
      let sTotals = { occurrences: 0, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, siteSpend: 0 };
      
      savedSheets.forEach(sheet => {
        const sMonth = sheet.sheetMonth || "2026-08";
        let sheetHasMatchingData = false;

        (sheet.dayData || []).forEach(d => {
          if (d.site.toLowerCase() === siteCode.toLowerCase()) {
            const recordDateStr = `${sMonth}-${String(d.day).padStart(2, '0')}`;
            let inRange = true;
            if (startDate && recordDateStr < startDate) inRange = false;
            if (endDate && recordDateStr > endDate) inRange = false;

            if (inRange) {
              sheetHasMatchingData = true;
              sTotals.masonReg += d.masonReg || 0; sTotals.masonOT += d.masonOT || 0;
              sTotals.halfMasonReg += d.halfMasonReg || 0; sTotals.halfMasonOT += d.halfMasonOT || 0;
              sTotals.helperReg += d.helperReg || 0; sTotals.helperOT += d.helperOT || 0;
              sTotals.siteSpend += (d.totalBaseCost || 0) + (d.totalOTCost || 0);
            }
          }
        });
        if (sheetHasMatchingData) sTotals.occurrences++;
      });
      
      const totalDays = sTotals.masonReg + sTotals.halfMasonReg + sTotals.helperReg;
      const totalOT = sTotals.masonOT + sTotals.halfMasonOT + sTotals.helperOT;

      totals.masonReg += sTotals.masonReg; totals.masonOT += sTotals.masonOT;
      totals.halfMasonReg += sTotals.halfMasonReg; totals.halfMasonOT += sTotals.halfMasonOT;
      totals.helperReg += sTotals.helperReg; totals.helperOT += sTotals.helperOT;
      totals.totalDays += totalDays; totals.totalOT += totalOT;
      totals.totalSpend += sTotals.siteSpend;

      matrix.push({ site: siteCode, ...sTotals, totalDays, totalOT });
    });
    return { matrix, totals };
  };
  const multiSiteData = getMultiSiteAnalytics();

  // --- ACTIVE DATA LOGIC ---
  const activeData = activeTab === 'site' ? siteData : dayData;
  const filteredData = activeData.filter(row => row.site.toLowerCase().includes(searchQuery.toLowerCase()));
  const totals = filteredData.reduce((acc, row) => {
    acc.masonReg += row.masonReg || 0; acc.masonOT += row.masonOT || 0;
    acc.halfMasonReg += row.halfMasonReg || 0; acc.halfMasonOT += row.halfMasonOT || 0;
    acc.helperReg += row.helperReg || 0; acc.helperOT += row.helperOT || 0;
    acc.totalBaseCost += row.totalBaseCost || 0; acc.totalOTCost += row.totalOTCost || 0;
    return acc;
  }, { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalBaseCost: 0, totalOTCost: 0 });

  const exportToExcel = () => {
    const formattedData = filteredData.map(row => {
      const base = {};
      if (activeTab === 'day') base["Date (Day)"] = row.day;
      return {
        ...base,
        "Site Code": row.site, "Mason Days": row.masonReg, "Mason Extra Hours": row.masonOT,
        "Half Mason Days": row.halfMasonReg, "Half Mason Extra": row.halfMasonOT,
        "Helper Days": row.helperReg, "Helper Extra Hours": row.helperOT,
        "Base Cost (Rs)": row.totalBaseCost || 0, "OT Cost (Rs)": row.totalOTCost || 0,
        "Total Cost (Rs)": (row.totalBaseCost || 0) + (row.totalOTCost || 0)
      };
    });
    const totalRow = { "Site Code": "GRAND TOTAL" };
    if (activeTab === 'day') totalRow["Date (Day)"] = "";
    totalRow["Mason Days"] = totals.masonReg; totalRow["Mason Extra Hours"] = totals.masonOT;
    totalRow["Half Mason Days"] = totals.halfMasonReg; totalRow["Half Mason Extra"] = totals.halfMasonOT;
    totalRow["Helper Days"] = totals.helperReg; totalRow["Helper Extra Hours"] = totals.helperOT;
    totalRow["Base Cost (Rs)"] = totals.totalBaseCost; totalRow["OT Cost (Rs)"] = totals.totalOTCost;
    totalRow["Total Cost (Rs)"] = totals.totalBaseCost + totals.totalOTCost;
    formattedData.push(totalRow);
    const ws = XLSX.utils.json_to_sheet(formattedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance Data");
    XLSX.writeFile(wb, `CRM_FIX_${sheetName}_${activeTab === 'day' ? 'DayWise' : 'SiteWise'}.xlsx`);
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    const viewTitle = activeTab === 'day' ? 'Day-Wise Breakdown' : 'Site-Wise Summary';
    doc.setFontSize(16); doc.setTextColor(37, 99, 235); doc.text("CRM_FIX FINANCIAL EXPORT", 14, 15);
    doc.setFontSize(10); doc.setTextColor(100); doc.text(`Sheet: ${sheetName} | ${viewTitle}`, 14, 22);
    
    const tableColumn = activeTab === 'day' 
      ? ["Day", "Site", "Mason D", "Mas Ex", "HM D", "HM Ex", "Help D", "Help Ex", "Cost (Rs)"]
      : ["Site Code", "Mason Days", "Mason Extra", "HM Days", "HM Extra", "Helper Days", "Helper Extra", "Total Cost"];
    const tableRows = filteredData.map(row => {
      const cost = formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0));
      return activeTab === 'day'
        ? [row.day, row.site, row.masonReg, row.masonOT, row.halfMasonReg, row.halfMasonOT, row.helperReg, row.helperOT, cost]
        : [row.site, row.masonReg, row.masonOT, row.halfMasonReg, row.halfMasonOT, row.helperReg, row.helperOT, cost]
    });
    const grandCost = formatCurrency(totals.totalBaseCost + totals.totalOTCost);
    const footRow = activeTab === 'day'
      ? ["", "TOTAL", totals.masonReg, totals.masonOT, totals.halfMasonReg, totals.halfMasonOT, totals.helperReg, totals.helperOT, grandCost]
      : ["TOTAL", totals.masonReg, totals.masonOT, totals.halfMasonReg, totals.halfMasonOT, totals.helperReg, totals.helperOT, grandCost];

    doc.autoTable({ head: [tableColumn], body: tableRows, foot: [footRow], startY: 28, theme: 'grid', headStyles: { fillColor: [37, 99, 235] }, footStyles: { fillColor: [243, 244, 246], textColor: [17, 24, 39], fontStyle: 'bold' } });
    doc.save(`CRM_FIX_Finance_${sheetName}.pdf`);
  };

  if (authLoading) {
    return <div className="min-h-screen flex items-center justify-center bg-[#f8fafc] text-gray-500 font-bold">Initializing Secure Environment...</div>;
  }

  if (!user) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#f8fafc] px-4 font-sans selection:bg-blue-100 selection:text-blue-900">
        <div className="w-full max-w-md bg-white p-8 md:p-10 rounded-[2rem] shadow-xl border border-gray-100">
          <div className="flex flex-col items-center mb-8 text-center">
            <div className="bg-blue-600 p-4 rounded-2xl shadow-lg shadow-blue-600/20 mb-5"><Lock className="w-8 h-8 text-white" /></div>
            <h1 className="text-3xl font-black text-gray-900 tracking-tight">CRM_FIX</h1>
            <p className="text-sm font-bold text-gray-400 uppercase tracking-widest mt-1">Authorized Access Only</p>
          </div>
          <form onSubmit={handleLogin} className="space-y-5">
            <div>
              <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Email Address</label>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" placeholder="Enter your email" />
            </div>
            <div>
              <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Password</label>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" placeholder="Enter your password" />
            </div>
            {loginError && <p className="text-red-500 text-xs font-bold text-center bg-red-50 py-2 rounded-lg">{loginError}</p>}
            <button type="submit" className="w-full py-4 bg-gray-900 hover:bg-blue-600 text-white font-bold rounded-xl shadow-md transition-colors duration-300">Access Dashboard</button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#f8fafc] text-gray-800 font-sans selection:bg-blue-100 selection:text-blue-900 pb-12">
      
      {/* GLOBAL HEADER */}
      <header className="bg-white border-b border-gray-200 px-4 py-3 flex justify-between items-center shadow-sm">
        <div className="flex items-center gap-2">
          <div className="bg-blue-600 p-1.5 rounded-lg"><LayoutGrid className="w-4 h-4 text-white" /></div>
          <span className="font-black text-gray-900 text-sm tracking-tight">CRM_FIX</span>
        </div>
        <div className="flex items-center gap-4">
          <span className="text-xs font-bold text-gray-500 hidden md:inline-block">{user.email}</span>
          <button onClick={handleLogout} className="text-xs font-bold bg-gray-100 hover:bg-red-50 hover:text-red-600 text-gray-600 px-3 py-1.5 rounded-lg flex items-center gap-2 transition-colors">
            <LogOut className="w-3.5 h-3.5" /> Sign Out
          </button>
        </div>
      </header>

      <input type="file" accept=".xlsx, .xls, .csv" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
      
      {/* --- DASHBOARD --- */}
      {activeData.length === 0 ? (
        <div className="max-w-6xl mx-auto pt-10 md:pt-16 px-4">
          
          <div className="text-center mb-10 space-y-3">
            <h1 className="text-3xl md:text-5xl font-black tracking-tight text-gray-900">Welcome Back.</h1>
            <p className="text-sm md:text-base font-bold text-gray-400 uppercase tracking-widest">Financial & Attendance Engine</p>
          </div>

          {/* SLEEK NAVIGATION PILL */}
          <div className="flex justify-center mb-12">
            <div className="bg-gray-100 p-1.5 rounded-full inline-flex shadow-inner overflow-x-auto max-w-full custom-scrollbar">
              <button onClick={() => setDashboardTab('upload')} className={`whitespace-nowrap px-6 py-3 rounded-full font-bold text-sm transition-all flex items-center gap-2.5 ${dashboardTab === 'upload' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                <UploadCloud className={`w-4 h-4 ${dashboardTab === 'upload' ? 'text-gray-900' : 'text-gray-400'}`} /> Upload New
              </button>
              <button onClick={() => { setDashboardTab('records'); fetchSavedSheets(); }} className={`whitespace-nowrap px-6 py-3 rounded-full font-bold text-sm transition-all flex items-center gap-2.5 ${dashboardTab === 'records' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                <Database className={`w-4 h-4 ${dashboardTab === 'records' ? 'text-gray-900' : 'text-gray-400'}`} /> Saved Records
              </button>
              <button onClick={() => { setDashboardTab('analytics'); fetchSavedSheets(); }} className={`whitespace-nowrap px-6 py-3 rounded-full font-bold text-sm transition-all flex items-center gap-2.5 ${dashboardTab === 'analytics' ? 'bg-white text-blue-600 shadow-sm border border-gray-50' : 'text-gray-500 hover:text-gray-700'}`}>
                <BarChart3 className={`w-4 h-4 ${dashboardTab === 'analytics' ? 'text-blue-600' : 'text-gray-400'}`} /> Global Analytics
              </button>
            </div>
          </div>

          {/* TAB 1: UPLOAD ZONE */}
          {dashboardTab === 'upload' && (
            <div onClick={() => fileInputRef.current?.click()} className="group w-full max-w-2xl mx-auto cursor-pointer bg-white rounded-[2.5rem] border-2 border-dashed border-gray-200 hover:border-blue-500 hover:bg-blue-50/50 transition-all duration-300 p-8 md:p-20 shadow-sm hover:shadow-xl text-center">
              <div className="flex flex-col items-center justify-center space-y-6">
                <div className="bg-blue-50 text-blue-600 p-6 rounded-full group-hover:scale-110 group-hover:bg-blue-100 transition-all duration-300"><UploadCloud className="w-12 h-12" /></div>
                <div><p className="text-xl md:text-2xl font-black text-gray-800">Upload Financial Sheet</p><p className="text-sm md:text-base font-medium text-gray-500 mt-2">Parse and visualize site labor costs instantly.</p></div>
                <span className="mt-4 px-8 py-3.5 bg-gray-900 text-white font-bold rounded-2xl shadow-md group-hover:bg-blue-600 transition-colors duration-300 flex items-center gap-2"><Upload className="w-4 h-4" />Select File</span>
              </div>
            </div>
          )}

          {/* TAB 2: SAVED RECORDS */}
          {dashboardTab === 'records' && (
            <div className="bg-white rounded-[2.5rem] p-6 md:p-10 shadow-sm border border-gray-100 min-h-[400px]">
              {isLoadingRecords ? (
                <div className="flex justify-center items-center h-48 text-gray-400 font-bold">Loading cloud records...</div>
              ) : savedSheets.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-48 text-gray-400 space-y-4"><Database className="w-12 h-12 text-gray-200" /><p className="font-semibold text-center">No saved sheets found.<br/><span className="text-sm font-normal">Upload and save a sheet to see it here.</span></p></div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
                  {savedSheets.map((sheet) => (
                    <div key={sheet.id} onClick={() => loadSavedRecord(sheet)} className="group bg-white hover:bg-blue-50/50 border border-gray-100 hover:border-blue-200 rounded-[1.5rem] p-6 cursor-pointer transition-all flex flex-col justify-between space-y-5 shadow-[0_2px_10px_-3px_rgba(0,0,0,0.05)] hover:shadow-lg">
                      <div className="flex items-start justify-between">
                        <div className="bg-blue-50 text-blue-600 p-3.5 rounded-2xl"><FileText className="w-5 h-5" /></div>
                        <div className="flex gap-1.5">
                          <button onClick={(e) => renameRecord(sheet.id, sheet.sheetName, e)} className="text-gray-400 hover:text-blue-600 bg-gray-50 hover:bg-white shadow-sm border border-transparent hover:border-blue-100 rounded-xl p-2 transition-all" title="Rename"><Edit2 className="w-4 h-4" /></button>
                          <button onClick={(e) => deleteRecord(sheet.id, e)} className="text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-white shadow-sm border border-transparent hover:border-red-100 rounded-xl p-2 transition-all" title="Delete"><Trash2 className="w-4 h-4" /></button>
                        </div>
                      </div>
                      <div>
                        <h3 className="font-black text-gray-800 text-base md:text-lg truncate group-hover:text-blue-700" title={sheet.sheetName}>{sheet.sheetName}</h3>
                        <div className="flex items-center justify-between mt-2">
                          <p className="text-[10px] md:text-xs font-bold text-gray-400 flex items-center gap-1.5"><Clock className="w-3 h-3" /> {new Date(sheet.createdAt).toLocaleDateString()}</p>
                          {sheet.sheetMonth && <span className="bg-gray-100 text-gray-500 text-[10px] font-bold px-2 py-1 rounded-md">{sheet.sheetMonth}</span>}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* TAB 3: GLOBAL ANALYTICS (PRISTINE UI) */}
          {dashboardTab === 'analytics' && (
            <div className="max-w-4xl mx-auto space-y-6">
              
              {/* WHITE DASHBOARD CARD */}
              <div className="bg-white px-6 py-10 md:p-12 rounded-[2.5rem] shadow-[0_2px_15px_-3px_rgba(0,0,0,0.07),0_10px_20px_-2px_rgba(0,0,0,0.04)] border border-gray-100 text-center">
                <h2 className="text-xl md:text-2xl font-black text-gray-900 mb-8 tracking-tight">Master Site Analytics</h2>
                
                {/* --- SEGMENTED TOGGLE --- */}
                <div className="flex justify-center mb-8">
                  <div className="bg-gray-50 border border-gray-100 p-1.5 rounded-2xl inline-flex shadow-inner">
                    <button onClick={() => setAnalyticsMode('single')} className={`px-6 py-2.5 rounded-xl font-bold text-sm transition-all flex items-center gap-2 ${analyticsMode === 'single' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                      <Search className={`w-4 h-4 ${analyticsMode === 'single' ? 'text-blue-600' : 'text-gray-400'}`}/> Single Site Lookup
                    </button>
                    <button onClick={() => setAnalyticsMode('compare')} className={`px-6 py-2.5 rounded-xl font-bold text-sm transition-all flex items-center gap-2 ${analyticsMode === 'compare' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                      <Layers className={`w-4 h-4 ${analyticsMode === 'compare' ? 'text-blue-600' : 'text-gray-400'}`}/> Multi-Site Compare
                    </button>
                  </div>
                </div>

                {/* --- INPUTS --- */}
                {analyticsMode === 'single' ? (
                  <div className="relative w-full max-w-md mx-auto mb-8">
                    <Search className="w-5 h-5 absolute left-5 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input type="text" placeholder="Enter site code..." value={globalSearch} onChange={(e) => setGlobalSearch(e.target.value)} className="w-full pl-14 pr-4 py-4 bg-gray-50/80 border border-gray-200 rounded-2xl text-base font-bold focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:bg-white focus:border-blue-500 transition-all shadow-sm"/>
                  </div>
                ) : (
                  <div className="w-full max-w-xl mx-auto mb-8">
                    <form onSubmit={handleAddCompareSite} className="flex flex-col sm:flex-row items-center gap-3">
                      <div className="relative flex-1 w-full">
                        <Plus className="w-5 h-5 absolute left-5 top-1/2 -translate-y-1/2 text-gray-400" />
                        <input type="text" placeholder="Type site code (e.g. S01) & press Enter..." value={compareSiteInput} onChange={(e) => setCompareSiteInput(e.target.value)} className="w-full pl-14 pr-4 py-4 bg-gray-50/80 border border-gray-200 rounded-2xl text-base font-bold focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:bg-white focus:border-blue-500 transition-all shadow-sm"/>
                      </div>
                      <button type="submit" className="w-full sm:w-auto bg-blue-600 hover:bg-blue-700 text-white px-8 py-4 rounded-2xl font-black flex items-center justify-center gap-2 transition-colors shadow-md">Add</button>
                    </form>
                    {compareSitesList.length > 0 && (
                      <div className="flex flex-wrap gap-2 justify-center mt-5">
                        {compareSitesList.map(site => (
                          <div key={site} className="bg-blue-50 border border-blue-200 text-blue-800 px-4 py-2.5 rounded-xl text-sm font-black flex items-center gap-2 shadow-sm">
                            {site}
                            <button onClick={() => removeCompareSite(site)} className="text-blue-400 hover:text-red-500 transition-colors bg-white rounded-md p-0.5"><X className="w-4 h-4"/></button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}

                {/* --- PERFECT DATE FILTER BAR --- */}
                <div className="flex flex-col sm:flex-row items-center justify-center gap-3 pt-8 border-t border-gray-100 max-w-2xl mx-auto">
                  <div className="flex items-center bg-gray-50 border border-gray-200 rounded-[1.25rem] p-1.5 shadow-inner w-full sm:w-auto">
                    
                    <div className="flex items-center bg-white px-4 py-2.5 rounded-xl shadow-sm border border-gray-100 flex-1 sm:w-40 relative">
                      <Calendar className="w-4 h-4 text-blue-500 mr-3 shrink-0" />
                      <div className="flex flex-col items-start w-full">
                        <span className="text-[8px] font-black text-gray-400 uppercase tracking-widest leading-none mb-1">Start Date</span>
                        <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="w-full text-xs font-black text-gray-800 outline-none bg-transparent cursor-pointer leading-none" />
                      </div>
                    </div>
                    
                    <span className="px-3 text-gray-300 font-black text-xs shrink-0">to</span>

                    <div className="flex items-center bg-white px-4 py-2.5 rounded-xl shadow-sm border border-gray-100 flex-1 sm:w-40 relative">
                      <Calendar className="w-4 h-4 text-gray-500 mr-3 shrink-0" />
                      <div className="flex flex-col items-start w-full">
                        <span className="text-[8px] font-black text-gray-400 uppercase tracking-widest leading-none mb-1">End Date</span>
                        <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="w-full text-xs font-black text-gray-800 outline-none bg-transparent cursor-pointer leading-none" />
                      </div>
                    </div>
                  </div>

                  {(startDate || endDate) && (
                    <button onClick={() => { setStartDate(""); setEndDate(""); }} className="p-4 bg-red-50 hover:bg-red-100 text-red-500 rounded-[1.25rem] transition-colors border border-red-100 shrink-0 shadow-sm" title="Clear Dates">
                      <X className="w-5 h-5" />
                    </button>
                  )}
                </div>
              </div>

              {/* SINGLE SITE MODE RESULTS */}
              {analyticsMode === 'single' && globalSearch.trim() && analyticsData && (
                <>
                  {analyticsData.results.length === 0 ? (
                    <div className="text-center py-12 text-gray-500 font-medium bg-white rounded-3xl border border-gray-100">No records found for site: "{globalSearch}" in this period.</div>
                  ) : (
                    <div className="space-y-6">
                      <div className="bg-white rounded-3xl shadow-sm border border-gray-200 overflow-hidden">
                        <div className="bg-gray-50 border-b border-gray-100 p-4 font-bold text-gray-600 text-xs md:text-sm uppercase tracking-wider truncate">Occurrence Timeline: <span className="text-blue-600">"{globalSearch.toUpperCase()}"</span></div>
                        <div className="divide-y divide-gray-50">
                          {analyticsData.results.map((record, idx) => (
                            <div key={idx} className="p-4 hover:bg-gray-50 transition-colors flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                              <div className="max-w-full md:max-w-[200px]">
                                <h4 className="font-black text-gray-800 text-base md:text-lg truncate" title={record.sheetName}>{record.sheetName}</h4>
                                <p className="text-xs font-semibold text-gray-400 flex items-center gap-1 mt-0.5">
                                  <Clock className="w-3 h-3" /> {new Date(record.date).toLocaleDateString()}
                                </p>
                              </div>
                              <div className="flex gap-2 md:gap-4 flex-wrap w-full md:w-auto justify-between md:justify-end items-center">
                                <div className="text-center bg-white p-2 rounded-xl border border-gray-100 flex-1 md:flex-none min-w-[75px]">
                                  <p className="text-[9px] md:text-[10px] font-bold text-blue-500 uppercase">Mason</p>
                                  <p className="font-black text-gray-900 text-sm">{record.masonReg} <span className="text-[10px] text-gray-400 font-medium">({record.masonOT}h)</span></p>
                                </div>
                                <div className="text-center bg-white p-2 rounded-xl border border-gray-100 flex-1 md:flex-none min-w-[75px]">
                                  <p className="text-[9px] md:text-[10px] font-bold text-purple-500 uppercase">H. Mason</p>
                                  <p className="font-black text-gray-900 text-sm">{record.halfMasonReg} <span className="text-[10px] text-gray-400 font-medium">({record.halfMasonOT}h)</span></p>
                                </div>
                                <div className="text-center bg-white p-2 rounded-xl border border-gray-100 flex-1 md:flex-none min-w-[75px]">
                                  <p className="text-[9px] md:text-[10px] font-bold text-orange-500 uppercase">Helper</p>
                                  <p className="font-black text-gray-900 text-sm">{record.helperReg} <span className="text-[10px] text-gray-400 font-medium">({record.helperOT}h)</span></p>
                                </div>
                                <div className="text-center bg-emerald-50 p-2 rounded-xl border border-emerald-100 flex-1 md:flex-none min-w-[95px]">
                                  <p className="text-[9px] md:text-[10px] font-bold text-emerald-600 uppercase">Spend</p>
                                  <p className="font-black text-emerald-900 text-sm">{formatCurrency((record.totalBaseCost||0) + (record.totalOTCost||0))}</p>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div className="bg-gray-900 p-6 md:p-8 rounded-[2rem] shadow-2xl border border-gray-800">
                        <h3 className="text-white font-black text-center text-lg md:text-xl mb-6 tracking-wider flex items-center justify-center gap-2">
                          <Database className="w-5 h-5 text-blue-400" /> TOTALS FOR "{globalSearch.toUpperCase()}"
                        </h3>
                        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-4 text-xs">
                          <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-gray-400 font-bold mb-1 uppercase text-[10px] tracking-wider">Mason Days</p><p className="text-2xl md:text-3xl font-black text-white">{analyticsData.totals.masonReg}</p></div>
                          <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-gray-400 font-bold mb-1 uppercase text-[10px] tracking-wider">Mason Extra</p><p className="text-2xl md:text-3xl font-black text-white">{analyticsData.totals.masonOT} <span className="text-xs font-medium text-gray-500">h</span></p></div>
                          <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-purple-400 font-bold mb-1 uppercase text-[10px] tracking-wider">HM Days</p><p className="text-2xl md:text-3xl font-black text-purple-200">{analyticsData.totals.halfMasonReg}</p></div>
                          <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-purple-400 font-bold mb-1 uppercase text-[10px] tracking-wider">HM Extra</p><p className="text-2xl md:text-3xl font-black text-purple-200">{analyticsData.totals.halfMasonOT} <span className="text-xs font-medium text-gray-500">h</span></p></div>
                          <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-orange-400 font-bold mb-1 uppercase text-[10px] tracking-wider">Helper Days</p><p className="text-2xl md:text-3xl font-black text-white">{analyticsData.totals.helperReg}</p></div>
                          <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-orange-400 font-bold mb-1 uppercase text-[10px] tracking-wider">Helper Extra</p><p className="text-2xl md:text-3xl font-black text-white">{analyticsData.totals.helperOT} <span className="text-xs font-medium text-gray-500">h</span></p></div>
                          <div className="bg-blue-600/90 p-4 rounded-2xl border border-blue-500"><p className="text-blue-100 font-bold mb-1 uppercase text-[10px] tracking-wider">Total Man-Days</p><p className="text-2xl md:text-3xl font-black text-white">{analyticsData.totals.masonReg + analyticsData.totals.halfMasonReg + analyticsData.totals.helperReg}</p></div>
                          <div className="bg-emerald-600 p-4 rounded-2xl border border-emerald-500"><p className="text-emerald-100 font-bold mb-1 uppercase text-[10px] tracking-wider">Total Spend</p><p className="text-2xl md:text-3xl font-black text-white">{formatCurrency(analyticsData.totals.totalSpend)}</p></div>
                        </div>
                      </div>
                    </div>
                  )}
                </>
              )}

              {/* MULTI-SITE COMPARE RESULTS */}
              {analyticsMode === 'compare' && compareSitesList.length > 0 && multiSiteData && (
                <div className="space-y-6">
                  <div className="bg-white rounded-3xl shadow-sm border border-gray-200 overflow-hidden">
                    <div className="bg-gray-50 border-b border-gray-100 p-4 font-bold text-gray-600 text-xs md:text-sm uppercase tracking-wider">
                      Comparison Matrix <span className="text-gray-400 normal-case font-medium">({compareSitesList.length} sites)</span>
                    </div>
                    <div className="overflow-x-auto custom-scrollbar">
                      <table className="w-full text-sm text-left">
                        <thead className="bg-white uppercase text-[10px] md:text-[11px] font-black tracking-wider text-gray-500 border-b border-gray-100">
                          <tr>
                            <th className="px-3.5 py-4 border-r border-gray-50">Site Code</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-center text-gray-400">Sheets</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-blue-800 bg-blue-50/40">Mason D</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-blue-800 bg-blue-50/40">Mason OT</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-purple-800 bg-purple-50/40">HM Days</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-purple-800 bg-purple-50/40">HM OT</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-orange-800 bg-orange-50/40">Helper D</th>
                            <th className="px-3.5 py-4 border-r border-gray-50 text-orange-800 bg-orange-50/40">Helper OT</th>
                            <th className="px-3.5 py-4 text-gray-900 bg-gray-100/60 font-black border-r border-gray-50">Total Days</th>
                            <th className="px-3.5 py-4 text-emerald-800 bg-emerald-50/60 font-black">Labor Cost</th>
                          </tr>
                        </thead>
                        <tbody>
                          {multiSiteData.matrix.map((row, idx) => (
                            <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/60 transition-colors">
                              <td className="px-3.5 py-4 font-black text-gray-800 border-r border-gray-50">{row.site}</td>
                              <td className="px-3.5 py-4 text-center font-bold text-gray-400 border-r border-gray-50">{row.occurrences}</td>
                              <td className="px-3.5 py-4 font-bold text-blue-900 border-r border-gray-50">{row.masonReg}</td>
                              <td className="px-3.5 py-4 text-blue-700 border-r border-gray-50">{row.masonOT} <span className="text-[10px] text-gray-400">h</span></td>
                              <td className="px-3.5 py-4 font-bold text-purple-900 border-r border-gray-50">{row.halfMasonReg}</td>
                              <td className="px-3.5 py-4 text-purple-700 border-r border-gray-50">{row.halfMasonOT} <span className="text-[10px] text-gray-400">h</span></td>
                              <td className="px-3.5 py-4 font-bold text-orange-900 border-r border-gray-50">{row.helperReg}</td>
                              <td className="px-3.5 py-4 text-orange-700 border-r border-gray-50">{row.helperOT} <span className="text-[10px] text-gray-400">h</span></td>
                              <td className="px-3.5 py-4 font-black text-gray-900 bg-gray-50/50 border-r border-gray-100">{row.totalDays}</td>
                              <td className="px-3.5 py-4 font-black text-emerald-700 bg-emerald-50/30">{formatCurrency(row.siteSpend)}</td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot className="bg-gray-100/80 font-black text-gray-900 border-t-2 border-gray-200">
                          <tr>
                            <td className="px-3.5 py-4 border-r border-gray-200">COMBINED:</td>
                            <td className="px-3.5 py-4 text-center border-r border-gray-200">-</td>
                            <td className="px-3.5 py-4 text-blue-900 border-r border-gray-200 text-base">{multiSiteData.totals.masonReg}</td>
                            <td className="px-3.5 py-4 text-blue-700 border-r border-gray-200">{multiSiteData.totals.masonOT} h</td>
                            <td className="px-3.5 py-4 text-purple-900 border-r border-gray-200 text-base">{multiSiteData.totals.halfMasonReg}</td>
                            <td className="px-3.5 py-4 text-purple-700 border-r border-gray-200">{multiSiteData.totals.halfMasonOT} h</td>
                            <td className="px-3.5 py-4 text-orange-900 border-r border-gray-200 text-base">{multiSiteData.totals.helperReg}</td>
                            <td className="px-3.5 py-4 text-orange-700 border-r border-gray-200">{multiSiteData.totals.helperOT} h</td>
                            <td className="px-3.5 py-4 text-gray-900 text-base bg-gray-200/70 border-r border-gray-300">{multiSiteData.totals.totalDays}</td>
                            <td className="px-3.5 py-4 text-emerald-900 text-base bg-emerald-100/70">{formatCurrency(multiSiteData.totals.totalSpend)}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>

                  <div className="bg-gray-900 p-6 md:p-8 rounded-[2rem] shadow-2xl border border-gray-800">
                    <h3 className="text-white font-black text-center text-lg md:text-xl mb-6 tracking-wider flex items-center justify-center gap-2"><Database className="w-5 h-5 text-blue-400" /> SELECTED SITES TOTALS</h3>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 text-xs">
                      <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-gray-400 font-bold mb-1 uppercase text-[10px] tracking-wider">Mason Days</p><p className="text-2xl md:text-3xl font-black text-white">{multiSiteData.totals.masonReg}</p></div>
                      <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-purple-400 font-bold mb-1 uppercase text-[10px] tracking-wider">HM Days</p><p className="text-2xl md:text-3xl font-black text-purple-200">{multiSiteData.totals.halfMasonReg}</p></div>
                      <div className="bg-gray-800/80 p-4 rounded-2xl"><p className="text-orange-400 font-bold mb-1 uppercase text-[10px] tracking-wider">Helper Days</p><p className="text-2xl md:text-3xl font-black text-white">{multiSiteData.totals.helperReg}</p></div>
                      <div className="bg-emerald-600 border border-emerald-500 p-4 rounded-2xl"><p className="text-emerald-100 font-bold mb-1 uppercase text-[10px] tracking-wider">Total Financial Spend</p><p className="text-2xl md:text-3xl font-black text-white">{formatCurrency(multiSiteData.totals.totalSpend)}</p></div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      ) : (

      /* --- ACTIVE PARSED SHEET VIEW (WITH FINANCIALS) --- */
        <div className="max-w-7xl mx-auto space-y-4 md:space-y-6 p-2 md:p-8 pt-4 md:pt-8 pb-12">
          
          <div className="bg-white p-4 md:p-5 rounded-[2rem] shadow-sm border border-gray-100 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div className="flex flex-col md:flex-row gap-2 md:gap-4 items-start md:items-center w-full md:w-auto">
              <div className="bg-blue-600 p-2.5 rounded-xl text-white shadow-sm cursor-pointer shrink-0" onClick={() => { setSiteData([]); setDayData([]); }}>
                <LayoutGrid className="w-5 h-5 md:w-6 md:h-6" />
              </div>
              <div className="min-w-0 flex-1">
                <h1 className="text-xl md:text-2xl font-black tracking-tight text-gray-900 leading-none truncate" title={sheetName}>{sheetName}</h1>
                <p className="text-[10px] md:text-xs font-bold uppercase tracking-wider text-emerald-500 mt-1 flex items-center gap-1"><IndianRupee className="w-3 h-3"/> Financial Engine Active</p>
              </div>
              
              {/* EDITABLE SHEET MONTH SELECTOR */}
              <div className="flex items-center gap-2 bg-gray-50 border border-gray-200 px-3 py-2 rounded-xl shadow-inner mt-2 md:mt-0">
                <Calendar className="w-4 h-4 text-gray-400" />
                <input type="month" value={sheetMonth} onChange={(e) => setSheetMonth(e.target.value)} className="bg-transparent text-sm font-black text-gray-700 outline-none cursor-pointer uppercase"/>
              </div>
            </div>
            
            <div className="flex w-full md:w-auto gap-2 md:gap-3">
              <button onClick={saveToDatabase} disabled={hasSavedCurrent || isSaving} className={`flex-1 justify-center px-4 md:px-6 py-3 rounded-xl font-black text-sm flex items-center gap-2 transition-all shadow-sm ${hasSavedCurrent ? 'bg-gray-100 text-gray-500 cursor-not-allowed border border-gray-200' : 'bg-indigo-600 hover:bg-indigo-700 text-white'}`}>
                {hasSavedCurrent ? <Check className="w-4 h-4" /> : <Save className="w-4 h-4" />}
                {hasSavedCurrent ? "Saved" : isSaving ? "Saving..." : "Save Insights"}
              </button>
              <button onClick={() => { setSiteData([]); setDayData([]); fileInputRef.current?.click(); }} className="flex-1 justify-center bg-gray-100 hover:bg-gray-200 text-gray-700 px-4 md:px-5 py-3 rounded-xl font-black text-sm flex items-center gap-2 transition-all">
                <Upload className="w-4 h-4" /> Upload
              </button>
            </div>
          </div>

          <div className="bg-white rounded-[2rem] shadow-sm border border-gray-100 overflow-hidden">
            <div className="flex flex-col lg:flex-row justify-between items-stretch lg:items-center border-b border-gray-100 bg-gray-50/50 p-2 md:p-4 gap-3">
              <div className="flex w-full lg:w-auto bg-gray-200/60 p-1 rounded-[1.25rem]">
                <button className={`flex-1 px-4 md:px-6 py-2 md:py-2.5 font-bold text-xs md:text-sm rounded-xl transition-all ${activeTab === 'day' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`} onClick={() => setActiveTab('day')}>Day-Wise</button>
                <button className={`flex-1 px-4 md:px-6 py-2 md:py-2.5 font-bold text-xs md:text-sm rounded-xl transition-all ${activeTab === 'site' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`} onClick={() => setActiveTab('site')}>Site-Wise</button>
              </div>
              <div className="relative w-full lg:flex-1 lg:max-w-xs">
                <Search className="w-4 h-4 absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" />
                <input type="text" placeholder="Search site code..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full pl-11 pr-4 py-2.5 md:py-3 bg-white border border-gray-200 rounded-2xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all shadow-sm"/>
              </div>
              <div className="flex w-full lg:w-auto gap-2">
                <button onClick={exportToExcel} className="flex-1 lg:flex-none justify-center text-xs md:text-sm font-semibold bg-emerald-50 text-emerald-700 hover:bg-emerald-100 hover:text-emerald-800 px-4 py-2.5 md:py-3 rounded-xl flex items-center gap-2 transition-colors border border-emerald-200"><Download className="w-4 h-4" /> Excel</button>
                <button onClick={exportToPDF} className="flex-1 lg:flex-none justify-center text-xs md:text-sm font-semibold bg-rose-50 text-rose-700 hover:bg-rose-100 hover:text-rose-800 px-4 py-2.5 md:py-3 rounded-xl flex items-center gap-2 transition-colors border border-rose-200"><FileText className="w-4 h-4" /> PDF</button>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[70vh] bg-gray-50 md:bg-white custom-scrollbar relative">
              <table className="w-full text-sm text-left hidden md:table">
                <thead className="bg-gray-50/90 uppercase text-[11px] font-black tracking-wider text-gray-500 sticky top-0 z-10 backdrop-blur-md shadow-sm">
                  <tr>
                    {activeTab === 'day' && <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100">Day</th>}
                    <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100">Site Code</th>
                    <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100 text-blue-800 bg-blue-50/30">Mason D</th>
                    <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100 text-purple-800 bg-purple-50/30">HM Days</th>
                    <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100 text-orange-800 bg-orange-50/30">Helper D</th>
                    <th className="px-4 lg:px-6 py-4 border-b border-r border-emerald-200 text-emerald-800 bg-emerald-50/50">Base Cost (Rs)</th>
                    <th className="px-4 lg:px-6 py-4 border-b border-r border-emerald-200 text-emerald-800 bg-emerald-50/50">OT Cost (Rs)</th>
                    <th className="px-4 lg:px-6 py-4 border-b border-gray-200 text-gray-900 bg-gray-200/50">Total Site Cost</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.length > 0 ? (
                    filteredData.map((row, idx) => (
                      <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/80 bg-white transition-colors">
                        {activeTab === 'day' && <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-900">{row.day}</td>}
                        <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-700">{row.site}</td>
                        <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.masonReg} <span className="text-xs text-gray-400">({row.masonOT}h)</span></td>
                        <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.halfMasonReg} <span className="text-xs text-gray-400">({row.halfMasonOT}h)</span></td>
                        <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.helperReg} <span className="text-xs text-gray-400">({row.helperOT}h)</span></td>
                        <td className="px-4 lg:px-6 py-4 border-r border-gray-50 text-emerald-700 font-medium">{formatCurrency(row.totalBaseCost || 0)}</td>
                        <td className="px-4 lg:px-6 py-4 border-r border-gray-100 text-emerald-700 font-medium">{formatCurrency(row.totalOTCost || 0)}</td>
                        <td className="px-4 lg:px-6 py-4 text-gray-900 font-black bg-gray-50/50">{formatCurrency((row.totalBaseCost||0) + (row.totalOTCost||0))}</td>
                      </tr>
                    ))
                  ) : (
                    <tr><td colSpan="9" className="px-6 py-12 text-center text-gray-500 font-medium">No sites found</td></tr>
                  )}
                </tbody>
                {filteredData.length > 0 && (
                  <tfoot className="bg-gray-100/90 sticky bottom-0 z-10 backdrop-blur-md shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.05)] border-t-2 border-gray-200">
                    <tr>
                      <td colSpan={activeTab === 'day' ? 2 : 1} className="px-4 lg:px-6 py-4 text-right font-black text-gray-900 uppercase">Grand Total:</td>
                      <td className="px-4 lg:px-6 py-4 font-black text-blue-900 text-base">{totals.masonReg}</td>
                      <td className="px-4 lg:px-6 py-4 font-black text-purple-900 text-base">{totals.halfMasonReg}</td>
                      <td className="px-4 lg:px-6 py-4 font-black text-orange-900 text-base">{totals.helperReg}</td>
                      <td className="px-4 lg:px-6 py-4 font-bold text-emerald-700">{formatCurrency(totals.totalBaseCost)}</td>
                      <td className="px-4 lg:px-6 py-4 font-bold text-emerald-700">{formatCurrency(totals.totalOTCost)}</td>
                      <td className="px-4 lg:px-6 py-4 font-black text-gray-900 text-lg bg-gray-200/50">{formatCurrency(totals.totalBaseCost + totals.totalOTCost)}</td>
                    </tr>
                  </tfoot>
                )}
              </table>

              {/* MOBILE CARDS FOR DATA VIEW */}
              <div className="block md:hidden p-3 space-y-4">
                {filteredData.length > 0 ? (
                  filteredData.map((row, idx) => (
                    <div key={idx} className="bg-white p-4 rounded-2xl shadow-sm border border-gray-100 space-y-3">
                      <div className="flex justify-between items-center border-b border-gray-50 pb-2">
                        <span className="text-sm font-black text-gray-800 bg-gray-100 px-3 py-1.5 rounded-lg truncate max-w-[200px]">Site: {row.site}</span>
                        {activeTab === 'day' && <span className="text-xs font-bold text-blue-700 bg-blue-50 px-3 py-1.5 rounded-lg whitespace-nowrap">Day {row.day}</span>}
                      </div>
                      <div className="grid grid-cols-3 gap-2 text-xs">
                        <div className="bg-blue-50/50 p-2.5 rounded-xl border border-blue-100"><p className="text-blue-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">Mason</p><p className="text-base font-black text-blue-900">{row.masonReg} <span className="text-[10px] font-medium text-blue-400">({row.masonOT}h)</span></p></div>
                        <div className="bg-purple-50/50 p-2.5 rounded-xl border border-purple-100"><p className="text-purple-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">H. Mason</p><p className="text-base font-black text-purple-900">{row.halfMasonReg} <span className="text-[10px] font-medium text-purple-400">({row.halfMasonOT}h)</span></p></div>
                        <div className="bg-orange-50/50 p-2.5 rounded-xl border border-orange-100"><p className="text-orange-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">Helper</p><p className="text-base font-black text-orange-900">{row.helperReg} <span className="text-[10px] font-medium text-orange-400">({row.helperOT}h)</span></p></div>
                      </div>
                      <div className="bg-emerald-50 p-3 rounded-xl border border-emerald-100 flex justify-between items-center">
                        <p className="text-emerald-700 font-bold text-[10px] uppercase tracking-wider">Total Site Cost</p>
                        <p className="text-emerald-900 font-black text-lg">{formatCurrency((row.totalBaseCost||0) + (row.totalOTCost||0))}</p>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="text-center py-12 text-gray-500 font-medium bg-white rounded-2xl border border-gray-100">No sites found</div>
                )}
                
                {filteredData.length > 0 && (
                  <div className="bg-gray-900 p-5 rounded-[2rem] shadow-xl mt-4 border border-gray-800">
                    <h3 className="text-white font-black text-center text-sm mb-4 tracking-wider">SHEET TOTALS</h3>
                    <div className="grid grid-cols-3 gap-2 text-xs mb-3">
                      <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-gray-400 font-bold mb-0.5 uppercase text-[9px]">Total Mason</p><p className="text-base font-black text-white">{totals.masonReg}</p></div>
                      <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-purple-400 font-bold mb-0.5 uppercase text-[9px]">Total HM</p><p className="text-base font-black text-purple-200">{totals.halfMasonReg}</p></div>
                      <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-gray-400 font-bold mb-0.5 uppercase text-[9px]">Total Helper</p><p className="text-base font-black text-white">{totals.helperReg}</p></div>
                    </div>
                    <div className="bg-emerald-600 p-4 rounded-2xl border border-emerald-500 text-center">
                      <p className="text-emerald-100 font-bold mb-1 uppercase text-[10px] tracking-wider">Grand Financial Total</p>
                      <p className="text-2xl font-black text-white">{formatCurrency(totals.totalBaseCost + totals.totalOTCost)}</p>
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
import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import {
  Upload, Download, FileText, UploadCloud, LayoutGrid, Search,
  Save, Trash2, Database, Clock, Check, BarChart3, Edit2,
  LogOut, Lock, Plus, X, Layers, IndianRupee, Calendar, Shield, Users, MapPin, HardHat, RefreshCw
} from 'lucide-react';
import { collection, addDoc, getDocs, deleteDoc, doc, updateDoc, setDoc } from "firebase/firestore";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { db, auth } from './firebase';

// Helper: Calculates which 15-day Bucket a date belongs to
const getPeriodKey = (dateString) => {
  const d = new Date(dateString);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = d.getDate();
  const half = day <= 15 ? 'H1' : 'H2';

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const displayMonth = monthNames[d.getMonth()];
  const displayRange = half === 'H1' ? '1-15' : '16-31';

  return {
    id: `${year}-${month}-${half}`,
    displayName: `${displayMonth} ${displayRange} ${year}`
  };
};

export default function App() {
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [activeRole, setActiveRole] = useState('master_admin');

  const [masterSites, setMasterSites] = useState([]);
  const [masterWorkers, setMasterWorkers] = useState([]);
  
  // Dynamically learn available contractors for Supervisor Dropdown
  const defaultContractors = ["Arvind", "Laljeet", "Deepak"];
  const dynamicContractors = Array.from(new Set([...defaultContractors, ...masterWorkers.map(w => w.contractor)])).filter(Boolean);

  // --- MASTER ADMIN STATES ---
  const [siteData, setSiteData] = useState([]);
  const [dayData, setDayData] = useState([]);
  const [workerData, setWorkerData] = useState([]);
  const [activeTab, setActiveTab] = useState('day');
  const [sheetName, setSheetName] = useState("");
  const [sheetMonth, setSheetMonth] = useState("");
  const [sheetContractor, setSheetContractor] = useState(""); 
  const [selectedRecordContractor, setSelectedRecordContractor] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const fileInputRef = useRef(null);
  const rosterInputRef = useRef(null);
  const [dashboardTab, setDashboardTab] = useState('upload');
  const [savedSheets, setSavedSheets] = useState([]);
  const [isSaving, setIsSaving] = useState(false);
  const [hasSavedCurrent, setHasSavedCurrent] = useState(false);
  const [isLoadingRecords, setIsLoadingRecords] = useState(false);

  // --- ANALYTICS STATES ---
  const [globalSearch, setGlobalSearch] = useState("");
  const [analyticsMode, setAnalyticsMode] = useState('single');
  const [compareSiteInput, setCompareSiteInput] = useState("");
  const [compareSitesList, setCompareSitesList] = useState([]);
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  // --- SUPERVISOR FIELD PORTAL STATES ---
  const [supDate, setSupDate] = useState(new Date().toISOString().split('T')[0]);
  const [supSite, setSupSite] = useState("");
  const [supContractor, setSupContractor] = useState("");
  const [workerSearch, setWorkerSearch] = useState("");
  const [supAttendance, setSupAttendance] = useState({});
  const [isSubmittingLog, setIsSubmittingLog] = useState(false);

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
    try { await signInWithEmailAndPassword(auth, email, password); }
    catch (err) { setLoginError("Invalid email or password. Please try again."); }
  };

  const handleLogout = async () => {
    await signOut(auth);
    goHome();
  };

  const goHome = () => {
    setSiteData([]); setDayData([]); setWorkerData([]); 
    setSelectedRecordContractor(null); setDashboardTab('upload');
  };

  const fetchData = async () => {
    if (!user) return;
    setIsLoadingRecords(true);
    try {
      const sheetsSnap = await getDocs(collection(db, "attendance_sheets"));
      const sheets = [];
      let aggregatedSites = new Set();

      sheetsSnap.forEach(docSnap => {
        const data = docSnap.data();
        sheets.push({ id: docSnap.id, ...data });
        if (data.siteData && Array.isArray(data.siteData)) {
          data.siteData.forEach(s => { if (s.site) aggregatedSites.add(s.site); });
        }
      });
      // Sort buckets descending by ID (e.g. 2026-08-H2 before 2026-08-H1)
      sheets.sort((a, b) => b.id.localeCompare(a.id));
      setSavedSheets(sheets);

      let loadedSites = Array.from(aggregatedSites);
      let loadedWorkers = [];

      try {
        const masterSnap = await getDocs(collection(db, "master_data"));
        masterSnap.forEach(docSnap => {
          if (docSnap.id === "sites" && docSnap.data().list) loadedSites = Array.from(new Set([...loadedSites, ...docSnap.data().list]));
          if (docSnap.id === "workers" && docSnap.data().list) loadedWorkers = docSnap.data().list;
        });
      } catch (err) { console.warn("master_data missing:", err); }

      setMasterSites(loadedSites.sort());
      setMasterWorkers(loadedWorkers);
    } catch (error) { console.error("Error fetching data:", error); }
    setIsLoadingRecords(false);
  };

  useEffect(() => { if (user) fetchData(); }, [user, activeRole]);

  useEffect(() => {
    if (supContractor) {
      const contractorWorkers = masterWorkers.filter(w => w.contractor === supContractor);
      const defaultState = contractorWorkers.reduce((acc, w) => {
        acc[w.name] = { status: 'absent', ot: '' };
        return acc;
      }, {});
      setSupAttendance(defaultState);
    }
  }, [supContractor, masterWorkers]);

  const formatCurrency = (amount) => {
    return new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(amount);
  };

  // --- EXCEL PARSER ---
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const rawName = file.name.replace(/\.[^/.]+$/, "");
    const firstWord = rawName.split(/[\s_.-]+/)[0];
    const detectedContractor = firstWord.charAt(0).toUpperCase() + firstWord.slice(1).toLowerCase();

    const monthNames = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };
    const mMatch = rawName.match(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i);
    const yMatch = rawName.match(/(202\d)/);
    const m = mMatch ? monthNames[mMatch[0].toLowerCase()] : String(new Date().getMonth() + 1).padStart(2, '0');
    const y = yMatch ? yMatch[0] : String(new Date().getFullYear());
    const detectedMonth = `${y}-${m}`;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      let rawData = [];

      setSheetName(rawName);
      setSheetMonth(detectedMonth);
      setSheetContractor(detectedContractor);
      setSelectedRecordContractor(null);

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

        if (col0.includes("HELPER") || col1.includes("HELPER") || col0.includes("LABOUR") || col1.includes("LABOUR")) { activeSection = "Helper"; }
        else if ((col0.includes("MASON") || col1.includes("MASON")) && !col0.includes("HALF") && !col1.includes("HALF")) { activeSection = "Mason"; }

        let rowWorkerType = activeSection;
        if (/\bHM\b/.test(col0) || /\bHM\b/.test(col1) || col1.includes("HALF MASON") || col0.includes("HALF MASON")) { rowWorkerType = "HalfMason"; }

        const col0Clean = col0.replace(/[^A-Z]/g, '');
        if (col0Clean === "SN" || col0Clean === "SNO" || col0 === "S.N.") {
          let tempHeaders = {}; wagesColIdx = -1; otWagesColIdx = -1;
          for (let c = 2; c < row.length; c++) {
            const val = String(row[c]).trim();
            if (/^\d+$/.test(val)) tempHeaders[c] = val;
            else if (val.toUpperCase() === "WAGES") wagesColIdx = c;
            else if (val.toUpperCase().includes("OT WAGES") || val.toUpperCase() === "OT (IN RS)") otWagesColIdx = c;
          }
          if (Object.keys(tempHeaders).length > 0) dayHeaders = tempHeaders;
          continue;
        }

        if (/^\d+$/.test(col0) && col1 !== "" && col1 !== "OT" && col1 !== "NAN") {
          const nextRow = rows[r + 1] || [];
          const isOTRow = String(nextRow[1]).trim().toUpperCase() === "OT";
          const workerName = String(row[1]).trim();

          let dailyWage = 0; let hourlyOTRate = 0;
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
                  rawData.push({ day: dayNum, site, type: rowWorkerType, reg: regDays, ot: otHrs, baseCost: regDays * dailyWage, otCost: otHrs * hourlyOTRate, worker: workerName, contractor: detectedContractor });
                }
              }
            }
          }
        }
      }

      const sMap = {}; const dMap = {}; const wMap = {};
      rawData.forEach(r => {
        const sKey = `${r.site}|${r.contractor}`;
        if (!sMap[sKey]) sMap[sKey] = { site: r.site, contractor: r.contractor, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalBaseCost: 0, totalOTCost: 0 };
        sMap[sKey].totalBaseCost += r.baseCost; sMap[sKey].totalOTCost += r.otCost;
        if (r.type === 'Mason') { sMap[sKey].masonReg += r.reg; sMap[sKey].masonOT += r.ot; }
        else if (r.type === 'HalfMason') { sMap[sKey].halfMasonReg += r.reg; sMap[sKey].halfMasonOT += r.ot; }
        else { sMap[sKey].helperReg += r.reg; sMap[sKey].helperOT += r.ot; }

        const dKey = `${r.day}|${r.site}|${r.contractor}`;
        if (!dMap[dKey]) dMap[dKey] = { day: r.day, site: r.site, contractor: r.contractor, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalBaseCost: 0, totalOTCost: 0 };
        dMap[dKey].totalBaseCost += r.baseCost; dMap[dKey].totalOTCost += r.otCost;
        if (r.type === 'Mason') { dMap[dKey].masonReg += r.reg; dMap[dKey].masonOT += r.ot; }
        else if (r.type === 'HalfMason') { dMap[dKey].halfMasonReg += r.reg; dMap[dKey].halfMasonOT += r.ot; }
        else { dMap[dKey].helperReg += r.reg; dMap[dKey].helperOT += r.ot; }

        const wKey = `${r.worker}|${r.contractor}`;
        if (!wMap[wKey]) wMap[wKey] = { worker: r.worker, type: r.type, contractor: r.contractor, regDays: 0, otHours: 0, totalBaseCost: 0, totalOTCost: 0 };
        wMap[wKey].regDays += r.reg; wMap[wKey].otHours += r.ot;
        wMap[wKey].totalBaseCost += r.baseCost; wMap[wKey].totalOTCost += r.otCost;
      });

      setSiteData(Object.values(sMap).sort((a, b) => a.site.localeCompare(b.site)));
      setDayData(Object.values(dMap).sort((a, b) => a.day - b.day || a.site.localeCompare(b.site)));
      setWorkerData(Object.values(wMap).sort((a, b) => a.worker.localeCompare(b.worker)));
      setHasSavedCurrent(false); setSearchQuery("");
    };
    reader.readAsBinaryString(file);
  };

  const handleMasterRosterUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const rawName = file.name.replace(/\.[^/.]+$/, "");
    const firstWord = rawName.split(/[\s_.-]+/)[0];
    const detectedContractor = firstWord.charAt(0).toUpperCase() + firstWord.slice(1).toLowerCase();

    setIsLoadingRecords(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const currentSheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(currentSheet, { header: 1, defval: "" });

      let newRosterWorkers = {};
      let activeSection = "Helper";
      let wagesColIdx = -1;

      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;
        const col0 = String(row[0]).trim().toUpperCase();
        const col1 = String(row[1]).trim().toUpperCase();
        if (col0.includes("HELPER") || col1.includes("HELPER") || col0.includes("LABOUR") || col1.includes("LABOUR")) { activeSection = "Helper"; }
        else if ((col0.includes("MASON") || col1.includes("MASON")) && !col0.includes("HALF") && !col1.includes("HALF")) { activeSection = "Mason"; }
        let rowWorkerType = activeSection;
        if (/\bHM\b/.test(col0) || /\bHM\b/.test(col1) || col1.includes("HALF MASON") || col0.includes("HALF MASON")) { rowWorkerType = "HalfMason"; }
        const col0Clean = col0.replace(/[^A-Z]/g, '');
        if (col0Clean === "SN" || col0Clean === "SNO" || col0 === "S.N.") {
          wagesColIdx = -1;
          for (let c = 0; c < row.length; c++) { if (String(row[c]).trim().toUpperCase() === "WAGES") wagesColIdx = c; }
          continue;
        }
        if (/^\d+$/.test(col0) && col1 !== "" && col1 !== "OT" && col1 !== "NAN") {
          const workerName = String(row[1]).trim();
          let dailyWage = 0;
          if (wagesColIdx !== -1 && row[wagesColIdx]) {
            const parsedWage = parseFloat(String(row[wagesColIdx]).replace(/[^0-9.]/g, ''));
            if (!isNaN(parsedWage)) dailyWage = parsedWage;
          }
          if (workerName && workerName.length > 1) {
            newRosterWorkers[`${detectedContractor}_${workerName}`] = {
              name: workerName, type: rowWorkerType, contractor: detectedContractor, wage: dailyWage || 0, otRate: (dailyWage || 0) / 8
            };
          }
        }
      }

      if (Object.keys(newRosterWorkers).length > 0) {
        try {
          let combinedWorkers = [...masterWorkers];
          combinedWorkers = combinedWorkers.filter(w => w.contractor !== detectedContractor);
          Object.values(newRosterWorkers).forEach(newW => combinedWorkers.push(newW));
          await setDoc(doc(db, "master_data", "workers"), { list: combinedWorkers });
          setMasterWorkers(combinedWorkers);
          alert(`Success! Updated ${Object.keys(newRosterWorkers).length} workers for ${detectedContractor}.`);
        } catch (err) { alert("Error saving master roster to cloud."); }
      } else { alert("No workers or wages found in this sheet."); }
      setIsLoadingRecords(false); fetchData();
    };
    reader.readAsBinaryString(file);
  };

  // SMART MERGE FOR EXCEL UPLOADS
  const saveToDatabase = async () => {
    if (hasSavedCurrent || !sheetMonth || !sheetContractor) return;
    setIsSaving(true);

    const isH2 = dayData.some(d => d.day > 15);
    const mockDate = `${sheetMonth}-${isH2 ? '16' : '01'}`;
    const period = getPeriodKey(mockDate);

    try {
      const docRef = doc(db, "attendance_sheets", period.id);
      let existingData = { sheetName: period.displayName, siteData: [], dayData: [], workerData: [], createdAt: Date.now() };
      const docSnap = await getDocs(collection(db, "attendance_sheets"));
      docSnap.forEach(d => { if (d.id === period.id) existingData = d.data(); });

      const newSiteData = (existingData.siteData || []).filter(s => s.contractor !== sheetContractor);
      const newDayData = (existingData.dayData || []).filter(d => d.contractor !== sheetContractor);
      const newWorkerData = (existingData.workerData || []).filter(w => w.contractor !== sheetContractor);

      const taggedSiteData = siteData.map(s => ({ ...s, contractor: sheetContractor }));
      const taggedDayData = dayData.map(d => ({ ...d, contractor: sheetContractor }));
      const taggedWorkerData = workerData.map(w => ({ ...w, contractor: sheetContractor }));

      const finalData = {
        sheetName: period.displayName,
        siteData: [...newSiteData, ...taggedSiteData],
        dayData: [...newDayData, ...taggedDayData],
        workerData: [...newWorkerData, ...taggedWorkerData],
        updatedAt: Date.now()
      };

      await setDoc(docRef, finalData);
      setHasSavedCurrent(true);
      fetchData();
      alert(`Successfully merged ${sheetContractor}'s data into ${period.displayName}!`);
    } catch (error) { console.error(error); alert("Failed to save to cloud bucket."); }
    setIsSaving(false);
  };

  const deleteRecord = async (id, e) => {
    e.stopPropagation();
    if (!window.confirm("Are you sure you want to permanently delete this entire 15-day period?")) return;
    try {
      await deleteDoc(doc(db, "attendance_sheets", id));
      setSavedSheets(savedSheets.filter(sheet => sheet.id !== id));
      if (hasSavedCurrent) goHome();
    } catch (error) { console.error("Error deleting:", error); }
  };

  const loadSavedRecord = (sheet) => {
    setSheetName(sheet.sheetName); setSheetMonth(sheet.id.substring(0, 7)); // e.g. "2026-08"
    setSiteData(sheet.siteData || []); setDayData(sheet.dayData || []); setWorkerData(sheet.workerData || []);
    setHasSavedCurrent(true); setSearchQuery(""); setSelectedRecordContractor(null);
  };

  // --- ANALYTICS ENGINE ---
  const getGlobalAnalytics = () => {
    if (!globalSearch.trim()) return null;
    const targetSite = globalSearch.toLowerCase().trim();
    let results = [];
    let totals = { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalSpend: 0 };

    savedSheets.forEach(sheet => {
      let sheetMason = 0, sheetMasonOT = 0, sheetHM = 0, sheetHMOT = 0, sheetHelper = 0, sheetHelperOT = 0, sheetSpend = 0;
      let sheetHasMatchingData = false;
      const sMonth = sheet.id.substring(0, 7);

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
          sheetName: sheet.sheetName, date: sheet.updatedAt || sheet.createdAt,
          masonReg: sheetMason, masonOT: sheetMasonOT, halfMasonReg: sheetHM, halfMasonOT: sheetHMOT,
          helperReg: sheetHelper, helperOT: sheetHelperOT, totalBaseCost: sheetSpend, totalOTCost: 0
        });
        totals.masonReg += sheetMason; totals.masonOT += sheetMasonOT; totals.halfMasonReg += sheetHM; totals.halfMasonOT += sheetHMOT;
        totals.helperReg += sheetHelper; totals.helperOT += sheetHelperOT; totals.totalSpend += sheetSpend;
      }
    });
    return { results, totals };
  };
  const analyticsData = getGlobalAnalytics();

  const handleAddCompareSite = (e) => {
    e.preventDefault();
    const code = compareSiteInput.trim();
    if (code && !compareSitesList.includes(code.toUpperCase())) setCompareSitesList([...compareSitesList, code.toUpperCase()]);
    setCompareSiteInput('');
  };
  const removeCompareSite = (siteToRemove) => setCompareSitesList(compareSitesList.filter(s => s !== siteToRemove));

  const getMultiSiteAnalytics = () => {
    if (compareSitesList.length === 0) return null;
    let matrix = [];
    let totals = { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalDays: 0, totalOT: 0, totalSpend: 0 };

    compareSitesList.forEach(siteCode => {
      let sTotals = { occurrences: 0, masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, siteSpend: 0 };
      savedSheets.forEach(sheet => {
        let sheetHasMatchingData = false;
        const sMonth = sheet.id.substring(0, 7);
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
      totals.masonReg += sTotals.masonReg; totals.masonOT += sTotals.masonOT;
      totals.halfMasonReg += sTotals.halfMasonReg; totals.halfMasonOT += sTotals.halfMasonOT;
      totals.helperReg += sTotals.helperReg; totals.helperOT += sTotals.helperOT;
      totals.totalDays += sTotals.masonReg + sTotals.halfMasonReg + sTotals.helperReg;
      totals.totalOT += sTotals.masonOT + sTotals.halfMasonOT + sTotals.helperOT;
      totals.totalSpend += sTotals.siteSpend;
      matrix.push({ site: siteCode, ...sTotals, totalDays: sTotals.masonReg + sTotals.halfMasonReg + sTotals.helperReg, totalOT: sTotals.masonOT + sTotals.halfMasonOT + sTotals.helperOT });
    });
    return { matrix, totals };
  };
  const multiSiteData = getMultiSiteAnalytics();

  // --- DRILL DOWN DATA FILTERS ---
  const activeData = activeTab === 'site' ? siteData : activeTab === 'day' ? dayData : workerData;
  const contractorFilteredData = selectedRecordContractor ? activeData.filter(row => row.contractor === selectedRecordContractor) : activeData;
  const filteredData = contractorFilteredData.filter(row => row.site ? row.site.toLowerCase().includes(searchQuery.toLowerCase()) : row.worker.toLowerCase().includes(searchQuery.toLowerCase()));

  const totals = filteredData.reduce((acc, row) => {
    if (activeTab === 'worker') {
      if (row.type === 'Mason') { acc.masonReg += row.regDays; acc.masonOT += row.otHours; }
      else if (row.type === 'HalfMason') { acc.halfMasonReg += row.regDays; acc.halfMasonOT += row.otHours; }
      else { acc.helperReg += row.regDays; acc.helperOT += row.otHours; }
      acc.totalBaseCost += row.totalBaseCost; acc.totalOTCost += row.totalOTCost;
    } else {
      acc.masonReg += row.masonReg || 0; acc.masonOT += row.masonOT || 0;
      acc.halfMasonReg += row.halfMasonReg || 0; acc.halfMasonOT += row.halfMasonOT || 0;
      acc.helperReg += row.helperReg || 0; acc.helperOT += row.helperOT || 0;
      acc.totalBaseCost += row.totalBaseCost || 0; acc.totalOTCost += row.totalOTCost || 0;
    }
    return acc;
  }, { masonReg: 0, masonOT: 0, halfMasonReg: 0, halfMasonOT: 0, helperReg: 0, helperOT: 0, totalBaseCost: 0, totalOTCost: 0 });

  const exportToExcel = () => { }; const exportToPDF = () => { }; 

  // --- SUPERVISOR ROSTER LOGIC & SUBMIT ---
  const handleAttendanceChange = (name, status) => {
    setSupAttendance(prev => ({ ...prev, [name]: { ...prev[name], status: status, ot: status === 'absent' ? '' : prev[name].ot } }));
  };

  const handleOTChange = (name, otValue) => {
    setSupAttendance(prev => ({ ...prev, [name]: { ...prev[name], ot: otValue } }));
  };

  const submitDailyLog = async () => {
    if (!supSite || !supContractor) return alert("Please select a site and contractor.");
    setIsSubmittingLog(true);

    const period = getPeriodKey(supDate);
    const dayNum = parseInt(supDate.split('-')[2]);

    const logWorkers = [];
    let dMasonReg = 0, dMasonOT = 0, dHMMasonReg = 0, dHMMasonOT = 0, dHelperReg = 0, dHelperOT = 0;
    let dTotalBaseCost = 0, dTotalOTCost = 0;

    Object.keys(supAttendance).forEach(name => {
      const rec = supAttendance[name];
      if (rec.status === 'present' || rec.status === 'half') {
        const workerInfo = masterWorkers.find(w => w.name === name && w.contractor === supContractor);
        if (workerInfo) {
          const isHalf = rec.status === 'half';
          const regDays = isHalf ? 0.5 : 1.0;
          const otHours = parseFloat(rec.ot) || 0;
          const bCost = regDays * (workerInfo.wage || 0);
          const oCost = otHours * (workerInfo.otRate || ((workerInfo.wage || 0) / 8));

          logWorkers.push({ worker: workerInfo.name, type: workerInfo.type, contractor: supContractor, regDays: regDays, otHours: otHours, totalBaseCost: bCost, totalOTCost: oCost });
          dTotalBaseCost += bCost; dTotalOTCost += oCost;
          if (workerInfo.type === 'Mason') { dMasonReg += regDays; dMasonOT += otHours; }
          else if (workerInfo.type === 'HalfMason') { dHMMasonReg += regDays; dHMMasonOT += otHours; }
          else { dHelperReg += regDays; dHelperOT += otHours; }
        }
      }
    });

    if (logWorkers.length === 0) { alert("No workers marked present. Nothing to save."); setIsSubmittingLog(false); return; }

    const todaySiteData = { site: supSite, contractor: supContractor, masonReg: dMasonReg, masonOT: dMasonOT, halfMasonReg: dHMMasonReg, halfMasonOT: dHMMasonOT, helperReg: dHelperReg, helperOT: dHelperOT, totalBaseCost: dTotalBaseCost, totalOTCost: dTotalOTCost };
    const todayDayData = { day: dayNum, site: supSite, contractor: supContractor, masonReg: dMasonReg, masonOT: dMasonOT, halfMasonReg: dHMMasonReg, halfMasonOT: dHMMasonOT, helperReg: dHelperReg, helperOT: dHelperOT, totalBaseCost: dTotalBaseCost, totalOTCost: dTotalOTCost };

    try {
      const docRef = doc(db, "attendance_sheets", period.id);
      let existingData = { sheetName: period.displayName, siteData: [], dayData: [], workerData: [], createdAt: Date.now() };
      const docSnap = await getDocs(collection(db, "attendance_sheets"));
      docSnap.forEach(d => { if (d.id === period.id) existingData = d.data(); });

      const safeDayData = (existingData.dayData || []).filter(d => !(d.day === dayNum && d.site === supSite && d.contractor === supContractor));

      let currentWorkerData = [...(existingData.workerData || [])];
      logWorkers.forEach(newW => {
        const existingW = currentWorkerData.find(w => w.worker === newW.worker && w.contractor === supContractor);
        if (existingW) {
          existingW.regDays += newW.regDays; existingW.otHours += newW.otHours;
          existingW.totalBaseCost += newW.totalBaseCost; existingW.totalOTCost += newW.totalOTCost;
        } else { currentWorkerData.push(newW); }
      });

      let currentSiteData = [...(existingData.siteData || [])];
      const existingSite = currentSiteData.find(s => s.site === supSite && s.contractor === supContractor);
      if (existingSite) {
        existingSite.masonReg += dMasonReg; existingSite.masonOT += dMasonOT;
        existingSite.halfMasonReg += dHMMasonReg; existingSite.halfMasonOT += dHMMasonOT;
        existingSite.helperReg += dHelperReg; existingSite.helperOT += dHelperOT;
        existingSite.totalBaseCost += dTotalBaseCost; existingSite.totalOTCost += dTotalOTCost;
      } else { currentSiteData.push(todaySiteData); }

      await setDoc(docRef, { sheetName: period.displayName, siteData: currentSiteData, dayData: [...safeDayData, todayDayData], workerData: currentWorkerData, updatedAt: Date.now() });
      await addDoc(collection(db, "daily_logs"), { date: supDate, site: supSite, contractor: supContractor, workers: logWorkers, timestamp: Date.now() });

      alert(`Attendance locked into bucket: ${period.displayName}!`);
      setSupContractor(""); setWorkerSearch(""); fetchData();
    } catch (err) { alert("Error saving attendance to bucket."); console.error(err); }
    setIsSubmittingLog(false);
  };

  // --- RENDER LOGIC ---
  if (authLoading) return <div className="min-h-screen flex items-center justify-center bg-[#f8fafc] text-gray-500 font-bold">Initializing Secure Environment...</div>;
  if (!user) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#f8fafc] px-4 font-sans selection:bg-blue-100 selection:text-blue-900">
        <div className="w-full max-w-md bg-white p-8 md:p-10 rounded-[2rem] shadow-xl border border-gray-100">
          <div className="flex flex-col items-center mb-8 text-center"><div className="bg-blue-600 p-4 rounded-2xl shadow-lg shadow-blue-600/20 mb-5"><Lock className="w-8 h-8 text-white" /></div><h1 className="text-3xl font-black text-gray-900 tracking-tight">CRM_FIX</h1><p className="text-sm font-bold text-gray-400 uppercase tracking-widest mt-1">Authorized Access Only</p></div>
          <form onSubmit={handleLogin} className="space-y-5">
            <div><label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Email</label><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" /></div>
            <div><label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" /></div>
            {loginError && <p className="text-red-500 text-xs font-bold text-center bg-red-50 py-2 rounded-lg">{loginError}</p>}
            <button type="submit" className="w-full py-4 bg-gray-900 hover:bg-blue-600 text-white font-bold rounded-xl shadow-md transition-colors duration-300">Access Dashboard</button>
          </form>
        </div>
      </div>
    );
  }

  const activeContractorWorkers = masterWorkers.filter(w => w.contractor === supContractor);
  const searchedWorkers = activeContractorWorkers.filter(w => w.name.toLowerCase().includes(workerSearch.toLowerCase()));
  const masons = searchedWorkers.filter(w => w.type === 'Mason');
  const halfMasons = searchedWorkers.filter(w => w.type === 'HalfMason');
  const helpers = searchedWorkers.filter(w => w.type === 'Helper');

  return (
    <div className="min-h-screen bg-[#f8fafc] text-gray-800 font-sans selection:bg-blue-100 selection:text-blue-900 pb-12">

      {/* GLOBAL HEADER */}
      <header className="bg-white border-b border-gray-200 px-4 py-3 flex justify-between items-center shadow-sm sticky top-0 z-50">
        <div className="flex items-center gap-2 cursor-pointer hover:opacity-80 transition-opacity" onClick={goHome} title="Go to Dashboard">
          <div className={`p-1.5 rounded-lg ${activeRole === 'master_admin' ? 'bg-blue-600' : 'bg-emerald-600'}`}><LayoutGrid className="w-4 h-4 text-white" /></div>
          <span className="font-black text-gray-900 text-sm tracking-tight">CRM_FIX <span className="text-gray-400 font-medium hidden sm:inline">| {activeRole === 'master_admin' ? 'Master' : 'Field'}</span></span>
        </div>
        <div className="flex items-center gap-3 md:gap-4">
          <button onClick={fetchData} disabled={isLoadingRecords} className="p-2 bg-gray-100 hover:bg-blue-50 text-gray-600 hover:text-blue-600 rounded-xl transition-all" title="Sync All Records">
            <RefreshCw className={`w-3.5 h-3.5 ${isLoadingRecords ? 'animate-spin text-blue-500' : ''}`} />
          </button>
          <div className="flex items-center bg-gray-100 p-1 rounded-xl shadow-inner">
            <button onClick={() => setActiveRole('master_admin')} className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1 ${activeRole === 'master_admin' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500'}`}><Shield className="w-3.5 h-3.5" /> Admin</button>
            <button onClick={() => setActiveRole('admin')} className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1 ${activeRole === 'admin' ? 'bg-white text-emerald-600 shadow-sm' : 'text-gray-500'}`}><HardHat className="w-3.5 h-3.5" /> Supervisor</button>
          </div>
          <button onClick={handleLogout} className="text-xs font-bold bg-gray-100 hover:bg-red-50 hover:text-red-600 text-gray-600 p-2 md:px-3 md:py-1.5 rounded-xl flex items-center gap-2 transition-colors"><LogOut className="w-4 h-4 md:w-3.5 md:h-3.5" /> <span className="hidden md:block">Sign Out</span></button>
        </div>
      </header>

      {/* SUPERVISOR FIELD PORTAL */}
      {activeRole === 'admin' ? (
        <div className="max-w-3xl mx-auto pt-8 md:pt-12 px-4 space-y-6">
          <div className="text-center mb-8">
            <h1 className="text-3xl font-black tracking-tight text-gray-900 mb-2">Hello Team,</h1>
            <p className="text-sm font-bold text-gray-400 uppercase tracking-widest">Daily Field Attendance</p>
          </div>

          <div className="bg-white rounded-[2rem] p-6 md:p-8 shadow-sm border border-gray-100 space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1 flex items-center gap-1"><Calendar className="w-3 h-3" /> Work Date</label>
                <input type="date" value={supDate} onChange={(e) => setSupDate(e.target.value)} className="w-full bg-gray-50 border border-gray-200 px-4 py-3.5 rounded-xl text-sm font-bold text-gray-800 outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all" />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1 flex items-center gap-1"><MapPin className="w-3 h-3" /> Select Site</label>
                <select value={supSite} onChange={(e) => setSupSite(e.target.value)} className="w-full bg-gray-50 border border-gray-200 px-4 py-3.5 rounded-xl text-sm font-bold text-gray-800 outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 appearance-none cursor-pointer">
                  <option value="" disabled>Choose a site...</option>
                  {masterSites.map(site => <option key={site} value={site}>{site}</option>)}
                </select>
              </div>
            </div>

            <div className="space-y-2 pt-2 border-t border-gray-50">
              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1 flex items-center gap-1"><Users className="w-3 h-3" /> Select Contractor</label>
              <select value={supContractor} onChange={(e) => setSupContractor(e.target.value)} className="w-full bg-gray-50 border border-gray-200 px-4 py-3.5 rounded-xl text-sm font-bold text-gray-800 outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 appearance-none cursor-pointer">
                <option value="" disabled>Choose contractor crew...</option>
                {dynamicContractors.map(contractor => <option key={contractor} value={contractor}>{contractor}'s Team</option>)}
              </select>
            </div>

            {supContractor && (
              <div className="pt-6 border-t border-gray-100">
                <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
                  <div>
                    <h3 className="font-black text-gray-900 text-lg">Daily Roster <span className="text-gray-400 font-medium text-sm">({activeContractorWorkers.length} total)</span></h3>
                    <span className="bg-gray-100 text-gray-500 text-[9px] font-bold px-2 py-1 rounded-md uppercase tracking-wider mt-1 inline-block">All Defaulted to Absent</span>
                  </div>
                  <div className="relative w-full md:w-64">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input type="text" placeholder="Search worker..." value={workerSearch} onChange={(e) => setWorkerSearch(e.target.value)} className="w-full pl-9 pr-4 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:bg-white transition-all" />
                  </div>
                </div>

                <div className="space-y-8">
                  {masons.length > 0 && (
                    <div className="space-y-3">
                      <h4 className="text-[11px] font-black text-blue-500 uppercase tracking-widest border-b border-gray-100 pb-2 flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-blue-500"></div> MASONS</h4>
                      {masons.map(worker => {
                        const rec = supAttendance[worker.name] || {};
                        const isPresent = rec.status === 'present'; const isHalf = rec.status === 'half'; const isAbsent = rec.status === 'absent';
                        return (
                          <div key={worker.name} className={`p-4 rounded-2xl border transition-all ${isAbsent ? 'bg-gray-50/50 border-gray-100' : 'bg-white border-blue-200 shadow-sm'}`}>
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                              <p className="font-black text-gray-900 text-base">{worker.name}</p>
                              <div className="flex flex-wrap items-center gap-2">
                                <div className="flex bg-gray-100 p-1 rounded-xl">
                                  <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isPresent ? 'bg-blue-600 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                                  <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isHalf ? 'bg-yellow-500 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                                  <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isAbsent ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                                </div>
                                {(isPresent || isHalf) && <input type="number" placeholder="OT hrs" value={rec.ot} onChange={(e) => handleOTChange(worker.name, e.target.value)} className="w-20 bg-gray-50 border border-gray-200 px-3 py-2 rounded-xl text-xs font-bold text-gray-800 outline-none focus:ring-2 focus:ring-blue-500/20 text-center" />}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {halfMasons.length > 0 && (
                    <div className="space-y-3">
                      <h4 className="text-[11px] font-black text-purple-500 uppercase tracking-widest border-b border-gray-100 pb-2 flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-purple-500"></div> HALF MASONS</h4>
                      {halfMasons.map(worker => {
                        const rec = supAttendance[worker.name] || {};
                        const isPresent = rec.status === 'present'; const isHalf = rec.status === 'half'; const isAbsent = rec.status === 'absent';
                        return (
                          <div key={worker.name} className={`p-4 rounded-2xl border transition-all ${isAbsent ? 'bg-gray-50/50 border-gray-100' : 'bg-white border-purple-200 shadow-sm'}`}>
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                              <p className="font-black text-gray-900 text-base">{worker.name}</p>
                              <div className="flex flex-wrap items-center gap-2">
                                <div className="flex bg-gray-100 p-1 rounded-xl">
                                  <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isPresent ? 'bg-purple-600 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                                  <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isHalf ? 'bg-yellow-500 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                                  <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isAbsent ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                                </div>
                                {(isPresent || isHalf) && <input type="number" placeholder="OT hrs" value={rec.ot} onChange={(e) => handleOTChange(worker.name, e.target.value)} className="w-20 bg-gray-50 border border-gray-200 px-3 py-2 rounded-xl text-xs font-bold text-gray-800 outline-none focus:ring-2 focus:ring-purple-500/20 text-center" />}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {helpers.length > 0 && (
                    <div className="space-y-3">
                      <h4 className="text-[11px] font-black text-orange-500 uppercase tracking-widest border-b border-gray-100 pb-2 flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-orange-500"></div> HELPERS</h4>
                      {helpers.map(worker => {
                        const rec = supAttendance[worker.name] || {};
                        const isPresent = rec.status === 'present'; const isHalf = rec.status === 'half'; const isAbsent = rec.status === 'absent';
                        return (
                          <div key={worker.name} className={`p-4 rounded-2xl border transition-all ${isAbsent ? 'bg-gray-50/50 border-gray-100' : 'bg-white border-orange-200 shadow-sm'}`}>
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                              <p className="font-black text-gray-900 text-base">{worker.name}</p>
                              <div className="flex flex-wrap items-center gap-2">
                                <div className="flex bg-gray-100 p-1 rounded-xl">
                                  <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isPresent ? 'bg-orange-600 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                                  <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isHalf ? 'bg-yellow-500 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                                  <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${isAbsent ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                                </div>
                                {(isPresent || isHalf) && <input type="number" placeholder="OT hrs" value={rec.ot} onChange={(e) => handleOTChange(worker.name, e.target.value)} className="w-20 bg-gray-50 border border-gray-200 px-3 py-2 rounded-xl text-xs font-bold text-gray-800 outline-none focus:ring-2 focus:ring-orange-500/20 text-center" />}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {activeContractorWorkers.length === 0 && (
                    <div className="text-center py-12 text-gray-400 font-bold bg-gray-50 rounded-2xl border border-gray-100">
                      No workers found for this contractor.<br />Upload an Excel sheet first to automatically learn their names.
                    </div>
                  )}
                </div>

                <button onClick={submitDailyLog} disabled={isSubmittingLog || activeContractorWorkers.length === 0} className={`w-full mt-10 font-black py-4 rounded-xl flex items-center justify-center gap-2 shadow-lg transition-colors ${isSubmittingLog || activeContractorWorkers.length === 0 ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : 'bg-gray-900 hover:bg-emerald-600 text-white'}`}>
                  <Save className="w-5 h-5" /> {isSubmittingLog ? 'Saving...' : "Secure Today's Attendance"}
                </button>
              </div>
            )}
          </div>
        </div>
      ) : (

        /* ============================================================== */
        /*                 MASTER ADMIN DASHBOARD (EXCEL)                 */
        /* ============================================================== */
        <>
          <input type="file" accept=".xlsx, .xls, .csv" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
          {activeData.length === 0 ? (
            <div className="max-w-6xl mx-auto pt-10 md:pt-16 px-4">
              <div className="text-center mb-10 space-y-3">
                <h1 className="text-3xl md:text-5xl font-black tracking-tight text-gray-900">Welcome Back.</h1>
                <p className="text-sm md:text-base font-bold text-gray-400 uppercase tracking-widest">Financial & Attendance Engine</p>
              </div>

              <div className="flex justify-center mb-12">
                <div className="bg-gray-100 p-1.5 rounded-full inline-flex shadow-inner overflow-x-auto max-w-full custom-scrollbar">
                  <button onClick={() => setDashboardTab('upload')} className={`whitespace-nowrap px-6 py-3 rounded-full font-bold text-sm transition-all flex items-center gap-2.5 ${dashboardTab === 'upload' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                    <UploadCloud className={`w-4 h-4 ${dashboardTab === 'upload' ? 'text-gray-900' : 'text-gray-400'}`} /> Upload New
                  </button>
                  <button onClick={() => { setDashboardTab('records'); fetchData(); }} className={`whitespace-nowrap px-6 py-3 rounded-full font-bold text-sm transition-all flex items-center gap-2.5 ${dashboardTab === 'records' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                    <Database className={`w-4 h-4 ${dashboardTab === 'records' ? 'text-gray-900' : 'text-gray-400'}`} /> Saved Records
                  </button>
                  <button onClick={() => { setDashboardTab('analytics'); fetchData(); }} className={`whitespace-nowrap px-6 py-3 rounded-full font-bold text-sm transition-all flex items-center gap-2.5 ${dashboardTab === 'analytics' ? 'bg-white text-blue-600 shadow-sm border border-gray-50' : 'text-gray-500 hover:text-gray-700'}`}>
                    <BarChart3 className={`w-4 h-4 ${dashboardTab === 'analytics' ? 'text-blue-600' : 'text-gray-400'}`} /> Global Analytics
                  </button>
                </div>
              </div>

              {/* TAB 1: UPLOAD ZONE */}
              {dashboardTab === 'upload' && (
                <div className="max-w-2xl mx-auto space-y-4">
                  <div onClick={() => fileInputRef.current?.click()} className="group w-full cursor-pointer bg-white rounded-[2.5rem] border-2 border-dashed border-gray-200 hover:border-blue-500 hover:bg-blue-50/50 transition-all duration-300 p-8 md:p-16 shadow-sm hover:shadow-xl text-center">
                    <div className="flex flex-col items-center justify-center space-y-6">
                      <div className="bg-blue-50 text-blue-600 p-6 rounded-full group-hover:scale-110 group-hover:bg-blue-100 transition-all duration-300"><UploadCloud className="w-12 h-12" /></div>
                      <div><p className="text-xl md:text-2xl font-black text-gray-800">Upload Financial Sheet</p><p className="text-sm md:text-base font-medium text-gray-500 mt-2">Parse attendance and generate site analytics instantly.</p></div>
                      <span className="mt-4 px-8 py-3.5 bg-gray-900 text-white font-bold rounded-2xl shadow-md group-hover:bg-blue-600 transition-colors duration-300 flex items-center gap-2"><Upload className="w-4 h-4" />Select Data File</span>
                    </div>
                  </div>

                  <input type="file" accept=".xlsx, .xls, .csv" className="hidden" ref={rosterInputRef} onChange={handleMasterRosterUpload} />
                  <div onClick={() => rosterInputRef.current?.click()} className="group w-full cursor-pointer bg-emerald-50 rounded-3xl border border-emerald-100 hover:border-emerald-500 hover:bg-emerald-100/50 transition-all duration-300 p-6 shadow-sm text-center flex items-center justify-center gap-4">
                    <div className="bg-white text-emerald-600 p-3 rounded-xl shadow-sm group-hover:scale-110 transition-all duration-300"><Users className="w-6 h-6" /></div>
                    <div className="text-left"><p className="text-base font-black text-emerald-900">Sync Master Roster</p><p className="text-xs font-bold text-emerald-600 mt-0.5">Upload a sheet to strictly update worker wages & names</p></div>
                  </div>
                </div>
              )}

              {/* TAB 2: SAVED RECORDS */}
              {dashboardTab === 'records' && (
                <div className="bg-white rounded-[2.5rem] p-6 md:p-10 shadow-sm border border-gray-100 min-h-[400px]">
                  {isLoadingRecords ? (
                    <div className="flex justify-center items-center h-48 text-gray-400 font-bold">Loading cloud records...</div>
                  ) : savedSheets.length === 0 ? (
                    <div className="flex flex-col items-center justify-center h-48 text-gray-400 space-y-4"><Database className="w-12 h-12 text-gray-200" /><p className="font-semibold text-center">No saved buckets found.<br /><span className="text-sm font-normal">Upload and save a sheet to see it here.</span></p></div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
                      {savedSheets.map((sheet) => (
                        <div key={sheet.id} onClick={() => loadSavedRecord(sheet)} className="group bg-white hover:bg-blue-50/50 border border-gray-100 hover:border-blue-200 rounded-[1.5rem] p-6 cursor-pointer transition-all flex flex-col justify-between space-y-5 shadow-[0_2px_10px_-3px_rgba(0,0,0,0.05)] hover:shadow-lg">
                          <div className="flex items-start justify-between">
                            <div className="bg-blue-50 text-blue-600 p-3.5 rounded-2xl"><FileText className="w-5 h-5" /></div>
                            <div className="flex gap-1.5">
                              <button onClick={(e) => deleteRecord(sheet.id, e)} className="text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-white shadow-sm border border-transparent hover:border-red-100 rounded-xl p-2 transition-all" title="Delete Period"><Trash2 className="w-4 h-4" /></button>
                            </div>
                          </div>
                          <div>
                            <h3 className="font-black text-gray-800 text-base md:text-lg group-hover:text-blue-700" title={sheet.sheetName}>{sheet.sheetName}</h3>
                            <div className="flex items-center justify-between mt-2">
                              <span className="bg-gray-100 text-gray-500 text-[10px] font-bold px-2 py-1 rounded-md">{sheet.id}</span>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* TAB 3: GLOBAL ANALYTICS */}
              {dashboardTab === 'analytics' && (
                <div className="max-w-4xl mx-auto space-y-6">
                  <div className="bg-white px-6 py-10 md:p-12 rounded-[2.5rem] shadow-[0_2px_15px_-3px_rgba(0,0,0,0.07),0_10px_20px_-2px_rgba(0,0,0,0.04)] border border-gray-100 text-center">
                    <h2 className="text-xl md:text-2xl font-black text-gray-900 mb-8 tracking-tight">Master Site Analytics</h2>

                    <div className="flex justify-center mb-8">
                      <div className="bg-gray-50 border border-gray-100 p-1.5 rounded-2xl inline-flex shadow-inner">
                        <button onClick={() => setAnalyticsMode('single')} className={`px-6 py-2.5 rounded-xl font-bold text-sm transition-all flex items-center gap-2 ${analyticsMode === 'single' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                          <Search className={`w-4 h-4 ${analyticsMode === 'single' ? 'text-blue-600' : 'text-gray-400'}`} /> Single Site Lookup
                        </button>
                        <button onClick={() => setAnalyticsMode('compare')} className={`px-6 py-2.5 rounded-xl font-bold text-sm transition-all flex items-center gap-2 ${analyticsMode === 'compare' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                          <Layers className={`w-4 h-4 ${analyticsMode === 'compare' ? 'text-blue-600' : 'text-gray-400'}`} /> Multi-Site Compare
                        </button>
                      </div>
                    </div>

                    {analyticsMode === 'single' ? (
                      <div className="relative w-full max-w-md mx-auto mb-8">
                        <Search className="w-5 h-5 absolute left-5 top-1/2 -translate-y-1/2 text-gray-400" />
                        <input type="text" placeholder="Enter site code..." value={globalSearch} onChange={(e) => setGlobalSearch(e.target.value)} className="w-full pl-14 pr-4 py-4 bg-gray-50/80 border border-gray-200 rounded-2xl text-base font-bold focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:bg-white focus:border-blue-500 transition-all shadow-sm" />
                      </div>
                    ) : (
                      <div className="w-full max-w-xl mx-auto mb-8">
                        <form onSubmit={handleAddCompareSite} className="flex flex-col sm:flex-row items-center gap-3">
                          <div className="relative flex-1 w-full">
                            <Plus className="w-5 h-5 absolute left-5 top-1/2 -translate-y-1/2 text-gray-400" />
                            <input type="text" placeholder="Type site code (e.g. S01) & press Enter..." value={compareSiteInput} onChange={(e) => setCompareSiteInput(e.target.value)} className="w-full pl-14 pr-4 py-4 bg-gray-50/80 border border-gray-200 rounded-2xl text-base font-bold focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:bg-white focus:border-blue-500 transition-all shadow-sm" />
                          </div>
                          <button type="submit" className="w-full sm:w-auto bg-blue-600 hover:bg-blue-700 text-white px-8 py-4 rounded-2xl font-black flex items-center justify-center gap-2 transition-colors shadow-md">Add</button>
                        </form>
                        {compareSitesList.length > 0 && (
                          <div className="flex flex-wrap gap-2 justify-center mt-5">
                            {compareSitesList.map(site => (
                              <div key={site} className="bg-blue-50 border border-blue-200 text-blue-800 px-4 py-2.5 rounded-xl text-sm font-black flex items-center gap-2 shadow-sm">
                                {site}
                                <button onClick={() => removeCompareSite(site)} className="text-blue-400 hover:text-red-500 transition-colors bg-white rounded-md p-0.5"><X className="w-4 h-4" /></button>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    )}

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
                        <button onClick={() => { setStartDate(""); setEndDate(""); }} className="p-4 bg-red-50 hover:bg-red-100 text-red-500 rounded-[1.25rem] transition-colors border border-red-100 shrink-0 shadow-sm" title="Clear Dates"><X className="w-5 h-5" /></button>
                      )}
                    </div>
                  </div>

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
                                      <p className="font-black text-emerald-900 text-sm">{formatCurrency((record.totalBaseCost || 0) + (record.totalOTCost || 0))}</p>
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

                  {analyticsMode === 'compare' && compareSitesList.length > 0 && multiSiteData && (
                    <div className="space-y-6">
                      <div className="bg-white rounded-3xl shadow-sm border border-gray-200 overflow-hidden">
                        <div className="bg-gray-50 border-b border-gray-100 p-4 font-bold text-gray-600 text-xs md:text-sm uppercase tracking-wider">Comparison Matrix <span className="text-gray-400 normal-case font-medium">({compareSitesList.length} sites)</span></div>
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
            <div className="max-w-7xl mx-auto space-y-4 md:space-y-6 p-2 md:p-8 pt-4 md:pt-8 pb-12">
              <div className="bg-white p-4 md:p-5 rounded-[2rem] shadow-sm border border-gray-100 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div className="flex flex-col md:flex-row gap-2 md:gap-4 items-start md:items-center w-full md:w-auto">
                  <div className="min-w-0 flex-1">
                    <h1 className="text-xl md:text-2xl font-black tracking-tight text-gray-900 leading-none truncate">{sheetName}</h1>
                    <p className="text-[10px] md:text-xs font-bold uppercase tracking-wider text-emerald-500 mt-1 flex items-center gap-1"><IndianRupee className="w-3 h-3" /> Financial Engine Active</p>
                  </div>
                </div>

                <div className="flex w-full md:w-auto gap-2 md:gap-3">
                  <button onClick={saveToDatabase} disabled={hasSavedCurrent || isSaving || !sheetContractor || sheetContractor === "Unknown"} className={`flex-1 justify-center px-4 md:px-6 py-3 rounded-xl font-black text-sm flex items-center gap-2 transition-all shadow-sm ${hasSavedCurrent || !sheetContractor || sheetContractor === "Unknown" ? 'bg-gray-100 text-gray-500 cursor-not-allowed border border-gray-200' : 'bg-indigo-600 hover:bg-indigo-700 text-white'}`}>
                    {hasSavedCurrent ? <Check className="w-4 h-4" /> : <Save className="w-4 h-4" />}
                    {hasSavedCurrent ? "Saved" : isSaving ? "Saving..." : "Save Insights"}
                  </button>
                  <button onClick={() => { goHome(); fileInputRef.current?.click(); }} className="flex-1 justify-center bg-gray-100 hover:bg-gray-200 text-gray-700 px-4 md:px-5 py-3 rounded-xl font-black text-sm flex items-center gap-2 transition-all">
                    <Upload className="w-4 h-4" /> Upload
                  </button>
                </div>
              </div>

              {/* UNIFIED BUCKET ERROR HANDLER */}
              {!sheetContractor || sheetContractor === "Unknown" ? (
                <div className="bg-red-50 border border-red-200 rounded-[2rem] p-8 text-center mt-6">
                  <h3 className="text-xl font-black text-red-700 mb-2">Unknown Contractor</h3>
                  <p className="text-sm font-bold text-red-500">The uploaded file name must include the contractor's name (e.g., "Arvind May 1-15.xlsx") to save it to the bucket.</p>
                </div>
              ) : null}

              {/* NEW DRILL-DOWN LOGIC: The Contractor Boxes */}
              {!selectedRecordContractor ? (
                <div className="bg-white rounded-[2.5rem] shadow-sm border border-gray-100 p-8 md:p-12 text-center max-w-4xl mx-auto mt-8">
                  <h2 className="text-2xl md:text-3xl font-black text-gray-900 mb-2">Select Contractor Team</h2>
                  <p className="text-sm font-bold text-gray-400 uppercase tracking-widest mb-10">Data Period: {sheetName}</p>
                  
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 md:gap-6">
                    {dynamicContractors.map(c => {
                      const hasData = workerData.some(w => w.contractor === c);
                      return (
                        <div key={c} onClick={() => hasData && setSelectedRecordContractor(c)} className={`group relative p-6 md:p-8 rounded-[2rem] border-2 transition-all ${hasData ? 'bg-white hover:bg-blue-50/50 border-gray-100 hover:border-blue-200 cursor-pointer shadow-sm hover:shadow-xl' : 'bg-gray-50 border-dashed border-gray-200 opacity-50 cursor-not-allowed'}`}>
                          <div className={`w-12 h-12 mx-auto rounded-full flex items-center justify-center mb-4 transition-transform ${hasData ? 'bg-blue-100 text-blue-600 group-hover:scale-110' : 'bg-gray-200 text-gray-400'}`}>
                            <Users className="w-6 h-6" />
                          </div>
                          <h3 className="text-lg font-black text-gray-900">{c}</h3>
                          <p className="text-[10px] font-bold uppercase tracking-widest mt-1 text-gray-400">{hasData ? 'View Analytics' : 'No Data'}</p>
                        </div>
                      )
                    })}
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded-[2rem] shadow-sm border border-gray-100 overflow-hidden mt-6">
                  {/* Internal Controls for the selected team */}
                  <div className="flex flex-col lg:flex-row justify-between items-stretch lg:items-center border-b border-gray-100 bg-gray-50/50 p-2 md:p-4 gap-3">
                    <div className="flex items-center gap-3 w-full lg:w-auto">
                      <button onClick={() => setSelectedRecordContractor(null)} className="p-2.5 bg-white border border-gray-200 rounded-xl text-gray-500 hover:text-blue-600 hover:border-blue-200 shadow-sm transition-all" title="Back to Teams">
                        <LayoutGrid className="w-4 h-4" />
                      </button>
                      <div className="flex w-full lg:w-auto bg-gray-200/60 p-1 rounded-[1.25rem]">
                        <button className={`flex-1 px-4 md:px-6 py-2 md:py-2.5 font-bold text-xs md:text-sm rounded-xl transition-all ${activeTab === 'day' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`} onClick={() => setActiveTab('day')}>Day-Wise</button>
                        <button className={`flex-1 px-4 md:px-6 py-2 md:py-2.5 font-bold text-xs md:text-sm rounded-xl transition-all ${activeTab === 'site' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`} onClick={() => setActiveTab('site')}>Site-Wise</button>
                        <button className={`flex-1 px-4 md:px-6 py-2 md:py-2.5 font-bold text-xs md:text-sm rounded-xl transition-all ${activeTab === 'worker' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`} onClick={() => setActiveTab('worker')}>Worker-Wise</button>
                      </div>
                    </div>
                    <div className="relative w-full lg:flex-1 lg:max-w-xs">
                      <Search className="w-4 h-4 absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input type="text" placeholder="Search site code..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full pl-11 pr-4 py-2.5 md:py-3 bg-white border border-gray-200 rounded-2xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all shadow-sm" />
                    </div>
                    <div className="flex w-full lg:w-auto gap-2">
                      <button onClick={exportToExcel} className="flex-1 lg:flex-none justify-center text-xs md:text-sm font-semibold bg-emerald-50 text-emerald-700 hover:bg-emerald-100 hover:text-emerald-800 px-4 py-2.5 md:py-3 rounded-xl flex items-center gap-2 transition-colors border border-emerald-200"><Download className="w-4 h-4" /> Excel</button>
                      <button onClick={exportToPDF} className="flex-1 lg:flex-none justify-center text-xs md:text-sm font-semibold bg-rose-50 text-rose-700 hover:bg-rose-100 hover:text-rose-800 px-4 py-2.5 md:py-3 rounded-xl flex items-center gap-2 transition-colors border border-rose-200"><FileText className="w-4 h-4" /> PDF</button>
                    </div>
                  </div>

                  <div className="overflow-x-auto max-h-[70vh] bg-gray-50 md:bg-white custom-scrollbar relative">
                    <table className="w-full text-sm text-left hidden md:table">
                      <thead className="bg-gray-50/90 uppercase text-[11px] font-black tracking-wider text-gray-500 sticky top-0 z-10 backdrop-blur-md shadow-sm">
                        {activeTab === 'worker' ? (
                          <tr>
                            <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100">Worker Name</th>
                            <th className="px-4 lg:px-6 py-4 border-b border-r border-gray-100">Category</th>
                            <th className="px-4 lg:px-6 py-4 border-b border-r border-blue-100 text-blue-800 bg-blue-50/30">Total Days</th>
                            <th className="px-4 lg:px-6 py-4 border-b border-r border-purple-100 text-purple-800 bg-purple-50/30">Total OT (Hrs)</th>
                            <th className="px-4 lg:px-6 py-4 border-b border-r border-emerald-200 text-emerald-800 bg-emerald-50/50">Base Pay</th>
                            <th className="px-4 lg:px-6 py-4 border-b border-r border-emerald-200 text-emerald-800 bg-emerald-50/50">OT Pay</th>
                            <th className="px-4 lg:px-6 py-4 border-b border-gray-200 text-gray-900 bg-gray-200/50">Total Payout</th>
                          </tr>
                        ) : (
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
                        )}
                      </thead>
                      <tbody>
                        {filteredData.length > 0 ? (
                          activeTab === 'worker' ? (
                            filteredData.map((row, idx) => (
                              <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/80 bg-white transition-colors">
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-900">{row.worker}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-700">{row.type}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.regDays}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-purple-600">{row.otHours}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 text-emerald-700 font-medium">{formatCurrency(row.totalBaseCost || 0)}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-100 text-emerald-700 font-medium">{formatCurrency(row.totalOTCost || 0)}</td>
                                <td className="px-4 lg:px-6 py-4 text-gray-900 font-black bg-gray-50/50">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</td>
                              </tr>
                            ))
                          ) : (
                            filteredData.map((row, idx) => (
                              <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/80 bg-white transition-colors">
                                {activeTab === 'day' && <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-900">{row.day}</td>}
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-700">{row.site}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.masonReg} <span className="text-xs text-gray-400">({row.masonOT}h)</span></td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.halfMasonReg} <span className="text-xs text-gray-400">({row.halfMasonOT}h)</span></td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.helperReg} <span className="text-xs text-gray-400">({row.helperOT}h)</span></td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-50 text-emerald-700 font-medium">{formatCurrency(row.totalBaseCost || 0)}</td>
                                <td className="px-4 lg:px-6 py-4 border-r border-gray-100 text-emerald-700 font-medium">{formatCurrency(row.totalOTCost || 0)}</td>
                                <td className="px-4 lg:px-6 py-4 text-gray-900 font-black bg-gray-50/50">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</td>
                              </tr>
                            ))
                          )
                        ) : (
                          <tr><td colSpan="9" className="px-6 py-12 text-center text-gray-500 font-medium">No records found</td></tr>
                        )}
                      </tbody>
                      {filteredData.length > 0 && (
                        <tfoot className="bg-gray-100/90 sticky bottom-0 z-10 backdrop-blur-md shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.05)] border-t-2 border-gray-200">
                          {activeTab === 'worker' ? (
                            <tr>
                              <td colSpan="2" className="px-4 lg:px-6 py-4 text-right font-black text-gray-900 uppercase">Grand Total:</td>
                              <td className="px-4 lg:px-6 py-4 font-black text-blue-900 text-base">{totals.masonReg + totals.halfMasonReg + totals.helperReg}</td>
                              <td className="px-4 lg:px-6 py-4 font-black text-purple-900 text-base">{totals.masonOT + totals.halfMasonOT + totals.helperOT} h</td>
                              <td className="px-4 lg:px-6 py-4 font-bold text-emerald-700">{formatCurrency(totals.totalBaseCost)}</td>
                              <td className="px-4 lg:px-6 py-4 font-bold text-emerald-700">{formatCurrency(totals.totalOTCost)}</td>
                              <td className="px-4 lg:px-6 py-4 font-black text-gray-900 text-lg bg-gray-200/50">{formatCurrency(totals.totalBaseCost + totals.totalOTCost)}</td>
                            </tr>
                          ) : (
                            <tr>
                              <td colSpan={activeTab === 'day' ? 2 : 1} className="px-4 lg:px-6 py-4 text-right font-black text-gray-900 uppercase">Grand Total:</td>
                              <td className="px-4 lg:px-6 py-4 font-black text-blue-900 text-base">{totals.masonReg} <span className="text-xs text-gray-500">({totals.masonOT}h)</span></td>
                              <td className="px-4 lg:px-6 py-4 font-black text-purple-900 text-base">{totals.halfMasonReg} <span className="text-xs text-gray-500">({totals.halfMasonOT}h)</span></td>
                              <td className="px-4 lg:px-6 py-4 font-black text-orange-900 text-base">{totals.helperReg} <span className="text-xs text-gray-500">({totals.helperOT}h)</span></td>
                              <td className="px-4 lg:px-6 py-4 font-bold text-emerald-700">{formatCurrency(totals.totalBaseCost)}</td>
                              <td className="px-4 lg:px-6 py-4 font-bold text-emerald-700">{formatCurrency(totals.totalOTCost)}</td>
                              <td className="px-4 lg:px-6 py-4 font-black text-gray-900 text-lg bg-gray-200/50">{formatCurrency(totals.totalBaseCost + totals.totalOTCost)}</td>
                            </tr>
                          )}
                        </tfoot>
                      )}
                    </table>

                    {/* MOBILE CARDS */}
                    <div className="block md:hidden p-3 space-y-4">
                      {filteredData.length > 0 ? (
                        filteredData.map((row, idx) => (
                          <div key={idx} className="bg-white p-4 rounded-2xl shadow-sm border border-gray-100 space-y-3">
                            <div className="flex justify-between items-center border-b border-gray-50 pb-2">
                              <span className="text-sm font-black text-gray-800 bg-gray-100 px-3 py-1.5 rounded-lg truncate max-w-[200px]">{activeTab === 'worker' ? `Name: ${row.worker}` : `Site: ${row.site}`}</span>
                              {activeTab === 'day' && <span className="text-xs font-bold text-blue-700 bg-blue-50 px-3 py-1.5 rounded-lg whitespace-nowrap">Day {row.day}</span>}
                            </div>
                            
                            {activeTab === 'worker' ? (
                              <div className="grid grid-cols-2 gap-2 text-xs">
                                <div className="bg-blue-50/50 p-2.5 rounded-xl border border-blue-100"><p className="text-blue-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">Total Days</p><p className="text-base font-black text-blue-900">{row.regDays}</p></div>
                                <div className="bg-purple-50/50 p-2.5 rounded-xl border border-purple-100"><p className="text-purple-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">OT Hours</p><p className="text-base font-black text-purple-900">{row.otHours}</p></div>
                              </div>
                            ) : (
                              <div className="grid grid-cols-3 gap-2 text-xs">
                                <div className="bg-blue-50/50 p-2.5 rounded-xl border border-blue-100"><p className="text-blue-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">Mason</p><p className="text-base font-black text-blue-900">{row.masonReg} <span className="text-[10px] font-medium text-blue-400">({row.masonOT}h)</span></p></div>
                                <div className="bg-purple-50/50 p-2.5 rounded-xl border border-purple-100"><p className="text-purple-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">H. Mason</p><p className="text-base font-black text-purple-900">{row.halfMasonReg} <span className="text-[10px] font-medium text-purple-400">({row.halfMasonOT}h)</span></p></div>
                                <div className="bg-orange-50/50 p-2.5 rounded-xl border border-orange-100"><p className="text-orange-600/80 font-bold mb-0.5 uppercase tracking-wider text-[9px]">Helper</p><p className="text-base font-black text-orange-900">{row.helperReg} <span className="text-[10px] font-medium text-orange-400">({row.helperOT}h)</span></p></div>
                              </div>
                            )}
                            
                            <div className="bg-emerald-50 p-3 rounded-xl border border-emerald-100 flex justify-between items-center">
                              <p className="text-emerald-700 font-bold text-[10px] uppercase tracking-wider">{activeTab === 'worker' ? 'Total Payout' : 'Total Site Cost'}</p>
                              <p className="text-emerald-900 font-black text-lg">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</p>
                            </div>
                          </div>
                        ))
                      ) : (
                        <div className="text-center py-12 text-gray-500 font-medium bg-white rounded-2xl border border-gray-100">No records found</div>
                      )}

                      {filteredData.length > 0 && (
                        <div className="bg-gray-900 p-5 rounded-[2rem] shadow-xl mt-4 border border-gray-800">
                          <h3 className="text-white font-black text-center text-sm mb-4 tracking-wider">SHEET TOTALS</h3>
                          
                          {activeTab === 'worker' ? (
                            <div className="grid grid-cols-2 gap-2 text-xs mb-3">
                              <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-gray-400 font-bold mb-0.5 uppercase text-[9px]">Total Days</p><p className="text-base font-black text-white">{totals.masonReg + totals.halfMasonReg + totals.helperReg}</p></div>
                              <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-purple-400 font-bold mb-0.5 uppercase text-[9px]">Total OT Hours</p><p className="text-base font-black text-purple-200">{totals.masonOT + totals.halfMasonOT + totals.helperOT} h</p></div>
                            </div>
                          ) : (
                            <div className="grid grid-cols-3 gap-2 text-xs mb-3">
                              <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-gray-400 font-bold mb-0.5 uppercase text-[9px]">Total Mason</p><p className="text-base font-black text-white">{totals.masonReg} <span className="text-[10px] text-gray-400">({totals.masonOT}h)</span></p></div>
                              <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-purple-400 font-bold mb-0.5 uppercase text-[9px]">Total HM</p><p className="text-base font-black text-purple-200">{totals.halfMasonReg} <span className="text-[10px] text-gray-400">({totals.halfMasonOT}h)</span></p></div>
                              <div className="bg-gray-800/80 p-3 rounded-xl"><p className="text-gray-400 font-bold mb-0.5 uppercase text-[9px]">Total Helper</p><p className="text-base font-black text-white">{totals.helperReg} <span className="text-[10px] text-gray-400">({totals.helperOT}h)</span></p></div>
                            </div>
                          )}
                          
                          <div className="bg-emerald-600 p-4 rounded-2xl border border-emerald-500 text-center">
                            <p className="text-emerald-100 font-bold mb-1 uppercase text-[10px] tracking-wider">Grand Financial Total</p>
                            <p className="text-2xl font-black text-white">{formatCurrency(totals.totalBaseCost + totals.totalOTCost)}</p>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}
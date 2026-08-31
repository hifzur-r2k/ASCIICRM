import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { auth } from '../firebase';
import { signInWithEmailAndPassword } from 'firebase/auth';
import {
  Upload, Download, FileText, UploadCloud, LayoutGrid, Search,
  Save, Trash2, Database, Clock, Check, BarChart3,
  LogOut, Plus, X, Layers, IndianRupee, Calendar, Shield, Users, RefreshCw, ClipboardList, AlertCircle, CheckCircle, Edit2, Eye, EyeOff, MapPin
} from 'lucide-react';
import { collection, getDocs, setDoc, doc, deleteDoc, updateDoc, onSnapshot, query, where, orderBy, limit } from "firebase/firestore";
import { createUserWithEmailAndPassword } from "firebase/auth";
import { db, secondaryAuth } from '../firebase';

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

export default function AdminDashboard({ currentUser, onLogout }) {
  const [masterSites, setMasterSites] = useState([]);
  const [masterWorkers, setMasterWorkers] = useState([]);

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
  const [dailyLogs, setDailyLogs] = useState([]);
  const [pendingRequests, setPendingRequests] = useState([]);
  const [expandedLogId, setExpandedLogId] = useState(null);
  // --- NEW WORKER APPROVAL STATES ---
  const [approvingWorkerReq, setApprovingWorkerReq] = useState(null);
  const [newWorkerWage, setNewWorkerWage] = useState("");

  // --- TEST MODE (SHADOW DATABASE) STATES ---
  const [isTestMode, setIsTestMode] = useState(false);
  const COLL_SHEETS = isTestMode ? "test_attendance_sheets" : "attendance_sheets";
  const COLL_LOGS = isTestMode ? "test_daily_logs" : "daily_logs";
  const COLL_REQS = isTestMode ? "test_edit_requests" : "edit_requests";

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

  // --- TEAM MANAGEMENT STATES ---
  const [newSupName, setNewSupName] = useState("");
  const [newSupEmail, setNewSupEmail] = useState("");
  const [newSupPassword, setNewSupPassword] = useState("");
  const [teamMessage, setTeamMessage] = useState({ text: "", type: "" });
  const [isCreatingSup, setIsCreatingSup] = useState(false);
  const [newSiteCode, setNewSiteCode] = useState("");
  const [newSiteName, setNewSiteName] = useState("");
  const [loadedSheetId, setLoadedSheetId] = useState(null);
  const [showPassword, setShowPassword] = useState(false);

  const defaultContractors = ["Arvind", "Laljeet", "Deepak"];
  const dynamicContractors = Array.from(new Set([...defaultContractors, ...masterWorkers.map(w => w.contractor)])).filter(Boolean);

  const goHome = () => {
    setSiteData([]); setDayData([]); setWorkerData([]);
    setSelectedRecordContractor(null); setDashboardTab('upload');
  };

  const fetchData = async () => {
    setIsLoadingRecords(true);
    try {
      // 1. Fetch from either LIVE or TEST sheets bucket based on switch
      const sheetsSnap = await getDocs(collection(db, COLL_SHEETS));
      const sheets = [];
      let aggregatedSites = new Set();

      sheetsSnap.forEach(docSnap => {
        const data = docSnap.data();
        sheets.push({ id: docSnap.id, ...data });
        if (data.siteData && Array.isArray(data.siteData)) {
          data.siteData.forEach(s => { if (s.site) aggregatedSites.add(s.site); });
        }
      });
      sheets.sort((a, b) => b.id.localeCompare(a.id));
      setSavedSheets(sheets);

      // 2. Fetch master data (Master data is shared between Live/Test so sites/supervisors stay accurate)
      let loadedWorkers = [];
      let finalSitesArray = [];
      try {
        const masterSnap = await getDocs(collection(db, "master_data"));
        const siteMap = new Map();

        aggregatedSites.forEach(code => siteMap.set(code, code));

        masterSnap.forEach(docSnap => {
          if (docSnap.id === "sites" && docSnap.data().list) {
            docSnap.data().list.forEach(siteString => {
              const code = siteString.includes('|') ? siteString.split('|')[0] : siteString;
              siteMap.set(code, siteString);
            });
          }
          if (docSnap.id === "workers" && docSnap.data().list) loadedWorkers = docSnap.data().list;
        });

        finalSitesArray = Array.from(siteMap.values()).sort();
        if (finalSitesArray.length > 0) {
          await setDoc(doc(db, "master_data", "sites"), { list: finalSitesArray });
        }
      } catch (err) { console.warn("master_data missing:", err); }

      setMasterSites(finalSitesArray);
      setMasterWorkers(loadedWorkers);
    } catch (error) { console.error("Error fetching data:", error); }
    setIsLoadingRecords(false);
  };

  // Re-fetch immediately if the Test Mode switch is flipped
  useEffect(() => { fetchData(); }, [isTestMode]);

  // --- QUOTA-PROOF LIVE LISTENERS ---
  useEffect(() => {
    const qLogs = query(collection(db, COLL_LOGS), orderBy("timestamp", "desc"), limit(100));
    const unsubLogs = onSnapshot(qLogs, (snap) => {
      const logs = [];
      snap.forEach(d => logs.push({ id: d.id, ...d.data() }));
      setDailyLogs(logs);
    });

    const qReqs = query(collection(db, COLL_REQS), where("status", "==", "pending"));
    const unsubReqs = onSnapshot(qReqs, (snap) => {
      const reqs = [];
      snap.forEach(d => reqs.push({ id: d.id, ...d.data() }));
      setPendingRequests(reqs.sort((a, b) => b.timestamp - a.timestamp));
    });

    return () => { unsubLogs(); unsubReqs(); };
  }, [isTestMode]);

  const handleRequestAction = async (req, action) => {
    try {
      // INTERCEPT 1: Open the wage box for new workers
      if (req.type === 'new_worker' && action === 'approved') {
        setApprovingWorkerReq(req);
        return;
      }

      // INTERCEPT 2: Surgically remove fake worker from the log if denied
      if (req.type === 'new_worker' && action === 'denied') {
        const targetLog = dailyLogs.find(l => l.id === req.logId);
        if (targetLog) {
          const cleanedWorkers = targetLog.workers.filter(w => !(w.worker === req.workerName && w.isProvisional));
          await updateDoc(doc(db, COLL_LOGS, req.logId), { workers: cleanedWorkers });
        }
      }

      await updateDoc(doc(db, COLL_REQS, req.id), { status: action, actionAt: Date.now() });
    } catch (error) { alert("Error updating request."); }
  };

  const handleApproveNewWorker = async (req) => {
    if (!newWorkerWage || isNaN(newWorkerWage)) return alert("Please enter a valid daily wage.");
    const wageNum = parseFloat(newWorkerWage);
    
    try {
      // 1. Add to Master Roster permanently
      const newWorker = {
        name: req.workerName,
        type: req.workerType,
        contractor: req.contractor,
        wage: wageNum,
        otRate: wageNum / 8
      };
      const updatedMaster = [...masterWorkers, newWorker];
      await setDoc(doc(db, "master_data", "workers"), { list: updatedMaster });
      setMasterWorkers(updatedMaster);

      let addedBase = 0;
      let addedOT = 0;
      let targetRegDays = 0;
      let targetOTHours = 0;

      // 2. Inject real wage into the Daily Log
      const targetLog = dailyLogs.find(l => l.id === req.logId);
      if (targetLog) {
        const updatedWorkers = targetLog.workers.map(w => {
          if (w.worker === req.workerName && w.isProvisional) {
            targetRegDays = w.regDays;
            targetOTHours = w.otHours;
            addedBase = w.regDays * wageNum;
            addedOT = w.otHours * (wageNum / 8);
            return { ...w, totalBaseCost: addedBase, totalOTCost: addedOT, isProvisional: false };
          }
          return w;
        });
        await updateDoc(doc(db, COLL_LOGS, req.logId), { workers: updatedWorkers });
      }

      // 3. Retroactively fix the Financial Master Sheet (This was missing!)
      if (targetRegDays > 0 || targetOTHours > 0) {
        const d = new Date(req.date);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = d.getDate();
        const half = day <= 15 ? 'H1' : 'H2';
        const periodId = `${year}-${month}-${half}`;

        const sheetToUpdate = savedSheets.find(s => s.id === periodId);
        if (sheetToUpdate) {
          let updatedSiteData = [...(sheetToUpdate.siteData || [])];
          let updatedDayData = [...(sheetToUpdate.dayData || [])];
          let updatedWorkerData = [...(sheetToUpdate.workerData || [])];

          // Inject missing money into Worker Data
          const wIdx = updatedWorkerData.findIndex(w => w.worker === req.workerName && w.contractor === req.contractor);
          if (wIdx > -1) {
            updatedWorkerData[wIdx].totalBaseCost += addedBase;
            updatedWorkerData[wIdx].totalOTCost += addedOT;
          }

          // Inject missing money into Site Data
          const sIdx = updatedSiteData.findIndex(s => s.site === req.site && s.contractor === req.contractor);
          if (sIdx > -1) {
            updatedSiteData[sIdx].totalBaseCost += addedBase;
            updatedSiteData[sIdx].totalOTCost += addedOT;
          }

          // Inject missing money into Day Data
          const dIdx = updatedDayData.findIndex(dayData => dayData.day === day && dayData.site === req.site && dayData.contractor === req.contractor);
          if (dIdx > -1) {
            updatedDayData[dIdx].totalBaseCost += addedBase;
            updatedDayData[dIdx].totalOTCost += addedOT;
          }

          await updateDoc(doc(db, COLL_SHEETS, periodId), {
            siteData: updatedSiteData,
            dayData: updatedDayData,
            workerData: updatedWorkerData,
            updatedAt: Date.now()
          });
          fetchData(); // Force UI refresh
        }
      }

      // 4. Clear the request
      await updateDoc(doc(db, COLL_REQS, req.id), { status: 'approved', actionAt: Date.now() });
      setApprovingWorkerReq(null); setNewWorkerWage("");
      alert(`${req.workerName} successfully registered at ₹${wageNum}/day!`);
    } catch (err) { alert("Error approving worker."); console.error(err); }
  };

  const formatCurrency = (amount) => {
    return new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }).format(amount);
  };

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

  const saveToDatabase = async () => {
    if (hasSavedCurrent || !sheetMonth || !sheetContractor) return;
    setIsSaving(true);

    const isH2 = dayData.some(d => d.day > 15);
    const mockDate = `${sheetMonth}-${isH2 ? '16' : '01'}`;
    const period = getPeriodKey(mockDate);

    try {
      const docRef = doc(db, COLL_SHEETS, period.id);
      let existingData = { sheetName: period.displayName, siteData: [], dayData: [], workerData: [], createdAt: Date.now() };
      const docSnap = await getDocs(collection(db, COLL_SHEETS));
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
      await deleteDoc(doc(db, COLL_SHEETS, id));
      setSavedSheets(savedSheets.filter(sheet => sheet.id !== id));
      if (hasSavedCurrent) goHome();
    } catch (error) { console.error("Error deleting:", error); }
  };

  const deleteContractorRecord = async (contractorName, e) => {
    e.stopPropagation();
    const enteredPassword = window.prompt(`SECURITY LOCK\nEnter your Master Admin password to PERMANENTLY delete ${contractorName}'s records:`);
    if (!enteredPassword) return;

    try {
      await signInWithEmailAndPassword(auth, currentUser.email, enteredPassword);
    } catch (error) {
      return alert("🚨 INCORRECT PASSWORD! Deletion blocked.");
    }

    if (!window.confirm(`FINAL WARNING: Are you absolutely sure you want to delete all records for ${contractorName}?`)) return;

    try {
      const newSiteData = siteData.filter(s => s.contractor !== contractorName);
      const newDayData = dayData.filter(d => d.contractor !== contractorName);
      const newWorkerData = workerData.filter(w => w.contractor !== contractorName);

      await updateDoc(doc(db, COLL_SHEETS, loadedSheetId), {
        siteData: newSiteData, dayData: newDayData, workerData: newWorkerData, updatedAt: Date.now()
      });

      setSiteData(newSiteData); setDayData(newDayData); setWorkerData(newWorkerData);
      alert(`Successfully deleted ${contractorName}'s records.`);
      if (newWorkerData.length === 0) goHome();
    } catch (error) { alert("Error deleting contractor."); }
  };

  const handleDeleteSite = async (siteToDelete) => {
    if (!window.confirm(`Are you sure you want to delete this site code?`)) return;
    const updatedSites = masterSites.filter(s => s !== siteToDelete);
    try {
      await setDoc(doc(db, "master_data", "sites"), { list: updatedSites });
      setMasterSites(updatedSites);
    } catch (err) { alert("Error deleting site."); }
  };

  const loadSavedRecord = (sheet) => {
    setLoadedSheetId(sheet.id);
    setSheetName(sheet.sheetName); setSheetMonth(sheet.id.substring(0, 7));
    setSiteData(sheet.siteData || []); setDayData(sheet.dayData || []); setWorkerData(sheet.workerData || []);
    setHasSavedCurrent(true); setSearchQuery(""); setSelectedRecordContractor(null);
  };

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

  const masonsData = filteredData.filter(r => r.type === 'Mason');
  const halfMasonsData = filteredData.filter(r => r.type === 'HalfMason');
  const helpersData = filteredData.filter(r => r.type === 'Helper');

  const calcSubtotals = (data) => data.reduce((acc, row) => ({
    days: acc.days + (row.regDays || 0),
    ot: acc.ot + (row.otHours || 0),
    base: acc.base + (row.totalBaseCost || 0),
    otPay: acc.otPay + (row.totalOTCost || 0)
  }), { days: 0, ot: 0, base: 0, otPay: 0 });

  const exportToExcel = () => {
    if (filteredData.length === 0) return alert("No data to export.");
    let aoa = [
      ["CRM_FIX - FINANCIAL & ATTENDANCE REPORT"],
      [`Period: ${sheetName}`],
      [`Contractor: ${selectedRecordContractor}`],
      [`View: ${activeTab.toUpperCase()}-WISE`],
      []
    ];

    if (activeTab === 'worker') {
      aoa.push(["Worker Name", "Category", "Total Days", "Total OT (Hrs)", "Base Pay (Rs)", "OT Pay (Rs)", "Total Payout (Rs)"]);
      const addCategoryToAOA = (data, catName) => {
        if (data.length === 0) return;
        data.forEach(row => {
          aoa.push([row.worker, row.type, row.regDays, row.otHours, row.totalBaseCost, row.totalOTCost, row.totalBaseCost + row.totalOTCost]);
        });
        const sub = calcSubtotals(data);
        aoa.push([`${catName.toUpperCase()} SUBTOTAL`, "", sub.days, sub.ot, sub.base, sub.otPay, sub.base + sub.otPay]);
        aoa.push([]);
      };
      addCategoryToAOA(masonsData, "Mason");
      addCategoryToAOA(halfMasonsData, "Half Mason");
      addCategoryToAOA(helpersData, "Helper");
      aoa.push(["GRAND TOTAL", "", totals.masonReg + totals.halfMasonReg + totals.helperReg, totals.masonOT + totals.halfMasonOT + totals.helperOT, totals.totalBaseCost, totals.totalOTCost, totals.totalBaseCost + totals.totalOTCost]);
    } else {
      const headers = activeTab === 'day'
        ? ["Day", "Site Code", "Mason Days", "Mason OT", "HM Days", "HM OT", "Helper Days", "Helper OT", "Base Cost (Rs)", "OT Cost (Rs)", "Total Site Cost (Rs)"]
        : ["Site Code", "Mason Days", "Mason OT", "HM Days", "HM OT", "Helper Days", "Helper OT", "Base Cost (Rs)", "OT Cost (Rs)", "Total Site Cost (Rs)"];
      aoa.push(headers);
      filteredData.forEach(row => {
        const r = activeTab === 'day' ? [row.day] : [];
        r.push(row.site, row.masonReg, row.masonOT, row.halfMasonReg, row.halfMasonOT, row.helperReg, row.helperOT, row.totalBaseCost, row.totalOTCost, row.totalBaseCost + row.totalOTCost);
        aoa.push(r);
      });
      aoa.push([]);
      const ft = activeTab === 'day' ? ["GRAND TOTAL", ""] : ["GRAND TOTAL"];
      ft.push(totals.masonReg, totals.masonOT, totals.halfMasonReg, totals.halfMasonOT, totals.helperReg, totals.helperOT, totals.totalBaseCost, totals.totalOTCost, totals.totalBaseCost + totals.totalOTCost);
      aoa.push(ft);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } }, { s: { r: 2, c: 0 }, e: { r: 2, c: 6 } }, { s: { r: 3, c: 0 }, e: { r: 3, c: 6 } }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `CRM_FIX_${selectedRecordContractor}_${activeTab}_${sheetName}.xlsx`);
  };

  const exportToPDF = () => {
    if (filteredData.length === 0) return alert("No data to export.");
    try {
      const doc = new jsPDF('l');
      doc.setFontSize(18); doc.setTextColor(15, 23, 42); doc.text("CRM_FIX Financial Report", 14, 22);
      doc.setFontSize(10); doc.setTextColor(100, 116, 139); doc.text(`Period: ${sheetName}   |   Contractor: ${selectedRecordContractor}   |   View: ${activeTab.toUpperCase()}-WISE`, 14, 30);
      let head = []; let body = []; let foot = [];
      const formatNum = (num) => Number(num).toLocaleString('en-IN', { maximumFractionDigits: 0 });

      if (activeTab === 'worker') {
        head = [["Worker Name", "Category", "Total Days", "OT (Hrs)", "Base Pay (Rs)", "OT Pay (Rs)", "Total Payout (Rs)"]];
        const addCategoryToPDF = (data, catName) => {
          if (data.length === 0) return;
          data.forEach(row => { body.push([row.worker, row.type, row.regDays, row.otHours, formatNum(row.totalBaseCost), formatNum(row.totalOTCost), formatNum(row.totalBaseCost + row.totalOTCost)]); });
          const sub = calcSubtotals(data);
          body.push([
            { content: `${catName.toUpperCase()} SUBTOTAL`, colSpan: 2, styles: { fillColor: [241, 245, 249], fontStyle: 'bold', textColor: [15, 23, 42] } },
            { content: sub.days.toString(), styles: { fillColor: [241, 245, 249], fontStyle: 'bold', textColor: [15, 23, 42] } },
            { content: sub.ot.toString(), styles: { fillColor: [241, 245, 249], fontStyle: 'bold', textColor: [15, 23, 42] } },
            { content: formatNum(sub.base), styles: { fillColor: [241, 245, 249], fontStyle: 'bold', textColor: [15, 23, 42] } },
            { content: formatNum(sub.otPay), styles: { fillColor: [241, 245, 249], fontStyle: 'bold', textColor: [15, 23, 42] } },
            { content: formatNum(sub.base + sub.otPay), styles: { fillColor: [241, 245, 249], fontStyle: 'bold', textColor: [15, 23, 42] } }
          ]);
        };
        addCategoryToPDF(masonsData, "Mason"); addCategoryToPDF(halfMasonsData, "Half Mason"); addCategoryToPDF(helpersData, "Helper");
        foot = [["GRAND TOTAL", "", (totals.masonReg + totals.halfMasonReg + totals.helperReg).toString(), (totals.masonOT + totals.halfMasonOT + totals.helperOT).toString(), formatNum(totals.totalBaseCost), formatNum(totals.totalOTCost), formatNum(totals.totalBaseCost + totals.totalOTCost)]];
      } else {
        head = activeTab === 'day' ? [["Day", "Site", "Mason", "M. OT", "HM", "HM OT", "Helper", "H. OT", "Base Cost", "OT Cost", "Total Cost"]] : [["Site", "Mason", "M. OT", "HM", "HM OT", "Helper", "H. OT", "Base Cost", "OT Cost", "Total Cost"]];
        body = filteredData.map(row => {
          const r = activeTab === 'day' ? [row.day] : [];
          r.push(row.site, row.masonReg, row.masonOT, row.halfMasonReg, row.halfMasonOT, row.helperReg, row.helperOT, formatNum(row.totalBaseCost || 0), formatNum(row.totalOTCost || 0), formatNum((row.totalBaseCost || 0) + (row.totalOTCost || 0)));
          return r;
        });
        const ft = activeTab === 'day' ? ["GRAND TOTAL", ""] : ["GRAND TOTAL"];
        ft.push(totals.masonReg, totals.masonOT, totals.halfMasonReg, totals.halfMasonOT, totals.helperReg, totals.helperOT, formatNum(totals.totalBaseCost), formatNum(totals.totalOTCost), formatNum(totals.totalBaseCost + totals.totalOTCost));
        foot = [ft];
      }
      autoTable(doc, { startY: 35, head: head, body: body, foot: foot, showFoot: 'lastPage', theme: 'striped', headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontStyle: 'bold' }, footStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontStyle: 'bold' }, styles: { fontSize: 8, cellPadding: 3 }, });
      doc.save(`CRM_FIX_${selectedRecordContractor}_${activeTab}_${sheetName}.pdf`);
    } catch (error) { alert("Error generating PDF."); }
  };

  const handleCreateSupervisor = async (e) => {
    e.preventDefault();
    if (!newSupEmail || !newSupPassword) return;
    setIsCreatingSup(true); setTeamMessage({ text: "Authorizing secure link...", type: "loading" });
    try {
      await createUserWithEmailAndPassword(secondaryAuth, newSupEmail, newSupPassword);
      await setDoc(doc(db, "supervisors", newSupEmail.toLowerCase()), { name: newSupName || "Unnamed Supervisor", email: newSupEmail.toLowerCase(), createdAt: Date.now(), createdBy: currentUser.email });
      setTeamMessage({ text: `Success! ${newSupName || newSupEmail} is authorized.`, type: "success" });
      setNewSupName(""); setNewSupEmail(""); setNewSupPassword("");
    } catch (error) { setTeamMessage({ text: error.message, type: "error" }); }
    setIsCreatingSup(false);
  };

  const handleAddNewSite = async (e) => {
    e.preventDefault();
    const code = newSiteCode.trim().toUpperCase(); const name = newSiteName.trim();
    if (!code) return;
    const finalSiteString = name ? `${code}|${name}` : code;
    const filteredSites = masterSites.filter(s => (s.includes('|') ? s.split('|')[0] : s) !== code);
    const updatedSites = [...filteredSites, finalSiteString].sort();
    try {
      await setDoc(doc(db, "master_data", "sites"), { list: updatedSites });
      setMasterSites(updatedSites); setNewSiteCode(""); setNewSiteName(""); alert(`Success! Site updated.`);
    } catch (err) { alert("Error pushing new site to cloud."); }
  };

  const renderDesktopWorkerRow = (row, idx) => (
    <tr key={idx} className="border-b border-gray-50 hover:bg-gray-50/80 bg-white transition-colors">
      <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-900">{row.worker}</td>
      <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-bold text-gray-700">{row.type}</td>
      <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-gray-900">{row.regDays}</td>
      <td className="px-4 lg:px-6 py-4 border-r border-gray-50 font-medium text-purple-600">{row.otHours}</td>
      <td className="px-4 lg:px-6 py-4 border-r border-gray-50 text-emerald-700 font-medium">{formatCurrency(row.totalBaseCost || 0)}</td>
      <td className="px-4 lg:px-6 py-4 border-r border-gray-100 text-emerald-700 font-medium">{formatCurrency(row.totalOTCost || 0)}</td>
      <td className="px-4 lg:px-6 py-4 text-gray-900 font-black bg-gray-50/50">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</td>
    </tr>
  );

  const renderWorkerSubtotal = (data, label, colorClass) => {
    if (data.length === 0) return null;
    const sub = calcSubtotals(data);
    return (
      <tr className={`${colorClass} border-b-2 border-gray-200`}>
        <td colSpan="2" className="px-4 lg:px-6 py-3 text-right font-black uppercase text-[10px] tracking-widest">{label} SUBTOTAL:</td>
        <td className="px-4 lg:px-6 py-3 font-black text-sm">{sub.days}</td>
        <td className="px-4 lg:px-6 py-3 font-black text-sm">{sub.ot} h</td>
        <td className="px-4 lg:px-6 py-3 font-black text-sm">{formatCurrency(sub.base)}</td>
        <td className="px-4 lg:px-6 py-3 font-black text-sm">{formatCurrency(sub.otPay)}</td>
        <td className="px-4 lg:px-6 py-3 font-black text-sm bg-black/5">{formatCurrency(sub.base + sub.otPay)}</td>
      </tr>
    );
  };

  const renderMobileWorkerCard = (row, idx) => (
    <div key={idx} className="flex justify-between items-center p-3 hover:bg-gray-50 transition-colors border-b border-gray-50">
      <div className="min-w-0 flex-1">
        <p className="text-xs font-black text-gray-900 truncate uppercase">{row.worker}</p>
        <p className="text-[10px] text-gray-400 font-bold mt-0.5">{row.regDays} Days <span className="text-blue-500 font-black">+{row.otHours}h OT</span></p>
      </div>
      <div className="text-right shrink-0 ml-3">
        <p className="text-[8px] uppercase font-bold text-emerald-600/70 mb-0.5">Total Payout</p>
        <p className="text-sm font-black text-emerald-700">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</p>
      </div>
    </div>
  );

  const renderMobileSubtotal = (data, label, colorClass) => {
    if (data.length === 0) return null;
    const sub = calcSubtotals(data);
    return (
      <div className={`${colorClass} p-3 flex justify-between items-center shadow-inner`}>
        <div className="text-[9px] font-black uppercase tracking-widest">{label} SUBTOTAL<br /><span className="opacity-70">{sub.days} Days | {sub.ot}h OT</span></div>
        <div className="text-sm font-black">{formatCurrency(sub.base + sub.otPay)}</div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] text-gray-800 font-sans selection:bg-blue-100 selection:text-blue-900 pb-12">
      <header className="bg-white border-b border-gray-200 px-4 py-3 flex justify-between items-center shadow-sm sticky top-0 z-50">
        <div className="flex items-center gap-2 cursor-pointer hover:opacity-80 transition-opacity" onClick={goHome} title="Go to Dashboard">
          <div className="p-1.5 rounded-lg bg-blue-600"><Shield className="w-4 h-4 text-white" /></div>
          <span className="font-black text-gray-900 text-sm tracking-tight">CRM_FIX <span className="text-gray-400 font-medium hidden sm:inline">| Master Admin</span></span>
        </div>
        <div className="flex items-center gap-3 md:gap-4">

          {/* NEW: TEST MODE TOGGLE */}
          <div className="flex items-center gap-1 md:gap-2 bg-gray-50 border border-gray-200 p-1.5 rounded-xl">
            <span className={`text-[9px] md:text-[10px] font-black px-1 md:px-2 ${!isTestMode ? 'text-emerald-600' : 'text-gray-400'}`}>LIVE</span>
            <button onClick={() => { setIsTestMode(!isTestMode); setDashboardTab('logs'); }} className={`relative w-10 h-5 rounded-full transition-colors shadow-inner ${isTestMode ? 'bg-yellow-400' : 'bg-gray-300'}`}>
              <span className={`absolute top-1 w-3 h-3 bg-white rounded-full transition-all shadow-sm ${isTestMode ? 'right-1' : 'left-1'}`} />
            </button>
            <span className={`text-[9px] md:text-[10px] font-black px-1 md:px-2 ${isTestMode ? 'text-yellow-600' : 'text-gray-400'}`}>TEST</span>
          </div>

          <button onClick={fetchData} disabled={isLoadingRecords} className="p-2 bg-gray-100 hover:bg-blue-50 text-gray-600 hover:text-blue-600 rounded-xl transition-all" title="Sync All Records">
            <RefreshCw className={`w-3.5 h-3.5 ${isLoadingRecords ? 'animate-spin text-blue-500' : ''}`} />
          </button>
          <button onClick={onLogout} className="text-xs font-bold bg-gray-100 hover:bg-red-50 hover:text-red-600 text-gray-600 p-2 md:px-3 md:py-1.5 rounded-xl flex items-center gap-2 transition-colors"><LogOut className="w-4 h-4 md:w-3.5 md:h-3.5" /> <span className="hidden md:block">Sign Out</span></button>
        </div>
      </header>

      {/* NEW: GLOBAL WARNING BANNER */}
      {isTestMode && (
        <div className="bg-yellow-400 text-yellow-900 py-2 px-4 text-center text-[11px] font-black uppercase tracking-widest shadow-md flex items-center justify-center gap-2 sticky top-[60px] z-40">
          <AlertCircle className="w-4 h-4" /> TEST MODE ACTIVE — ALL DATA IS SAVED TO A FAKE ISOLATED DATABASE
        </div>
      )}

      <>
        <input type="file" accept=".xlsx, .xls, .csv" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
        {activeData.length === 0 ? (
          <div className="max-w-6xl mx-auto pt-8 md:pt-16 px-4">
            <div className="text-center mb-8 md:mb-10 space-y-2 md:space-y-3">
              <h1 className="text-3xl md:text-5xl font-black tracking-tight text-gray-900">Welcome Back.</h1>
              <p className="text-xs md:text-base font-bold text-gray-400 uppercase tracking-widest">Financial & Attendance Engine</p>
            </div>

            <div className="flex justify-center mb-8 px-2">
              <div className="bg-gray-100 p-1.5 rounded-[1.25rem] sm:rounded-full flex flex-wrap w-full sm:w-auto shadow-inner overflow-hidden gap-1">
                <button onClick={() => setDashboardTab('upload')} className={`flex-1 sm:flex-none px-2 sm:px-6 py-2.5 sm:py-3 rounded-xl sm:rounded-full font-bold text-[10px] sm:text-sm transition-all flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2.5 ${dashboardTab === 'upload' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                  <UploadCloud className={`w-5 h-5 sm:w-4 sm:h-4 ${dashboardTab === 'upload' ? 'text-gray-900' : 'text-gray-400'}`} /> <span className="text-center leading-tight">Upload</span>
                </button>
                <button onClick={() => { setDashboardTab('records'); fetchData(); }} className={`flex-1 sm:flex-none px-2 sm:px-6 py-2.5 sm:py-3 rounded-xl sm:rounded-full font-bold text-[10px] sm:text-sm transition-all flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2.5 ${dashboardTab === 'records' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
                  <Database className={`w-5 h-5 sm:w-4 sm:h-4 ${dashboardTab === 'records' ? 'text-gray-900' : 'text-gray-400'}`} /> <span className="text-center leading-tight">Records</span>
                </button>
                <button onClick={() => { setDashboardTab('analytics'); fetchData(); }} className={`flex-1 sm:flex-none px-2 sm:px-6 py-2.5 sm:py-3 rounded-xl sm:rounded-full font-bold text-[10px] sm:text-sm transition-all flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2.5 ${dashboardTab === 'analytics' ? 'bg-white text-blue-600 shadow-sm border border-gray-50' : 'text-gray-500 hover:text-gray-700'}`}>
                  <BarChart3 className={`w-5 h-5 sm:w-4 sm:h-4 ${dashboardTab === 'analytics' ? 'text-blue-600' : 'text-gray-400'}`} /> <span className="text-center leading-tight">Analytics</span>
                </button>
                <button onClick={() => setDashboardTab('logs')} className={`flex-1 sm:flex-none px-2 sm:px-6 py-2.5 sm:py-3 rounded-xl sm:rounded-full font-bold text-[10px] sm:text-sm transition-all flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2.5 ${dashboardTab === 'logs' ? 'bg-white text-orange-600 shadow-sm border border-gray-50' : 'text-gray-500 hover:text-gray-700'}`}>
                  <ClipboardList className={`w-5 h-5 sm:w-4 sm:h-4 ${dashboardTab === 'logs' ? 'text-orange-600' : 'text-gray-400'}`} /> <span className="text-center leading-tight">Logs</span>
                </button>
                <button onClick={() => setDashboardTab('approvals')} className={`flex-1 sm:flex-none px-2 sm:px-6 py-2.5 sm:py-3 rounded-xl sm:rounded-full font-bold text-[10px] sm:text-sm transition-all flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2.5 relative ${dashboardTab === 'approvals' ? 'bg-white text-rose-600 shadow-sm border border-gray-50' : 'text-gray-500 hover:text-gray-700'}`}>
                  <AlertCircle className={`w-5 h-5 sm:w-4 sm:h-4 ${dashboardTab === 'approvals' ? 'text-rose-600' : 'text-gray-400'}`} /> <span className="text-center leading-tight">Approvals</span>
                  {pendingRequests.length > 0 && <span className="absolute top-1 right-1 sm:top-2 sm:right-2 w-2 h-2 rounded-full bg-red-500"></span>}
                </button>
                <button onClick={() => setDashboardTab('team')} className={`flex-1 sm:flex-none px-2 sm:px-6 py-2.5 sm:py-3 rounded-xl sm:rounded-full font-bold text-[10px] sm:text-sm transition-all flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2.5 ${dashboardTab === 'team' ? 'bg-white text-purple-600 shadow-sm border border-gray-50' : 'text-gray-500 hover:text-gray-700'}`}>
                  <Users className={`w-5 h-5 sm:w-4 sm:h-4 ${dashboardTab === 'team' ? 'text-purple-600' : 'text-gray-400'}`} /> <span className="text-center leading-tight">Team</span>
                </button>
              </div>
            </div>

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
                                  <p className="text-xs font-semibold text-gray-400 flex items-center gap-1 mt-0.5"><Clock className="w-3 h-3" /> {new Date(record.date).toLocaleDateString()}</p>
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

            {/* TAB 4: AUDIT LOGS */}
            {dashboardTab === 'logs' && (
              <div className="max-w-4xl mx-auto">
                <div className="bg-white rounded-[2rem] p-6 md:p-8 shadow-sm border border-gray-100">
                  <h2 className="text-xl md:text-2xl font-black text-gray-900 mb-6 flex items-center gap-2">
                    <Clock className="w-6 h-6 text-orange-500" /> Audit Logs & Submissions
                  </h2>

                  <div className="space-y-3 max-h-[600px] overflow-y-auto custom-scrollbar pr-2">
                    {dailyLogs.length === 0 ? (
                      <p className="text-center text-gray-500 py-8 font-medium">No attendance logs found.</p>
                    ) : (
                      dailyLogs.map(log => (
                        <div key={log.id} className={`p-4 rounded-2xl border flex flex-col gap-4 transition-colors ${log.isEdited ? 'bg-yellow-50/50 border-yellow-200 shadow-[inset_0_0_0_1px_rgba(250,204,21,0.2)]' : 'bg-gray-50 border-gray-100'}`}>
                          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                            <div>
                              <div className="flex items-center gap-2 mb-1">
                                <span className="bg-orange-100 text-orange-700 text-[10px] font-black px-2 py-0.5 rounded uppercase tracking-wider">
                                  {new Date(log.timestamp).toLocaleString()}
                                </span>
                                {log.isEdited && (
                                  <span className="bg-yellow-400 text-yellow-900 text-[9px] font-black px-1.5 py-0.5 rounded uppercase tracking-wider flex items-center gap-1">
                                    <Edit2 className="w-2.5 h-2.5" /> EDITED {log.editCount ? `(${log.editCount})` : ''}
                                  </span>
                                )}
                                <span className="font-bold text-gray-900 text-sm ml-1">{log.site}</span>
                              </div>
                              <p className="text-xs text-gray-500 font-bold flex items-center gap-1">
                                <Users className="w-3 h-3" /> {log.contractor}'s Team <span className="text-gray-300 mx-1">|</span> {log.workers?.length || 0} Workers Logged
                              </p>
                            </div>
                            <div className="flex gap-2 w-full md:w-auto">
                              {/* VIEW WORKERS TOGGLE */}
                              <button onClick={() => setExpandedLogId(expandedLogId === log.id ? null : log.id)} className={`flex-1 md:flex-none px-3 py-2 font-black text-xs rounded-xl transition-colors flex items-center justify-center gap-1.5 border ${expandedLogId === log.id ? 'bg-gray-800 text-white border-gray-800' : 'bg-white text-gray-600 border-gray-200 hover:bg-gray-100'}`}>
                                {expandedLogId === log.id ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                                {expandedLogId === log.id ? 'Hide Roster' : 'View Roster'}
                              </button>

                              <div className={`px-3 py-2 rounded-xl border text-right flex-1 md:flex-none ${log.isEdited ? 'bg-white border-yellow-200' : 'bg-white border-gray-200'}`}>
                                <p className="text-[9px] text-gray-400 font-black uppercase tracking-wider mb-0.5">Submitted By</p>
                                <p className="text-xs font-bold text-blue-600">{log.submittedBy || "Unknown"}</p>
                                {log.editTimestamp && <p className="text-[8px] text-yellow-600 font-bold mt-0.5">Edited: {new Date(log.editTimestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</p>}
                              </div>
                            </div>
                          </div>

                          {/* EXPANDABLE GROUPED WORKER LIST */}
                          {expandedLogId === log.id && log.workers && (
                            <div className="pt-3 border-t border-gray-200/60 animate-in slide-in-from-top-2 fade-in duration-300">
                              {['Mason', 'HalfMason', 'Helper'].map(type => {
                                const catWorkers = log.workers.filter(w => w.type === type);
                                if (catWorkers.length === 0) return null;

                                const title = type === 'HalfMason' ? 'Half Masons' : type + 's';
                                const badgeClass = type === 'Mason' ? 'bg-blue-100 text-blue-700' : type === 'HalfMason' ? 'bg-purple-100 text-purple-700' : 'bg-orange-100 text-orange-700';

                                return (
                                  <div key={type} className="mb-3 last:mb-0">
                                    <h5 className={`text-[10px] font-black uppercase tracking-widest px-2 py-1 rounded w-max mb-2 ${badgeClass}`}>{title}</h5>
                                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
                                      {catWorkers.map((w, i) => (
                                        <div key={i} className="flex justify-between items-center bg-white px-3 py-2 rounded-lg border border-gray-100 shadow-sm">
                                          <span className="text-xs font-black text-gray-800">{w.worker}</span>
                                          <div className="text-[10px] font-bold text-right tracking-tight">
                                            <span className={w.regDays === 0.5 ? "text-yellow-600" : w.regDays === 2.0 ? "text-indigo-600" : "text-emerald-600"}>
                                              {w.regDays === 0.5 ? 'HD' : w.regDays === 2.0 ? '2P' : 'P'}
                                            </span>
                                            {w.otHours > 0 && <span className="text-purple-600 ml-1.5">+{w.otHours}h OT</span>}
                                          </div>
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      ))
                    )}
                  </div>
                </div>
              </div>
            )}

            {/* TAB 5: EDIT APPROVALS INBOX */}
            {dashboardTab === 'approvals' && (
              <div className="max-w-4xl mx-auto">
                <div className="bg-white rounded-[2rem] p-6 md:p-8 shadow-sm border border-gray-100">
                  <div className="flex justify-between items-center mb-6">
                    <h2 className="text-xl md:text-2xl font-black text-gray-900 flex items-center gap-2">
                      <AlertCircle className="w-6 h-6 text-rose-500" /> Permissions Inbox
                    </h2>
                    <span className="bg-rose-100 text-rose-700 text-xs font-black px-3 py-1 rounded-full uppercase tracking-wider">
                      {pendingRequests.length} Pending
                    </span>
                  </div>

                  <div className="space-y-3">
                    {pendingRequests.length === 0 ? (
                      <div className="text-center bg-gray-50 rounded-2xl border border-gray-100 py-10">
                        <CheckCircle className="w-8 h-8 text-emerald-400 mx-auto mb-2" />
                        <p className="text-gray-500 font-bold">You are all caught up! No pending requests.</p>
                      </div>
                    ) : (
                      pendingRequests.map(req => (
                        <div key={req.id} className="p-4 md:p-5 rounded-2xl border bg-white border-rose-200 shadow-[0_4px_15px_-3px_rgba(244,63,94,0.1)] flex flex-col md:flex-row justify-between items-start md:items-center gap-4 transition-all">
                          <div className="flex-1">
                            <div className="flex items-center gap-2 mb-1.5">
                              <span className={`text-[10px] font-black px-2 py-0.5 rounded uppercase tracking-wider ${req.type === 'new_worker' ? 'bg-orange-100 text-orange-700' : 'bg-rose-100 text-rose-700'}`}>
                                {req.type === 'backdate' ? 'Past Date Unlock' : req.type === 'new_worker' ? 'New Worker Alert' : 'Edit Request'}
                              </span>
                              <span className="font-black text-gray-900 text-sm">{req.site}</span>
                              <span className="text-gray-400 font-bold text-xs uppercase tracking-wider">| {req.date}</span>
                            </div>
                            <h4 className="font-black text-gray-900 text-base">
                              {req.type === 'new_worker'
                                ? <>{req.submittedBy.split('@')[0]} added <span className="text-blue-600">{req.workerName}</span> ({req.workerType}) to {req.contractor}'s team.</>
                                : <>{req.submittedBy.split('@')[0]} wants to {req.type === 'backdate' ? 'submit missing attendance' : 'edit an existing log'}</>
                              }
                            </h4>
                          </div>

                          <div className="w-full md:w-auto flex flex-col items-end gap-2 shrink-0">
                            {approvingWorkerReq?.id === req.id ? (
                              <div className="flex items-center gap-2 bg-orange-50 p-2 rounded-xl border border-orange-200 w-full md:w-auto">
                                <input type="number" placeholder="Enter Daily Wage (₹)" value={newWorkerWage} onChange={e => setNewWorkerWage(e.target.value)} className="w-40 px-3 py-2 text-sm font-bold border border-orange-200 rounded-lg outline-none focus:ring-2 focus:ring-orange-500/30" autoFocus />
                                <button onClick={() => setApprovingWorkerReq(null)} className="px-3 py-2 text-xs font-bold text-gray-500 hover:bg-gray-200 rounded-lg">Cancel</button>
                                <button onClick={() => handleApproveNewWorker(req)} className="px-4 py-2 bg-emerald-600 text-white font-black text-xs rounded-lg shadow-sm hover:bg-emerald-700 flex items-center gap-1"><CheckCircle className="w-3.5 h-3.5" /> Confirm</button>
                              </div>
                            ) : (
                              <div className="flex items-center gap-2 w-full md:w-auto">
                                <button onClick={() => handleRequestAction(req, 'denied')} className="flex-1 md:flex-none px-4 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-black text-xs rounded-xl transition-colors">
                                  Deny
                                </button>
                                <button onClick={() => handleRequestAction(req, 'approved')} className="flex-1 md:flex-none px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs rounded-xl transition-colors flex items-center justify-center gap-2 shadow-sm">
                                  <CheckCircle className="w-4 h-4" /> Approve
                                </button>
                              </div>
                            )}
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              </div>
            )}

            {/* TAB 6: SECURE TEAM MANAGEMENT */}
            {dashboardTab === 'team' && (
              <div className="max-w-2xl mx-auto">
                <div className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-sm border border-gray-100 text-center">
                  <div className="w-16 h-16 bg-purple-50 text-purple-600 rounded-full flex items-center justify-center mx-auto mb-6">
                    <Shield className="w-8 h-8" />
                  </div>
                  <h2 className="text-2xl font-black text-gray-900 mb-2">Team Management</h2>
                  <p className="text-sm font-medium text-gray-500 mb-8">Authorize specific emails to access the Supervisor Field Portal.</p>

                  <div className="bg-gray-50 border border-gray-200 rounded-2xl p-6 text-left">
                    <h3 className="text-sm font-bold text-gray-800 mb-4">Add New Supervisor Route</h3>
                    <form onSubmit={handleCreateSupervisor} className="space-y-4">
                      <input
                        type="text"
                        value={newSupName}
                        onChange={(e) => setNewSupName(e.target.value)}
                        placeholder="Supervisor Name (e.g., Laljeet)"
                        required
                        className="w-full px-4 py-3 bg-white border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-purple-500/20 focus:border-purple-500 transition-all"
                      />
                      <input
                        type="email"
                        value={newSupEmail}
                        onChange={(e) => setNewSupEmail(e.target.value)}
                        placeholder="Supervisor Email"
                        required
                        className="w-full px-4 py-3 bg-white border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-purple-500/20 focus:border-purple-500 transition-all"
                      />
                      <div className="relative w-full">
                        <input
                          type={showPassword ? "text" : "password"}
                          value={newSupPassword}
                          onChange={(e) => setNewSupPassword(e.target.value)}
                          placeholder="Create Password (min 6 characters)"
                          required
                          minLength={6}
                          className="w-full px-4 py-3 bg-white border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-purple-500/20 focus:border-purple-500 transition-all pr-12"
                        />
                        <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute right-4 top-1/2 -translate-y-1/2 text-gray-400 hover:text-purple-600 outline-none">
                          {showPassword ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
                        </button>
                      </div>

                      {teamMessage.text && (
                        <p className={`text-xs font-bold text-center py-2 rounded-lg ${teamMessage.type === 'error' ? 'text-red-500 bg-red-50' :
                          teamMessage.type === 'loading' ? 'text-blue-500 bg-blue-50' :
                            'text-emerald-600 bg-emerald-50'
                          }`}>
                          {teamMessage.text}
                        </p>
                      )}

                      <button disabled={isCreatingSup} type="submit" className={`w-full py-3 text-white font-bold rounded-xl shadow-md transition-colors flex items-center justify-center gap-2 ${isCreatingSup ? 'bg-purple-400 cursor-not-allowed' : 'bg-purple-600 hover:bg-purple-700'}`}>
                        <Plus className="w-4 h-4" /> {isCreatingSup ? 'Authorizing...' : 'Authorize Account'}
                      </button>
                    </form>
                  </div>

                  <div className="bg-gray-50 border border-gray-200 rounded-2xl p-6 text-left mt-6">
                    <h3 className="text-sm font-bold text-gray-800 mb-4 flex items-center gap-2"><Layers className="w-4 h-4 text-emerald-500" /> Create New Site</h3>
                    <form onSubmit={handleAddNewSite} className="flex flex-col sm:flex-row gap-3">
                      <input
                        type="text"
                        value={newSiteCode}
                        onChange={(e) => setNewSiteCode(e.target.value)}
                        placeholder="Short Code (e.g. BN)"
                        required
                        className="w-full sm:flex-[0.5] px-4 py-3 bg-white border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all uppercase"
                      />
                      <input
                        type="text"
                        value={newSiteName}
                        onChange={(e) => setNewSiteName(e.target.value)}
                        placeholder="Full Site Name (e.g. By Nature Project)"
                        className="w-full sm:flex-1 px-4 py-3 bg-white border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all"
                      />
                      <button type="submit" className="w-full sm:w-auto px-6 py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl shadow-md transition-colors flex items-center justify-center gap-2">
                        <Plus className="w-4 h-4" /> Add Site
                      </button>
                    </form>
                  </div>

                  {/* NEW: ACTIVE SITES LIST & DELETE */}
                  {masterSites.length > 0 && (
                    <div className="bg-gray-50 border border-gray-200 rounded-2xl p-6 text-left mt-6">
                      <h3 className="text-sm font-bold text-gray-800 mb-4 flex items-center gap-2"><MapPin className="w-4 h-4 text-blue-500" /> Manage Active Sites</h3>
                      <div className="flex flex-wrap gap-2">
                        {masterSites.map((site, idx) => (
                          <div key={idx} className="bg-white border border-gray-200 text-gray-700 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-2 shadow-sm">
                            {site.includes('|') ? `${site.split('|')[0]} - ${site.split('|')[1]}` : site}
                            <button onClick={() => handleDeleteSite(site)} className="text-gray-400 hover:text-red-500 transition-colors bg-gray-50 hover:bg-red-50 rounded p-0.5" title="Delete Site">
                              <X className="w-3 h-3" />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                </div>
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

            {!selectedRecordContractor ? (
              <div className="bg-white rounded-[2.5rem] shadow-sm border border-gray-100 p-8 md:p-12 text-center max-w-4xl mx-auto mt-8">
                <h2 className="text-2xl md:text-3xl font-black text-gray-900 mb-2">Select Contractor Team</h2>
                <p className="text-sm font-bold text-gray-400 uppercase tracking-widest mb-10">Data Period: {sheetName}</p>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 md:gap-6">
                  {Array.from(new Set([...dynamicContractors, ...workerData.map(w => w.contractor)])).filter(Boolean).map(c => {
                    const hasData = workerData.some(w => w.contractor === c);
                    return (
                      <div key={c} onClick={() => hasData && setSelectedRecordContractor(c)} className={`group relative p-6 md:p-8 rounded-[2rem] border-2 transition-all ${hasData ? 'bg-white hover:bg-blue-50/50 border-gray-100 hover:border-blue-200 cursor-pointer shadow-sm hover:shadow-xl' : 'bg-gray-50 border-dashed border-gray-200 opacity-50 cursor-not-allowed'}`}>
                        <div className={`w-12 h-12 mx-auto rounded-full flex items-center justify-center mb-4 transition-transform ${hasData ? 'bg-blue-100 text-blue-600 group-hover:scale-110' : 'bg-gray-200 text-gray-400'}`}>
                          <Users className="w-6 h-6" />
                        </div>
                        <h3 className="text-lg font-black text-gray-900">{c}</h3>
                        <p className="text-[10px] font-bold uppercase tracking-widest mt-1 text-gray-400">{hasData ? 'View Analytics' : 'No Data'}</p>

                        {hasData && loadedSheetId && (
                          <button onClick={(e) => deleteContractorRecord(c, e)} className="absolute top-4 right-4 p-2 bg-red-50 text-red-500 rounded-lg hover:bg-red-500 hover:text-white transition-all shadow-sm" title="Delete Contractor Data">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                    )
                  })}
                </div>
              </div>
            ) : (
              <div className="bg-white rounded-[2rem] shadow-sm border border-gray-100 overflow-hidden mt-6">
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
                    <input type="text" placeholder="Search..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full pl-11 pr-4 py-2.5 md:py-3 bg-white border border-gray-200 rounded-2xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all shadow-sm" />
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
                          <>
                            {masonsData.length > 0 && (
                              <>
                                {masonsData.map(renderDesktopWorkerRow)}
                                {renderWorkerSubtotal(masonsData, "Mason", "bg-blue-50 text-blue-900")}
                              </>
                            )}
                            {halfMasonsData.length > 0 && (
                              <>
                                {halfMasonsData.map(renderDesktopWorkerRow)}
                                {renderWorkerSubtotal(halfMasonsData, "Half Mason", "bg-purple-50 text-purple-900")}
                              </>
                            )}
                            {helpersData.length > 0 && (
                              <>
                                {helpersData.map(renderDesktopWorkerRow)}
                                {renderWorkerSubtotal(helpersData, "Helper", "bg-orange-50 text-orange-900")}
                              </>
                            )}
                          </>
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

                  <div className="block md:hidden">
                    {filteredData.length > 0 ? (
                      activeTab === 'site' ? (
                        <div className="p-2 space-y-2.5 bg-gray-50/50">
                          {filteredData.map((row, idx) => (
                            <div key={idx} className="bg-white p-3 rounded-xl shadow-[0_1px_3px_rgba(0,0,0,0.05)] border border-gray-100 space-y-2">
                              <div className="flex justify-between items-center border-b border-gray-50 pb-1.5">
                                <span className="text-xs font-black text-gray-800 bg-gray-100 px-2 py-1 rounded-md truncate max-w-[200px]">Site: {row.site}</span>
                              </div>
                              <div className="grid grid-cols-3 gap-1.5 text-xs">
                                <div className="bg-blue-50/50 p-1.5 rounded-lg border border-blue-100"><p className="text-blue-600/80 font-bold mb-0.5 uppercase tracking-wider text-[8px]">Mason</p><p className="text-sm font-black text-blue-900">{row.masonReg} <span className="text-[9px] font-medium text-blue-400">({row.masonOT}h)</span></p></div>
                                <div className="bg-purple-50/50 p-1.5 rounded-lg border border-purple-100"><p className="text-purple-600/80 font-bold mb-0.5 uppercase tracking-wider text-[8px]">H. Mason</p><p className="text-sm font-black text-purple-900">{row.halfMasonReg} <span className="text-[9px] font-medium text-purple-400">({row.halfMasonOT}h)</span></p></div>
                                <div className="bg-orange-50/50 p-1.5 rounded-lg border border-orange-100"><p className="text-orange-600/80 font-bold mb-0.5 uppercase tracking-wider text-[8px]">Helper</p><p className="text-sm font-black text-orange-900">{row.helperReg} <span className="text-[9px] font-medium text-orange-400">({row.helperOT}h)</span></p></div>
                              </div>
                              <div className="bg-emerald-50 px-2.5 py-1.5 rounded-lg border border-emerald-100 flex justify-between items-center">
                                <p className="text-emerald-700 font-bold text-[9px] uppercase tracking-wider">Total Site Cost</p>
                                <p className="text-emerald-900 font-black text-base">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</p>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : activeTab === 'day' ? (
                        <div className="bg-white border-y border-gray-100 divide-y divide-gray-50 mt-2">
                          {filteredData.map((row, idx) => (
                            <div key={idx} className="flex justify-between items-center p-3 hover:bg-gray-50 transition-colors">
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-1.5 mb-0.5">
                                  <span className="text-[9px] font-black text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded uppercase">Day {row.day}</span>
                                  <p className="text-xs font-black text-gray-800 truncate">{row.site}</p>
                                </div>
                                <p className="text-[10px] text-gray-400 font-bold">
                                  M:{row.masonReg} <span className="text-gray-300 mx-0.5">|</span> HM:{row.halfMasonReg} <span className="text-gray-300 mx-0.5">|</span> H:{row.helperReg}
                                </p>
                              </div>
                              <div className="text-right shrink-0 ml-3">
                                <p className="text-[8px] uppercase font-bold text-emerald-600/70 mb-0.5">Daily Cost</p>
                                <p className="text-sm font-black text-emerald-700">{formatCurrency((row.totalBaseCost || 0) + (row.totalOTCost || 0))}</p>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="bg-gray-50/50 border-y border-gray-100 mt-2 flex flex-col">
                          {masonsData.length > 0 && (
                            <div className="mb-3 bg-white border-y border-gray-100 shadow-sm">
                              <div className="bg-blue-600 text-white text-[10px] font-black px-3 py-1.5 uppercase tracking-widest">Masons</div>
                              {masonsData.map(renderMobileWorkerCard)}
                              {renderMobileSubtotal(masonsData, "Mason", "bg-blue-50 text-blue-900")}
                            </div>
                          )}
                          {halfMasonsData.length > 0 && (
                            <div className="mb-3 bg-white border-y border-gray-100 shadow-sm">
                              <div className="bg-purple-600 text-white text-[10px] font-black px-3 py-1.5 uppercase tracking-widest">Half Masons</div>
                              {halfMasonsData.map(renderMobileWorkerCard)}
                              {renderMobileSubtotal(halfMasonsData, "Half Mason", "bg-purple-50 text-purple-900")}
                            </div>
                          )}
                          {helpersData.length > 0 && (
                            <div className="mb-3 bg-white border-y border-gray-100 shadow-sm">
                              <div className="bg-orange-600 text-white text-[10px] font-black px-3 py-1.5 uppercase tracking-widest">Helpers</div>
                              {helpersData.map(renderMobileWorkerCard)}
                              {renderMobileSubtotal(helpersData, "Helper", "bg-orange-50 text-orange-900")}
                            </div>
                          )}
                        </div>
                      )
                    ) : (
                      <div className="text-center py-10 text-gray-500 font-medium bg-white border-y border-gray-100 text-sm">No records found</div>
                    )}

                    {filteredData.length > 0 && (
                      <div className="bg-gray-900 p-5 shadow-[0_-4px_15px_-3px_rgba(0,0,0,0.1)]">
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

                        <div className="bg-emerald-600 p-4 rounded-xl border border-emerald-500 text-center">
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
    </div>
  );
}
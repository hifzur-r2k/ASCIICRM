import React, { useState, useEffect } from 'react';
import {
  LayoutGrid, Search, Save, LogOut, Calendar, Users, MapPin, HardHat, ChevronDown, Clock, Edit2, AlertCircle, RefreshCw, Eye, EyeOff, Plus, X
} from 'lucide-react';
import { collection, addDoc, getDocs, doc, query, where, runTransaction, updateDoc } from "firebase/firestore";
import { db } from '../firebase';

const getPeriodKey = (dateString) => {
  const d = new Date(dateString);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = d.getDate();
  const half = day <= 15 ? 'H1' : 'H2';

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const displayMonth = monthNames[d.getMonth()];
  const displayRange = half === 'H1' ? '1-15' : '16-31';

  return { id: `${year}-${month}-${half}`, displayName: `${displayMonth} ${displayRange} ${year}` };
};

const getSiteCode = (siteStr) => siteStr.includes('|') ? siteStr.split('|')[0] : siteStr;
const getSiteDisplay = (siteStr) => {
  if (siteStr.includes('|')) {
    const parts = siteStr.split('|');
    return `${parts[1]} (${parts[0]})`;
  }
  return siteStr;
};

export default function SupervisorPortal({ currentUser, onLogout }) {
  const [activePortalTab, setActivePortalTab] = useState('attendance');
  const [masterSites, setMasterSites] = useState([]);
  const [masterWorkers, setMasterWorkers] = useState([]);
  const [myLogs, setMyLogs] = useState([]);
  const [expandedLogId, setExpandedLogId] = useState(null);

  const [toastMessage, setToastMessage] = useState("");

  // --- TEST MODE DETECTOR ---
  const isTestMode = currentUser?.email?.toLowerCase()?.includes('test') || false;
  const COLL_LOGS = isTestMode ? "test_daily_logs" : "daily_logs";
  const COLL_REQS = isTestMode ? "test_edit_requests" : "edit_requests";
  const COLL_SHEETS = isTestMode ? "test_attendance_sheets" : "attendance_sheets";

  const [myRequests, setMyRequests] = useState([]);

  // Attendance Form States
  const [supDate, setSupDate] = useState(new Date().toISOString().split('T')[0]);
  const [supSite, setSupSite] = useState("");
  const [supContractor, setSupContractor] = useState("");
  const [workerSearch, setWorkerSearch] = useState("");
  const [supAttendance, setSupAttendance] = useState({});
  const [isSubmittingLog, setIsSubmittingLog] = useState(false);
  const [editingLog, setEditingLog] = useState(null);

  const [isSiteDropdownOpen, setIsSiteDropdownOpen] = useState(false);
  const [isTeamDropdownOpen, setIsTeamDropdownOpen] = useState(false);
  const [siteSearchQuery, setSiteSearchQuery] = useState("");

  // PROVISIONAL WORKER STATES
  const [showAddWorker, setShowAddWorker] = useState(false);
  const [newWorkerName, setNewWorkerName] = useState("");
  const [newWorkerType, setNewWorkerType] = useState("Helper");

  const defaultContractors = ["Arvind", "Laljeet", "Deepak"];
  const dynamicContractors = Array.from(new Set([...defaultContractors, ...masterWorkers.map(w => w.contractor)])).filter(Boolean);

  useEffect(() => {
    const fetchMasterData = async () => {
      try {
        const masterSnap = await getDocs(collection(db, "master_data"));
        let loadedSites = [];
        let loadedWorkers = [];
        masterSnap.forEach(docSnap => {
          if (docSnap.id === "sites" && docSnap.data().list) loadedSites = docSnap.data().list;
          if (docSnap.id === "workers" && docSnap.data().list) loadedWorkers = docSnap.data().list;
        });

        if (loadedSites.length === 0) {
          const sheetsSnap = await getDocs(collection(db, COLL_SHEETS));
          let aggregatedSites = new Set();
          sheetsSnap.forEach(docSnap => {
            const data = docSnap.data();
            if (data.siteData && Array.isArray(data.siteData)) {
              data.siteData.forEach(s => { if (s.site) aggregatedSites.add(s.site); });
            }
          });
          loadedSites = Array.from(aggregatedSites);
        }
        setMasterSites(loadedSites.sort());
        setMasterWorkers(loadedWorkers);
      } catch (err) { console.error("Error fetching master data:", err); }
    };
    fetchMasterData();
  }, []);

  const fetchHistoryData = async () => {
    if (!currentUser?.email) return;
    try {
      const qLogs = query(collection(db, COLL_LOGS), where("submittedBy", "==", currentUser.email));
      const snapLogs = await getDocs(qLogs);
      const logs = [];
      snapLogs.forEach(d => logs.push({ id: d.id, ...d.data() }));
      logs.sort((a, b) => b.timestamp - a.timestamp);
      setMyLogs(logs);

      const qReq = query(collection(db, COLL_REQS), where("submittedBy", "==", currentUser.email));
      const snapReq = await getDocs(qReq);
      const reqs = [];
      snapReq.forEach(d => reqs.push({ id: d.id, ...d.data() }));
      setMyRequests(reqs);
    } catch (err) { console.error("Error fetching history data:", err); }
  };

  useEffect(() => { fetchHistoryData(); }, [currentUser]);
  useEffect(() => { if (activePortalTab === 'history') fetchHistoryData(); }, [activePortalTab]);

  useEffect(() => {
    window.history.pushState({ noBackExitsApp: true }, '');
    const handlePopState = (e) => {
      window.history.pushState({ noBackExitsApp: true }, '');
      if (activePortalTab !== 'attendance') setActivePortalTab('attendance');
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, [activePortalTab]);

  useEffect(() => {
    if (supContractor && !editingLog) {
      const contractorWorkers = masterWorkers.filter(w => w.contractor === supContractor);
      const defaultState = contractorWorkers.reduce((acc, w) => {
        acc[w.name] = { status: 'absent', ot: '' };
        return acc;
      }, {});
      setSupAttendance(defaultState);
    }
  }, [supContractor, editingLog]);

  const handleAttendanceChange = (name, status) => {
    setSupAttendance(prev => ({ ...prev, [name]: { ...prev[name], status: status } }));
  };

  const handleOTChange = (name, otValue) => {
    if (otValue === '') { setSupAttendance(prev => ({ ...prev, [name]: { ...prev[name], ot: '' } })); return; }
    let val = parseFloat(otValue);
    if (val > 12) val = 12;
    if (val < 0) val = 0;
    setSupAttendance(prev => ({ ...prev, [name]: { ...prev[name], ot: isNaN(val) ? '' : val } }));
  };

  const handleAddNewWorkerLocal = () => {
    if (!newWorkerName.trim()) return;
    const nameCaps = newWorkerName.trim().toUpperCase();

    if (masterWorkers.some(w => w.name.toUpperCase() === nameCaps && w.contractor === supContractor)) {
      return alert("This worker already exists in the roster!");
    }

    const newW = {
      name: nameCaps,
      type: newWorkerType,
      contractor: supContractor,
      wage: 0,
      otRate: 0,
      isProvisional: true
    };

    setMasterWorkers(prev => [...prev, newW]);
    
    // 1. STRICTLY DEFAULT TO ABSENT (Not Present)
    setSupAttendance(prev => ({ ...prev, [nameCaps]: { status: 'absent', ot: '' } }));

    setShowAddWorker(false);
    setNewWorkerName("");

    // 2. FRONT & CENTER: Auto-search their name so they are the only one on screen
    setWorkerSearch(nameCaps);

    // 3. EXACT MESSAGE REQUESTED
    setToastMessage(`✅ New ${newWorkerType} "${nameCaps}" is added!`);
    setTimeout(() => setToastMessage(""), 3500);
  };
  
  const sortWorkers = (workersList) => {
    return [...workersList].sort((a, b) => {
      // 1. NEW UPGRADE: Always force newly added (provisional) workers to the absolute top of the list
      if (a.isProvisional && !b.isProvisional) return -1;
      if (!a.isProvisional && b.isProvisional) return 1;

      // 2. Float present workers
      const aRec = supAttendance[a.name];
      const bRec = supAttendance[b.name];
      const aAct = aRec && (aRec.status !== 'absent' || parseFloat(aRec.ot) > 0) ? 1 : 0;
      const bAct = bRec && (bRec.status !== 'absent' || parseFloat(bRec.ot) > 0) ? 1 : 0;
      if (aAct !== bAct) return bAct - aAct;
      
      // 3. Alphabetical order for everyone else
      return a.name.localeCompare(b.name);
    });
  };

  const activeContractorWorkers = masterWorkers.filter(w => w.contractor === supContractor);
  const searchedWorkers = activeContractorWorkers.filter(w => w.name.toLowerCase().includes(workerSearch.toLowerCase()));
  const masons = sortWorkers(searchedWorkers.filter(w => w.type === 'Mason'));
  const halfMasons = sortWorkers(searchedWorkers.filter(w => w.type === 'HalfMason'));
  const helpers = sortWorkers(searchedWorkers.filter(w => w.type === 'Helper'));
  const isToday = (dateStr) => dateStr === new Date().toISOString().split('T')[0];

  const currentHour = new Date().getHours();
  const hasSubmittedToday = myLogs.some(log => isToday(log.date));
  const isOverdue = !hasSubmittedToday && currentHour >= 12;
  const overdueMessage = currentHour >= 16
    ? "URGENT: 4:00 PM Deadline Passed! Submit attendance immediately."
    : "REMINDER: 12:00 PM Deadline Passed. Please submit morning attendance.";

  const submitDailyLog = async () => {
    if (!supSite || !supContractor) return alert("Please select a site and contractor.");
    setIsSubmittingLog(true);

    const todayStr = new Date().toISOString().split('T')[0];

    if (supDate > todayStr) {
      setIsSubmittingLog(false);
      return alert("🚨 NOT ALLOWED: You cannot submit attendance for tomorrow or any future date.");
    }

    if (supDate < todayStr && !editingLog) {
      const hasBackdateApproval = myRequests.find(r => r.date === supDate && r.status === 'approved' && r.type === 'backdate');
      if (!hasBackdateApproval) {
        setIsSubmittingLog(false);
        return alert("🚨You can only submit attendance for TODAY.\n\nPlease use the 'Request Past Date Unlock' link below the calendar to get Admin approval first.");
      }
    }

    if (!editingLog) {
      const duplicateCheckQuery = query(collection(db, COLL_LOGS),
        where("date", "==", supDate),
        where("site", "==", supSite),
        where("contractor", "==", supContractor)
      );
      const duplicateSnap = await getDocs(duplicateCheckQuery);

      if (!duplicateSnap.empty) {
        setIsSubmittingLog(false);
        return alert(`🛑 DUPLICATE ATTENDANCE: You have already submitted attendance for ${supContractor} at this site today!\n\nIf you need to make changes, go to "My Logs" and click 'Edit Log'.`);
      }
    }

    const period = getPeriodKey(supDate);
    const dayNum = parseInt(supDate.split('-')[2]);

    const logWorkers = [];
    let dMasonReg = 0, dMasonOT = 0, dHMMasonReg = 0, dHMMasonOT = 0, dHelperReg = 0, dHelperOT = 0;
    let dTotalBaseCost = 0, dTotalOTCost = 0;

    Object.keys(supAttendance).forEach(name => {
      const rec = supAttendance[name];
      const hasHours = parseFloat(rec.ot) > 0;

      // Let them through if they are present OR if they have hours logged
      if (rec.status !== 'absent' || hasHours) {
        const workerInfo = masterWorkers.find(w => w.name === name && w.contractor === supContractor);
        if (workerInfo) {
          // Strictly apply 0 base days if they are marked Absent
          const regDays = rec.status === '2P' ? 2.0 : rec.status === 'half' ? 0.5 : rec.status === 'present' ? 1.0 : 0;
          const otHours = parseFloat(rec.ot) || 0;
          const bCost = regDays * (workerInfo.wage || 0);
          const oCost = otHours * (workerInfo.otRate || ((workerInfo.wage || 0) / 8));

          logWorkers.push({
            worker: workerInfo.name,
            type: workerInfo.type,
            contractor: supContractor,
            regDays: regDays,
            otHours: otHours,
            totalBaseCost: bCost,
            totalOTCost: oCost,
            isProvisional: workerInfo.isProvisional || false // Flags new workers
          });

          dTotalBaseCost += bCost; dTotalOTCost += oCost;
          if (workerInfo.type === 'Mason') { dMasonReg += regDays; dMasonOT += otHours; }
          else if (workerInfo.type === 'HalfMason') { dHMMasonReg += regDays; dHMMasonOT += otHours; }
          else { dHelperReg += regDays; dHelperOT += otHours; }
        }
      }
    });

    if (logWorkers.length === 0 && !editingLog) {
      alert("No workers marked present or given hours. Nothing to save.");
      setIsSubmittingLog(false); return;
    }

    const todaySiteData = { site: supSite, contractor: supContractor, masonReg: dMasonReg, masonOT: dMasonOT, halfMasonReg: dHMMasonReg, halfMasonOT: dHMMasonOT, helperReg: dHelperReg, helperOT: dHelperOT, totalBaseCost: dTotalBaseCost, totalOTCost: dTotalOTCost };
    const todayDayData = { day: dayNum, site: supSite, contractor: supContractor, masonReg: dMasonReg, masonOT: dMasonOT, halfMasonReg: dHMMasonReg, halfMasonOT: dHMMasonOT, helperReg: dHelperReg, helperOT: dHelperOT, totalBaseCost: dTotalBaseCost, totalOTCost: dTotalOTCost };

    let createdLogId = editingLog ? editingLog.id : null;

    try {
      await runTransaction(db, async (transaction) => {
        const sheetRef = doc(db, COLL_SHEETS, period.id);
        const sheetSnap = await transaction.get(sheetRef);
        let data = sheetSnap.exists() ? sheetSnap.data() : { sheetName: period.displayName, siteData: [], dayData: [], workerData: [], createdAt: Date.now() };

        let currentSiteData = [...(data.siteData || [])];
        let currentDayData = [...(data.dayData || [])];
        let currentWorkerData = [...(data.workerData || [])];

        if (editingLog) {
          currentDayData = currentDayData.filter(d => !(d.day === dayNum && d.site === supSite && d.contractor === supContractor));
          const sIdx = currentSiteData.findIndex(s => s.site === supSite && s.contractor === supContractor);
          if (sIdx > -1) {
            let oldMasonReg = 0, oldMasonOT = 0, oldHMReg = 0, oldHMOT = 0, oldHelperReg = 0, oldHelperOT = 0, oldBaseCost = 0, oldOTCost = 0;
            editingLog.workers.forEach(w => {
              if (w.type === 'Mason') { oldMasonReg += w.regDays; oldMasonOT += w.otHours; }
              else if (w.type === 'HalfMason') { oldHMReg += w.regDays; oldHMOT += w.otHours; }
              else { oldHelperReg += w.regDays; oldHelperOT += w.otHours; }
              oldBaseCost += w.totalBaseCost; oldOTCost += w.totalOTCost;
            });
            currentSiteData[sIdx].masonReg -= oldMasonReg; currentSiteData[sIdx].masonOT -= oldMasonOT;
            currentSiteData[sIdx].halfMasonReg -= oldHMReg; currentSiteData[sIdx].halfMasonOT -= oldHMOT;
            currentSiteData[sIdx].helperReg -= oldHelperReg; currentSiteData[sIdx].helperOT -= oldHelperOT;
            currentSiteData[sIdx].totalBaseCost -= oldBaseCost; currentSiteData[sIdx].totalOTCost -= oldOTCost;
          }

          editingLog.workers.forEach(oldW => {
            const wIdx = currentWorkerData.findIndex(w => w.worker === oldW.worker && w.contractor === supContractor);
            if (wIdx > -1) {
              currentWorkerData[wIdx].regDays -= oldW.regDays; currentWorkerData[wIdx].otHours -= oldW.otHours;
              currentWorkerData[wIdx].totalBaseCost -= oldW.totalBaseCost; currentWorkerData[wIdx].totalOTCost -= oldW.totalOTCost;
            }
          });
        } else {
          currentDayData = currentDayData.filter(d => !(d.day === dayNum && d.site === supSite && d.contractor === supContractor));
        }

        if (logWorkers.length > 0) {
          currentDayData.push(todayDayData);
          const existingSiteIdx = currentSiteData.findIndex(s => s.site === supSite && s.contractor === supContractor);
          if (existingSiteIdx > -1) {
            currentSiteData[existingSiteIdx].masonReg += dMasonReg; currentSiteData[existingSiteIdx].masonOT += dMasonOT;
            currentSiteData[existingSiteIdx].halfMasonReg += dHMMasonReg; currentSiteData[existingSiteIdx].halfMasonOT += dHMMasonOT;
            currentSiteData[existingSiteIdx].helperReg += dHelperReg; currentSiteData[existingSiteIdx].helperOT += dHelperOT;
            currentSiteData[existingSiteIdx].totalBaseCost += dTotalBaseCost; currentSiteData[existingSiteIdx].totalOTCost += dTotalOTCost;
          } else { currentSiteData.push(todaySiteData); }

          logWorkers.forEach(newW => {
            const wIdx = currentWorkerData.findIndex(w => w.worker === newW.worker && w.contractor === supContractor);
            if (wIdx > -1) {
              currentWorkerData[wIdx].regDays += newW.regDays; currentWorkerData[wIdx].otHours += newW.otHours;
              currentWorkerData[wIdx].totalBaseCost += newW.totalBaseCost; currentWorkerData[wIdx].totalOTCost += newW.totalOTCost;
            } else { currentWorkerData.push(newW); }
          });
        }

        transaction.set(sheetRef, { sheetName: period.displayName, siteData: currentSiteData, dayData: currentDayData, workerData: currentWorkerData, updatedAt: Date.now() });
      });

      if (editingLog) {
        const newEditCount = (editingLog.editCount || 0) + 1;
        await updateDoc(doc(db, COLL_LOGS, editingLog.id), {
          workers: logWorkers,
          isEdited: true,
          editTimestamp: Date.now(),
          editCount: newEditCount
        });

        if (!isToday(editingLog.date)) {
          const reqsToClose = myRequests.filter(r => r.logId === editingLog.id && (r.status === 'approved' || r.status === 'pending'));
          for (const req of reqsToClose) {
            await updateDoc(doc(db, COLL_REQS, req.id), { status: 'completed' });
          }
        }
        alert(`Attendance securely updated!`);
      } else {
        const newLogRef = await addDoc(collection(db, COLL_LOGS), {
          date: supDate, site: supSite, contractor: supContractor, workers: logWorkers,
          timestamp: Date.now(), submittedBy: currentUser?.email || "Unknown User"
        });
        createdLogId = newLogRef.id;
        // Trigger the sleek mobile toast instead of the ugly browser alert
        setToastMessage("✅ Attendance Secured Successfully!");
        setTimeout(() => setToastMessage(""), 3000);
      }

      // --- SEND PROVISIONAL WORKER ADMIN REQUESTS ---
      const provWorkers = logWorkers.filter(w => w.isProvisional);
      for (const pw of provWorkers) {
        await addDoc(collection(db, COLL_REQS), {
          logId: createdLogId,
          date: supDate,
          site: supSite,
          contractor: supContractor,
          submittedBy: currentUser.email,
          timestamp: Date.now(),
          status: 'pending',
          type: 'new_worker',
          workerName: pw.worker,
          workerType: pw.type
        });
      }

      setSupContractor(""); setWorkerSearch(""); setEditingLog(null);
      setActivePortalTab('history');
    } catch (err) { alert("Error saving attendance."); console.error(err); }
    setIsSubmittingLog(false);
  };

  const startEditLog = (log) => {
    setSupDate(log.date); setSupSite(log.site); setSupContractor(log.contractor);
    const contractorWorkers = masterWorkers.filter(w => w.contractor === log.contractor);

    // Make sure we bring provisional workers back locally so they don't disappear on edit
    const provisionalInLog = log.workers.filter(w => w.isProvisional);
    const combinedWorkers = [...contractorWorkers];

    provisionalInLog.forEach(pw => {
      if (!combinedWorkers.some(cw => cw.name === pw.worker)) {
        combinedWorkers.push({ name: pw.worker, type: pw.type, contractor: log.contractor, wage: 0, otRate: 0, isProvisional: true });
      }
    });
    setMasterWorkers(prev => {
      const allOthers = prev.filter(p => p.contractor !== log.contractor);
      return [...allOthers, ...combinedWorkers];
    });

    const restoredState = combinedWorkers.reduce((acc, w) => {
      const logged = log.workers.find(lw => lw.worker === w.name);
      if (logged) {
        let st = 'present';
        if (logged.regDays === 0.5) st = 'half';
        else if (logged.regDays === 2.0) st = '2P';
        acc[w.name] = { status: st, ot: logged.otHours || '' };
      } else { acc[w.name] = { status: 'absent', ot: '' }; }
      return acc;
    }, {});

    setSupAttendance(restoredState);
    setEditingLog(log);
    setActivePortalTab('attendance');
  };

  const handleRequestEdit = async (log) => {
    try {
      await addDoc(collection(db, COLL_REQS), {
        logId: log.id, date: log.date, site: log.site, contractor: log.contractor,
        submittedBy: currentUser.email, timestamp: Date.now(), status: 'pending', type: 'edit'
      });
      fetchHistoryData();
      alert("Edit request sent to Master Admin successfully!");
    } catch (err) { alert("Error sending edit request."); }
  };

  const handleRequestBackdate = async () => {
    try {
      await addDoc(collection(db, COLL_REQS), {
        date: supDate, submittedBy: currentUser.email, timestamp: Date.now(), status: 'pending', type: 'backdate'
      });
      fetchHistoryData();
      alert(`Permission request for ${supDate} sent to Admin! Check back soon.`);
    } catch (err) { alert("Error sending request."); }
  };

  const cancelEdit = () => {
    setEditingLog(null); setSupContractor(""); setSupSite(""); setSupDate(new Date().toISOString().split('T')[0]);
  };

  const renderCategoryUI = (list, title, colorClass, dotColor) => {
    if (list.length === 0) return null;
    return (
      <div className="space-y-2">
        <h4 className={`text-[10px] font-black ${colorClass} uppercase tracking-widest border-b border-gray-100 pb-1.5 flex items-center gap-1.5`}><div className={`w-1.5 h-1.5 rounded-full ${dotColor}`}></div> {title}</h4>
        {list.map(worker => {
          const rec = supAttendance[worker.name] || {};
          const is2P = rec.status === '2P'; const isP = rec.status === 'present'; const isHD = rec.status === 'half'; const isA = rec.status === 'absent';

          // UPGRADE 1: Tint the entire card green if the worker is present
          const cardBg = isA ? 'bg-gray-50/50 border-gray-100' : 'bg-emerald-50/40 border-emerald-300 shadow-sm';

          return (
            <div key={worker.name} className={`px-2 md:px-3 py-2.5 rounded-xl border transition-all flex flex-col md:flex-row items-stretch md:items-center justify-between gap-2 md:gap-4 ${cardBg} ${worker.isProvisional ? 'border-dashed border-orange-300' : ''}`}>
              <div>
                <p className="font-black text-gray-900 text-sm min-w-0 truncate">{worker.name}</p>
                {worker.isProvisional && <p className="text-[9px] font-bold text-orange-500 uppercase tracking-widest mt-0.5">Pending Admin Approval</p>}
              </div>
              <div className="flex items-center justify-between md:justify-end gap-1.5 shrink-0">
                <div className="flex bg-gray-100/80 p-0.5 rounded-lg border border-gray-200/50 flex-1 md:flex-none justify-between shadow-inner">
                  {/* UPGRADE 2: High Contrast Active States */}
                  <button onClick={() => handleAttendanceChange(worker.name, '2P')} className={`flex-1 px-2.5 py-1.5 rounded-md text-[10px] font-black transition-all ${is2P ? 'bg-indigo-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>2P</button>
                  <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`flex-1 px-2.5 py-1.5 rounded-md text-[10px] font-black transition-all ${isP ? 'bg-emerald-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                  <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`flex-1 px-2.5 py-1.5 rounded-md text-[10px] font-black transition-all ${isHD ? 'bg-yellow-500 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                  <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`flex-1 px-2.5 py-1.5 rounded-md text-[10px] font-black transition-all ${isA ? 'bg-white text-gray-900 border border-gray-300 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                </div>
                <div className="w-16 shrink-0 relative">
                  {/* UPGRADE 3: inputMode="decimal" forces the phone number pad to open instead of the QWERTY keyboard */}
                  <input type="number" inputMode="decimal" pattern="[0-9]*" placeholder="Hrs" value={rec.ot || ''} max="12" min="0" onChange={(e) => handleOTChange(worker.name, e.target.value)} className={`w-full h-8 px-1 rounded-lg text-xs font-bold outline-none focus:ring-2 focus:ring-emerald-500/20 text-center transition-all ${rec.ot ? 'bg-emerald-100 text-emerald-900 border border-emerald-400' : 'bg-white border border-gray-200 text-gray-800'}`} />
                </div>
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  const selectedSiteFull = masterSites.find(s => getSiteCode(s) === supSite);
  const displaySelectedSite = selectedSiteFull ? getSiteDisplay(selectedSiteFull) : (supSite || 'Choose Site...');

  return (
    <div className="min-h-screen bg-[#f8fafc] text-gray-800 font-sans selection:bg-emerald-100 selection:text-emerald-900 pb-36">
      <header className="bg-white border-b border-gray-200 px-4 py-3 flex justify-between items-center shadow-sm sticky top-0 z-50">
        <div className="flex items-center gap-2" title="CRM_FIX">
          <div className="p-1.5 rounded-lg bg-emerald-600"><LayoutGrid className="w-4 h-4 text-white" /></div>
          <span className="font-black text-gray-900 text-sm tracking-tight">CRM_FIX <span className="text-gray-400 font-medium hidden sm:inline">| Field Portal</span></span>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-bold shadow-inner">
            <HardHat className="w-3.5 h-3.5 mr-1" /> {currentUser?.email?.split('@')[0]}
          </div>
          <button onClick={() => window.location.reload()} className="text-xs font-bold bg-gray-100 hover:bg-blue-50 hover:text-blue-600 text-gray-600 p-2 rounded-xl transition-colors" title="Sync Data">
            <RefreshCw className="w-4 h-4" />
          </button>
          <button onClick={onLogout} className="text-xs font-bold bg-gray-100 hover:bg-red-50 hover:text-red-600 text-gray-600 p-2 rounded-xl transition-colors" title="Sign Out">
            <LogOut className="w-4 h-4" />
          </button>
        </div>
      </header>
      {/* MOBILE-OPTIMIZED FLOATING TOAST */}
      {toastMessage && (
        <div className="fixed top-6 left-1/2 -translate-x-1/2 z-[100] animate-in slide-in-from-top-4 fade-in duration-300 w-[90%] max-w-sm">
          <div className="bg-gray-900/95 backdrop-blur-md text-white px-5 py-3.5 rounded-2xl shadow-2xl font-black text-sm flex items-center justify-center gap-2 text-center border border-gray-700">
            {toastMessage}
          </div>
        </div>
      )}
      {isTestMode && (
        <div className="bg-yellow-400 text-yellow-900 py-2 px-4 text-center text-[11px] font-black uppercase tracking-widest shadow-md flex items-center justify-center gap-2 sticky top-[60px] z-40">
          <AlertCircle className="w-4 h-4" /> TEST MODE ACTIVE — DATA WILL NOT AFFECT LIVE FINANCIALS
        </div>
      )}

      {isOverdue && !editingLog && (
        <div className="bg-red-500 text-white text-[10px] md:text-xs font-black text-center py-2.5 px-4 shadow-md animate-pulse flex items-center justify-center gap-1.5">
          <AlertCircle className="w-4 h-4" /> {overdueMessage}
        </div>
      )}

      {activePortalTab === 'attendance' ? (
        <div className="max-w-3xl mx-auto pt-4 md:pt-8 px-2 md:px-4 space-y-4 animate-in fade-in slide-in-from-bottom-2 duration-300">
          <div className={`bg-white rounded-[1.5rem] md:rounded-[2rem] p-4 md:p-8 shadow-sm border ${editingLog ? 'border-yellow-400 shadow-yellow-100' : 'border-gray-100'} space-y-4 md:space-y-6 relative`}>

            {editingLog && (
              <div className="absolute top-0 left-0 w-full bg-yellow-400 text-yellow-900 text-[10px] font-black uppercase tracking-widest text-center py-1 rounded-t-[1.5rem] md:rounded-t-[2rem]">
                Editing Mode Active
              </div>
            )}

            <div className={`grid grid-cols-1 md:grid-cols-3 gap-3 md:gap-4 mb-2 ${editingLog ? 'mt-4' : ''}`}>
              <div className="space-y-1.5 relative z-30">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Work Date</label>
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Calendar className="h-4 w-4 text-emerald-500" /></div>
                  <input type="date" disabled={editingLog} max={new Date().toISOString().split('T')[0]} value={supDate} onChange={(e) => setSupDate(e.target.value)} className="w-full bg-gray-50 hover:bg-white border border-gray-200 pl-10 pr-3 py-3 rounded-xl text-sm font-black text-gray-800 outline-none focus:ring-2 focus:ring-emerald-500/20 shadow-sm transition-all cursor-pointer relative [&::-webkit-calendar-picker-indicator]:opacity-0 [&::-webkit-calendar-picker-indicator]:absolute [&::-webkit-calendar-picker-indicator]:inset-0 [&::-webkit-calendar-picker-indicator]:w-full [&::-webkit-calendar-picker-indicator]:h-full [&::-webkit-calendar-picker-indicator]:cursor-pointer z-10 bg-transparent" />
                </div>
                {supDate < new Date().toISOString().split('T')[0] && !editingLog && (
                  <div className="mt-1 ml-1 text-[10px] font-bold">
                    {myRequests.find(r => r.date === supDate && r.status === 'approved' && r.type === 'backdate') ? (
                      <span className="text-emerald-600">✅ Unlocked by Admin! You can submit.</span>
                    ) : myRequests.find(r => r.date === supDate && r.status === 'pending' && r.type === 'backdate') ? (
                      <span className="text-orange-500">⏳ Request pending admin approval...</span>
                    ) : (
                      <button onClick={handleRequestBackdate} type="button" className="text-blue-600 hover:underline">Request Past Date Unlock</button>
                    )}
                  </div>
                )}
              </div>

              <div className="space-y-1.5 relative z-50">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Select Site</label>
                <div className="relative">
                  {isSiteDropdownOpen && <div className="fixed inset-0 z-40" onClick={() => setIsSiteDropdownOpen(false)}></div>}
                  <div className={`w-full ${editingLog ? 'bg-gray-100 cursor-not-allowed' : 'bg-gray-50 hover:bg-white cursor-pointer'} border border-gray-200 pl-10 pr-10 py-3 rounded-xl text-sm font-black text-gray-800 shadow-sm transition-all relative z-50 flex items-center`} onClick={() => { if (!editingLog) { setIsSiteDropdownOpen(!isSiteDropdownOpen); setIsTeamDropdownOpen(false); setSiteSearchQuery(""); } }}>
                    <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><MapPin className={`h-4 w-4 ${supSite ? 'text-emerald-500' : 'text-gray-400'}`} /></div>
                    <span className={supSite ? 'text-gray-900 truncate' : 'text-gray-400'}>{displaySelectedSite}</span>
                    <div className="absolute inset-y-0 right-0 pr-3.5 flex items-center pointer-events-none"><ChevronDown className={`h-4 w-4 transition-transform duration-200 ${isSiteDropdownOpen ? 'rotate-180 text-emerald-500' : 'text-gray-400'}`} /></div>
                  </div>
                  <div className={`absolute top-[calc(100%+8px)] left-0 w-full bg-white border border-gray-100 rounded-xl shadow-xl z-[60] overflow-hidden transition-all duration-200 origin-top ${isSiteDropdownOpen ? 'scale-y-100 opacity-100' : 'scale-y-0 opacity-0 pointer-events-none'}`}>
                    <div className="p-2 border-b border-gray-50 bg-gray-50/50">
                      <div className="relative">
                        <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                        <input type="text" placeholder="Search site..." value={siteSearchQuery} onChange={(e) => setSiteSearchQuery(e.target.value)} onClick={(e) => e.stopPropagation()} className="w-full pl-8 pr-3 py-2 bg-white border border-gray-200 rounded-lg text-xs font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20" />
                      </div>
                    </div>
                    <div className="max-h-52 overflow-y-auto custom-scrollbar py-1">
                      {masterSites.filter(site => getSiteDisplay(site).toLowerCase().includes(siteSearchQuery.toLowerCase())).map(site => {
                        const code = getSiteCode(site);
                        const display = getSiteDisplay(site);
                        return (
                          <div key={code} onClick={() => { setSupSite(code); setIsSiteDropdownOpen(false); setSiteSearchQuery(""); }} className={`px-4 py-2.5 text-sm font-black cursor-pointer transition-colors ${supSite === code ? 'bg-emerald-50 text-emerald-700' : 'text-gray-600 hover:bg-gray-50 hover:text-blue-600'}`}>
                            {display}
                          </div>
                        )
                      })}
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-1.5 relative z-40">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Contractor Team</label>
                <div className="relative">
                  {isTeamDropdownOpen && <div className="fixed inset-0 z-40" onClick={() => setIsTeamDropdownOpen(false)}></div>}
                  <div className={`w-full ${editingLog ? 'bg-gray-100 cursor-not-allowed' : 'bg-gray-50 hover:bg-white cursor-pointer'} border border-gray-200 pl-10 pr-10 py-3 rounded-xl text-sm font-black text-gray-800 shadow-sm transition-all relative z-50 flex items-center`} onClick={() => { if (!editingLog) { setIsTeamDropdownOpen(!isTeamDropdownOpen); setIsSiteDropdownOpen(false); } }}>
                    <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Users className={`h-4 w-4 ${supContractor ? 'text-emerald-500' : 'text-gray-400'}`} /></div>
                    <span className={supContractor ? 'text-gray-900 truncate' : 'text-gray-400'}>{supContractor ? `${supContractor}'s Team` : 'Select Team...'}</span>
                    <div className="absolute inset-y-0 right-0 pr-3.5 flex items-center pointer-events-none"><ChevronDown className={`h-4 w-4 transition-transform duration-200 ${isTeamDropdownOpen ? 'rotate-180 text-emerald-500' : 'text-gray-400'}`} /></div>
                  </div>
                  <div className={`absolute top-[calc(100%+8px)] left-0 w-full bg-white border border-gray-100 rounded-xl shadow-xl z-[60] overflow-hidden transition-all duration-200 origin-top ${isTeamDropdownOpen ? 'scale-y-100 opacity-100' : 'scale-y-0 opacity-0 pointer-events-none'}`}>
                    <div className="max-h-60 overflow-y-auto custom-scrollbar py-2">
                      {dynamicContractors.map(contractor => (
                        <div key={contractor} onClick={() => { setSupContractor(contractor); setIsTeamDropdownOpen(false); }} className={`px-4 py-3 text-sm font-black cursor-pointer transition-colors ${supContractor === contractor ? 'bg-emerald-50 text-emerald-700' : 'text-gray-600 hover:bg-gray-50 hover:text-blue-600'}`}>{contractor}'s Team</div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {supContractor && (
              <div className="pt-4 md:pt-6 border-t border-gray-100">
                <div className="flex flex-col sm:flex-row justify-between items-stretch sm:items-center gap-2 mb-4">
                  <div className="flex items-center justify-between sm:justify-start gap-3">
                    <h3 className="font-black text-gray-900 text-base md:text-lg">Roster <span className="text-gray-400">({activeContractorWorkers.length})</span></h3>
                  </div>
                  <div className="relative w-full sm:w-48 md:w-64">
                    <Search className="w-3.5 h-3.5 md:w-4 md:h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input type="text" placeholder="Search name..." value={workerSearch} onChange={(e) => setWorkerSearch(e.target.value)} className="w-full pl-8 md:pl-9 pr-3 md:pr-4 py-2 bg-gray-50 border border-gray-200 rounded-lg md:rounded-xl text-xs md:text-sm font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:bg-white" />
                  </div>
                </div>
                {/* 1. THE ACTIVE CART (Selected Workers Box) */}
                {activeContractorWorkers.filter(w => supAttendance[w.name] && (supAttendance[w.name].status !== 'absent' || parseFloat(supAttendance[w.name].ot) > 0)).length > 0 && (
                  <div className="bg-emerald-50/50 border border-emerald-100 rounded-xl p-3 mb-4 max-h-32 overflow-y-auto custom-scrollbar shadow-inner">
                    <p className="text-[9px] font-black text-emerald-600 uppercase tracking-widest mb-2">Marked Today ({activeContractorWorkers.filter(w => supAttendance[w.name] && (supAttendance[w.name].status !== 'absent' || parseFloat(supAttendance[w.name].ot) > 0)).length})</p>
                    <div className="flex flex-wrap gap-1.5">
                      {activeContractorWorkers.filter(w => supAttendance[w.name] && (supAttendance[w.name].status !== 'absent' || parseFloat(supAttendance[w.name].ot) > 0)).map(w => {
                        const rec = supAttendance[w.name];
                        const val = rec.status === 'present' ? 'P' : rec.status === 'half' ? 'HD' : rec.status === '2P' ? '2P' : '';
                        const otVal = parseFloat(rec.ot) > 0 ? `+${rec.ot}h` : '';
                        return (
                          <div key={w.name} className="bg-white border border-emerald-200 text-gray-800 text-[10px] font-black pl-2 pr-1 py-1 rounded-lg shadow-sm flex items-center gap-1.5">
                            <span className="truncate max-w-[80px] sm:max-w-[100px]">{w.name}</span>
                            <span className="text-emerald-600">{val}{otVal}</span>
                            <button onClick={() => { handleAttendanceChange(w.name, 'absent'); handleOTChange(w.name, ''); }} className="text-gray-400 hover:text-red-500 ml-0.5 bg-gray-50 hover:bg-red-50 rounded-md p-0.5"><X className="w-3 h-3" /></button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="space-y-6">
                  {renderCategoryUI(masons, "MASONS", "text-blue-500", "bg-blue-500")}
                  {renderCategoryUI(halfMasons, "HALF MASONS", "text-purple-500", "bg-purple-500")}
                  {renderCategoryUI(helpers, "HELPERS", "text-orange-500", "bg-orange-500")}
                </div>

                {/* 2. THE MOVED ADD WORKER FORM */}
                {showAddWorker ? (
                  <div className="mt-6 bg-blue-50/50 p-4 rounded-[1.5rem] border border-blue-200 shadow-inner flex flex-col sm:flex-row gap-3 animate-in fade-in slide-in-from-bottom-2">
                    <input type="text" placeholder="Worker Name..." value={newWorkerName} onChange={(e) => setNewWorkerName(e.target.value)} className="flex-1 px-4 py-3 text-sm font-bold rounded-xl border border-blue-200 focus:outline-none focus:ring-2 focus:ring-blue-500/30" />
                    <select value={newWorkerType} onChange={(e) => setNewWorkerType(e.target.value)} className="px-4 py-3 text-sm font-bold rounded-xl border border-blue-200 focus:outline-none focus:ring-2 focus:ring-blue-500/30 bg-white cursor-pointer">
                      <option value="Helper">Helper</option>
                      <option value="HalfMason">Half Mason</option>
                      <option value="Mason">Mason</option>
                    </select>
                    <div className="flex gap-2">
                      <button onClick={() => setShowAddWorker(false)} className="px-4 py-3 text-xs font-black text-gray-500 hover:bg-white border border-transparent hover:border-gray-200 rounded-xl transition-all flex-1 sm:flex-none">Cancel</button>
                      <button onClick={handleAddNewWorkerLocal} className="px-6 py-3 text-xs font-black bg-blue-600 hover:bg-blue-700 text-white rounded-xl shadow-md transition-all flex-1 sm:flex-none">Save</button>
                    </div>
                  </div>
                ) : (
                  <div className="mt-6 flex justify-center w-full">
                    <button onClick={() => setShowAddWorker(true)} className="w-full text-xs font-bold text-gray-500 hover:text-blue-600 bg-white active:bg-blue-50 py-3 rounded-2xl border-2 border-gray-200 border-dashed transition-all flex items-center justify-center gap-2">
                      Worker missing? <span className="text-blue-600">Register New Profile</span>
                    </button>
                  </div>
                )}

                {/* 3. THE BULLETPROOF FLOATING SAVE BAR */}
                <div className="fixed bottom-[60px] md:bottom-0 left-0 w-full p-4 z-40 bg-gradient-to-t from-white via-white to-transparent pointer-events-none flex justify-center pb-6">
                  <div className="w-full max-w-3xl pointer-events-auto flex flex-col md:flex-row gap-3 px-2">
                    {editingLog && (
                      <button onClick={cancelEdit} className="w-full md:w-auto font-black py-3.5 px-6 rounded-xl border-2 border-gray-200 bg-white text-gray-500 hover:bg-gray-50 transition-colors shadow-sm">
                        Cancel Edit
                      </button>
                    )}
                    <button onClick={submitDailyLog} disabled={isSubmittingLog || activeContractorWorkers.length === 0} className={`w-full flex-1 font-black py-3.5 rounded-xl flex items-center justify-center gap-2 shadow-lg transition-all ${isSubmittingLog || activeContractorWorkers.length === 0 ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : editingLog ? 'bg-yellow-500 hover:bg-yellow-600 text-white shadow-yellow-500/30' : 'bg-gray-900 hover:bg-emerald-600 text-white shadow-gray-900/20'}`}>
                      {editingLog ? <Edit2 className="w-4 h-4" /> : <Save className="w-4 h-4" />}
                      {isSubmittingLog ? 'Saving...' : editingLog ? "Update Attendance" : "Secure Today's Attendance"}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      ) : (
        <div className="max-w-3xl mx-auto pt-4 md:pt-8 px-2 md:px-4 space-y-4 animate-in fade-in slide-in-from-right-2 duration-300">
          <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] p-4 md:p-8 shadow-sm border border-gray-100">
            <h2 className="text-xl md:text-2xl font-black text-gray-900 mb-6 flex items-center gap-2">
              <Clock className="w-6 h-6 text-emerald-500" /> My Submission History
            </h2>

            <div className="space-y-3">
              {myLogs.length === 0 ? (
                <div className="text-center py-10 bg-gray-50 rounded-2xl border border-gray-100">
                  <p className="text-gray-400 font-bold text-sm">You haven't submitted any logs yet.</p>
                </div>
              ) : (
                myLogs.map(log => {
                  const approvedReq = myRequests.find(r => r.logId === log.id && r.status === 'approved');
                  const pendingReq = myRequests.find(r => r.logId === log.id && r.status === 'pending');
                  const editCount = log.editCount || 0;
                  const canEdit = (isToday(log.date) && editCount < 3) || approvedReq;

                  const sFull = masterSites.find(s => getSiteCode(s) === log.site);
                  const displayLogSite = sFull ? getSiteDisplay(sFull) : log.site;

                  return (
                    <div key={log.id} className="bg-white p-4 rounded-2xl border border-gray-200 shadow-sm flex flex-col hover:border-emerald-200 transition-colors">
                      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                        <div>
                          <div className="flex items-center gap-2 mb-1">
                            <span className="bg-emerald-100 text-emerald-700 text-[10px] font-black px-2 py-0.5 rounded uppercase tracking-wider">
                              {log.date}
                            </span>
                            {log.isEdited && <span className="bg-yellow-100 text-yellow-800 text-[9px] font-black px-1.5 py-0.5 rounded uppercase tracking-wider">Edited {editCount > 0 ? `(${editCount}/3)` : ''}</span>}
                            {log.timestamp && <span className="text-[10px] text-gray-400 font-bold ml-1">{new Date(log.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>}
                          </div>
                          <h4 className="font-black text-gray-900 text-base">{displayLogSite}</h4>
                          <p className="text-xs text-gray-500 font-bold flex items-center gap-1 mt-0.5">
                            <Users className="w-3 h-3" /> {log.contractor}'s Team <span className="text-gray-300 mx-1">|</span> {log.workers?.length || 0} Present
                          </p>
                        </div>

                        <div className="w-full sm:w-auto flex flex-wrap gap-2">
                          <button onClick={() => setExpandedLogId(expandedLogId === log.id ? null : log.id)} className={`flex-1 sm:flex-none px-3 py-2 font-black text-xs rounded-xl transition-colors flex items-center justify-center gap-1.5 border ${expandedLogId === log.id ? 'bg-gray-800 text-white border-gray-800' : 'bg-white text-gray-600 border-gray-200 hover:bg-gray-50'}`}>
                            {expandedLogId === log.id ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                            {expandedLogId === log.id ? 'Hide Roster' : 'View Roster'}
                          </button>

                          {canEdit ? (
                            <button onClick={() => startEditLog(log)} className="flex-1 sm:flex-none px-4 py-2 bg-blue-50 hover:bg-blue-100 text-blue-700 font-black text-xs rounded-xl transition-colors flex items-center justify-center gap-1.5">
                              {approvedReq ? <Edit2 className="w-3.5 h-3.5" /> : <Edit2 className="w-3.5 h-3.5" />}
                              {approvedReq ? 'Edit Unlocked' : 'Edit Log'}
                            </button>
                          ) : pendingReq ? (
                            <button disabled className="flex-1 sm:flex-none px-4 py-2 bg-gray-100 text-gray-400 font-black text-xs rounded-xl cursor-not-allowed flex items-center justify-center gap-1.5 opacity-70">
                              <Clock className="w-3.5 h-3.5" /> Pending
                            </button>
                          ) : (
                            <button onClick={() => handleRequestEdit(log)} className="flex-1 sm:flex-none px-4 py-2 bg-orange-50 hover:bg-orange-100 text-orange-700 font-black text-xs rounded-xl transition-colors flex items-center justify-center gap-1.5">
                              <AlertCircle className="w-3.5 h-3.5" /> Request Edit
                            </button>
                          )}
                        </div>
                      </div>

                      {expandedLogId === log.id && log.workers && (
                        <div className="mt-4 pt-4 border-t border-gray-100 animate-in slide-in-from-top-2 fade-in duration-300">
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
                                    <div key={i} className="flex justify-between items-center bg-gray-50/80 px-3 py-2 rounded-lg border border-gray-100">
                                      <div>
                                        <span className="text-xs font-black text-gray-800">{w.worker}</span>
                                        {w.isProvisional && <span className="ml-1 text-[8px] bg-orange-100 text-orange-600 px-1 py-0.5 rounded font-black uppercase">Pending</span>}
                                      </div>
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
                  );
                })
              )}
            </div>
          </div>
        </div>
      )}

      <div className="fixed bottom-0 left-0 w-full bg-white border-t border-gray-200 px-4 py-3 shadow-[0_-4px_20px_-10px_rgba(0,0,0,0.1)] z-50 md:flex md:justify-center">
        <div className="flex max-w-md w-full mx-auto gap-2 bg-gray-100 p-1 rounded-2xl">
          <button onClick={() => setActivePortalTab('attendance')} className={`flex-1 flex flex-col items-center justify-center py-2 rounded-xl transition-all ${activePortalTab === 'attendance' ? 'bg-white text-emerald-600 shadow-sm' : 'text-gray-500 hover:text-gray-900'}`}>
            <LayoutGrid className="w-5 h-5 mb-0.5" />
            <span className="text-[10px] font-black uppercase tracking-wider">Attendance</span>
          </button>
          <button onClick={() => setActivePortalTab('history')} className={`flex-1 flex flex-col items-center justify-center py-2 rounded-xl transition-all ${activePortalTab === 'history' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-900'}`}>
            <Clock className="w-5 h-5 mb-0.5" />
            <span className="text-[10px] font-black uppercase tracking-wider">My Logs</span>
          </button>
        </div>
      </div>
    </div>
  );
}
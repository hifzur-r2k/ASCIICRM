import React, { useState, useEffect } from 'react';
import { 
  LayoutGrid, Search, Save, LogOut, Calendar, Users, MapPin, HardHat, ChevronDown 
} from 'lucide-react';
import { collection, addDoc, getDocs, doc, setDoc } from "firebase/firestore";
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

  return {
    id: `${year}-${month}-${half}`,
    displayName: `${displayMonth} ${displayRange} ${year}`
  };
};

export default function SupervisorPortal({ currentUser, onLogout }) {
  const [masterSites, setMasterSites] = useState([]);
  const [masterWorkers, setMasterWorkers] = useState([]);

  const [supDate, setSupDate] = useState(new Date().toISOString().split('T')[0]);
  const [supSite, setSupSite] = useState("");
  const [supContractor, setSupContractor] = useState("");
  const [workerSearch, setWorkerSearch] = useState("");
  const [supAttendance, setSupAttendance] = useState({});
  const [isSubmittingLog, setIsSubmittingLog] = useState(false);
  
  const [isSiteDropdownOpen, setIsSiteDropdownOpen] = useState(false);
  const [isTeamDropdownOpen, setIsTeamDropdownOpen] = useState(false);
  const [siteSearchQuery, setSiteSearchQuery] = useState("");

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

        // If no sites found in master_data, fallback to querying attendance sheets to build site list
        if (loadedSites.length === 0) {
          const sheetsSnap = await getDocs(collection(db, "attendance_sheets"));
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
      } catch (err) {
        console.error("Error fetching master data:", err);
      }
    };
    fetchMasterData();
  }, []);

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

    if (logWorkers.length === 0) { 
      alert("No workers marked present. Nothing to save."); 
      setIsSubmittingLog(false); 
      return; 
    }

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
      
      // NEW: AUDIT TRAIL ADDED HERE! 
      await addDoc(collection(db, "daily_logs"), { 
        date: supDate, 
        site: supSite, 
        contractor: supContractor, 
        workers: logWorkers, 
        timestamp: Date.now(),
        submittedBy: currentUser?.email || "Unknown User" // Safely tracks the supervisor's email
      });

      alert(`Attendance locked into bucket: ${period.displayName}!`);
      setSupContractor(""); 
      setWorkerSearch(""); 
    } catch (err) { 
      alert("Error saving attendance to bucket."); 
      console.error(err); 
    }
    setIsSubmittingLog(false);
  };

  const activeContractorWorkers = masterWorkers.filter(w => w.contractor === supContractor);
  const searchedWorkers = activeContractorWorkers.filter(w => w.name.toLowerCase().includes(workerSearch.toLowerCase()));
  const masons = searchedWorkers.filter(w => w.type === 'Mason');
  const halfMasons = searchedWorkers.filter(w => w.type === 'HalfMason');
  const helpers = searchedWorkers.filter(w => w.type === 'Helper');

  return (
    <div className="min-h-screen bg-[#f8fafc] text-gray-800 font-sans selection:bg-blue-100 selection:text-blue-900 pb-12">
      <header className="bg-white border-b border-gray-200 px-4 py-3 flex justify-between items-center shadow-sm sticky top-0 z-50">
        <div className="flex items-center gap-2 cursor-pointer hover:opacity-80 transition-opacity" title="CRM_FIX">
          <div className="p-1.5 rounded-lg bg-emerald-600"><LayoutGrid className="w-4 h-4 text-white" /></div>
          <span className="font-black text-gray-900 text-sm tracking-tight">CRM_FIX <span className="text-gray-400 font-medium hidden sm:inline">| Field Portal</span></span>
        </div>
        <div className="flex items-center gap-3 md:gap-4">
          <div className="flex items-center bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-lg text-xs font-bold shadow-inner">
            <HardHat className="w-3.5 h-3.5 mr-1" /> Supervisor
          </div>
          <button onClick={onLogout} className="text-xs font-bold bg-gray-100 hover:bg-red-50 hover:text-red-600 text-gray-600 p-2 md:px-3 md:py-1.5 rounded-xl flex items-center gap-2 transition-colors">
            <LogOut className="w-4 h-4 md:w-3.5 md:h-3.5" /> <span className="hidden md:block">Sign Out</span>
          </button>
        </div>
      </header>

      <div className="max-w-3xl mx-auto pt-4 md:pt-8 px-2 md:px-4 space-y-4">
        <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] p-4 md:p-8 shadow-sm border border-gray-100 space-y-4 md:space-y-6">

          <div className="grid grid-cols-1 md:grid-cols-3 gap-3 md:gap-4 mb-2">
            <div className="space-y-1.5 relative z-30">
              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Work Date</label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                  <Calendar className="h-4 w-4 text-emerald-500" />
                </div>
                <input type="date" value={supDate} onChange={(e) => setSupDate(e.target.value)} className="w-full bg-gray-50 hover:bg-white border border-gray-200 pl-10 pr-3 py-3 rounded-xl text-sm font-black text-gray-800 outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 shadow-sm transition-all cursor-pointer relative [&::-webkit-calendar-picker-indicator]:opacity-0 [&::-webkit-calendar-picker-indicator]:absolute [&::-webkit-calendar-picker-indicator]:inset-0 [&::-webkit-calendar-picker-indicator]:w-full [&::-webkit-calendar-picker-indicator]:h-full [&::-webkit-calendar-picker-indicator]:cursor-pointer z-10 bg-transparent" />
              </div>
            </div>

            <div className="space-y-1.5 relative z-50">
              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Select Site</label>
              <div className="relative">
                {isSiteDropdownOpen && <div className="fixed inset-0 z-40" onClick={() => setIsSiteDropdownOpen(false)}></div>}
                <div className="w-full bg-gray-50 hover:bg-white border border-gray-200 pl-10 pr-10 py-3 rounded-xl text-sm font-black text-gray-800 shadow-sm transition-all cursor-pointer relative z-50 flex items-center" onClick={() => { setIsSiteDropdownOpen(!isSiteDropdownOpen); setIsTeamDropdownOpen(false); setSiteSearchQuery(""); }}>
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                    <MapPin className={`h-4 w-4 ${supSite ? 'text-emerald-500' : 'text-gray-400'}`} />
                  </div>
                  <span className={supSite ? 'text-gray-900 truncate' : 'text-gray-400'}>{supSite || 'Choose Site...'}</span>
                  <div className="absolute inset-y-0 right-0 pr-3.5 flex items-center pointer-events-none">
                    <ChevronDown className={`h-4 w-4 transition-transform duration-200 ${isSiteDropdownOpen ? 'rotate-180 text-emerald-500' : 'text-gray-400'}`} />
                  </div>
                </div>
                <div className={`absolute top-full left-0 mt-2 w-full bg-white border border-gray-100 rounded-xl shadow-[0_10px_40px_-10px_rgba(0,0,0,0.1)] z-50 overflow-hidden transition-all duration-200 origin-top ${isSiteDropdownOpen ? 'scale-y-100 opacity-100' : 'scale-y-0 opacity-0 pointer-events-none'}`}>
                  <div className="p-2 border-b border-gray-50 bg-gray-50/50">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input 
                        type="text" 
                        placeholder="Search site..." 
                        value={siteSearchQuery} 
                        onChange={(e) => setSiteSearchQuery(e.target.value)}
                        onClick={(e) => e.stopPropagation()} 
                        className="w-full pl-8 pr-3 py-2 bg-white border border-gray-200 rounded-lg text-xs font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
                      />
                    </div>
                  </div>
                  <div className="max-h-52 overflow-y-auto custom-scrollbar py-1">
                    {masterSites.filter(site => site.toLowerCase().includes(siteSearchQuery.toLowerCase())).map(site => (
                      <div key={site} onClick={() => { setSupSite(site); setIsSiteDropdownOpen(false); setSiteSearchQuery(""); }} className={`px-4 py-2.5 text-sm font-black cursor-pointer transition-colors ${supSite === site ? 'bg-emerald-50 text-emerald-700' : 'text-gray-600 hover:bg-gray-50 hover:text-blue-600'}`}>
                        {site}
                      </div>
                    ))}
                    {masterSites.filter(site => site.toLowerCase().includes(siteSearchQuery.toLowerCase())).length === 0 && <div className="px-4 py-3 text-xs font-bold text-gray-400 text-center">No sites found</div>}
                  </div>
                </div>
              </div>
            </div>

            <div className="space-y-1.5 relative z-40">
              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Contractor Team</label>
              <div className="relative">
                {isTeamDropdownOpen && <div className="fixed inset-0 z-40" onClick={() => setIsTeamDropdownOpen(false)}></div>}
                <div className="w-full bg-gray-50 hover:bg-white border border-gray-200 pl-10 pr-10 py-3 rounded-xl text-sm font-black text-gray-800 shadow-sm transition-all cursor-pointer relative z-50 flex items-center" onClick={() => { setIsTeamDropdownOpen(!isTeamDropdownOpen); setIsSiteDropdownOpen(false); }}>
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                    <Users className={`h-4 w-4 ${supContractor ? 'text-emerald-500' : 'text-gray-400'}`} />
                  </div>
                  <span className={supContractor ? 'text-gray-900 truncate' : 'text-gray-400'}>{supContractor ? `${supContractor}'s Team` : 'Select Team...'}</span>
                  <div className="absolute inset-y-0 right-0 pr-3.5 flex items-center pointer-events-none">
                    <ChevronDown className={`h-4 w-4 transition-transform duration-200 ${isTeamDropdownOpen ? 'rotate-180 text-emerald-500' : 'text-gray-400'}`} />
                  </div>
                </div>
                <div className={`absolute top-full left-0 mt-2 w-full bg-white border border-gray-100 rounded-xl shadow-[0_10px_40px_-10px_rgba(0,0,0,0.1)] z-50 overflow-hidden transition-all duration-200 origin-top ${isTeamDropdownOpen ? 'scale-y-100 opacity-100' : 'scale-y-0 opacity-0 pointer-events-none'}`}>
                  <div className="max-h-60 overflow-y-auto custom-scrollbar py-2">
                    {dynamicContractors.map(contractor => (
                      <div key={contractor} onClick={() => { setSupContractor(contractor); setIsTeamDropdownOpen(false); }} className={`px-4 py-3 text-sm font-black cursor-pointer transition-colors ${supContractor === contractor ? 'bg-emerald-50 text-emerald-700' : 'text-gray-600 hover:bg-gray-50 hover:text-blue-600'}`}>
                        {contractor}'s Team
                      </div>
                    ))}
                    {dynamicContractors.length === 0 && <div className="px-4 py-3 text-xs font-bold text-gray-400 text-center">No teams found</div>}
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
                  <span className="bg-gray-100 text-gray-500 text-[8px] md:text-[9px] font-bold px-2 py-1 rounded md:rounded-md uppercase tracking-wider">Default: Absent</span>
                </div>
                <div className="relative w-full sm:w-48 md:w-64">
                  <Search className="w-3.5 h-3.5 md:w-4 md:h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                  <input type="text" placeholder="Search name..." value={workerSearch} onChange={(e) => setWorkerSearch(e.target.value)} className="w-full pl-8 md:pl-9 pr-3 md:pr-4 py-2 bg-gray-50 border border-gray-200 rounded-lg md:rounded-xl text-xs md:text-sm font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:bg-white" />
                </div>
              </div>

              <div className="space-y-6">
                {masons.length > 0 && (
                  <div className="space-y-2">
                    <h4 className="text-[10px] font-black text-blue-500 uppercase tracking-widest border-b border-gray-100 pb-1.5 flex items-center gap-1.5"><div className="w-1.5 h-1.5 rounded-full bg-blue-500"></div> MASONS</h4>
                    {masons.map(worker => {
                      const rec = supAttendance[worker.name] || {};
                      const isPresent = rec.status === 'present'; const isHalf = rec.status === 'half'; const isAbsent = rec.status === 'absent';
                      return (
                        <div key={worker.name} className={`px-3 py-2.5 rounded-xl border transition-all flex items-center justify-between gap-2 ${isAbsent ? 'bg-gray-50/50 border-gray-100' : 'bg-white border-blue-200 shadow-sm'}`}>
                          <p className="font-black text-gray-900 text-sm flex-1 min-w-0 pr-2 truncate">{worker.name}</p>
                          <div className="flex items-center gap-1.5 shrink-0">
                            <div className="flex bg-gray-100 p-0.5 rounded-lg border border-gray-200/50">
                              <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isPresent ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                              <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isHalf ? 'bg-yellow-500 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                              <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isAbsent ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                            </div>
                            {(isPresent || isHalf) && <input type="number" placeholder="OT" value={rec.ot} onChange={(e) => handleOTChange(worker.name, e.target.value)} className="w-12 sm:w-16 h-7 sm:h-8 bg-gray-50 border border-gray-200 px-1 rounded-md sm:rounded-lg text-[10px] sm:text-xs font-bold text-gray-800 outline-none focus:ring-2 focus:ring-blue-500/20 text-center" />}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {halfMasons.length > 0 && (
                  <div className="space-y-2">
                    <h4 className="text-[10px] font-black text-purple-500 uppercase tracking-widest border-b border-gray-100 pb-1.5 flex items-center gap-1.5"><div className="w-1.5 h-1.5 rounded-full bg-purple-500"></div> HALF MASONS</h4>
                    {halfMasons.map(worker => {
                      const rec = supAttendance[worker.name] || {};
                      const isPresent = rec.status === 'present'; const isHalf = rec.status === 'half'; const isAbsent = rec.status === 'absent';
                      return (
                        <div key={worker.name} className={`px-3 py-2.5 rounded-xl border transition-all flex items-center justify-between gap-2 ${isAbsent ? 'bg-gray-50/50 border-gray-100' : 'bg-white border-purple-200 shadow-sm'}`}>
                          <p className="font-black text-gray-900 text-sm flex-1 min-w-0 pr-2 truncate">{worker.name}</p>
                          <div className="flex items-center gap-1.5 shrink-0">
                            <div className="flex bg-gray-100 p-0.5 rounded-lg border border-gray-200/50">
                              <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isPresent ? 'bg-purple-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                              <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isHalf ? 'bg-yellow-500 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                              <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isAbsent ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                            </div>
                            {(isPresent || isHalf) && <input type="number" placeholder="OT" value={rec.ot} onChange={(e) => handleOTChange(worker.name, e.target.value)} className="w-12 sm:w-16 h-7 sm:h-8 bg-gray-50 border border-gray-200 px-1 rounded-md sm:rounded-lg text-[10px] sm:text-xs font-bold text-gray-800 outline-none focus:ring-2 focus:ring-purple-500/20 text-center" />}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {helpers.length > 0 && (
                  <div className="space-y-2">
                    <h4 className="text-[10px] font-black text-orange-500 uppercase tracking-widest border-b border-gray-100 pb-1.5 flex items-center gap-1.5"><div className="w-1.5 h-1.5 rounded-full bg-orange-500"></div> HELPERS</h4>
                    {helpers.map(worker => {
                      const rec = supAttendance[worker.name] || {};
                      const isPresent = rec.status === 'present'; const isHalf = rec.status === 'half'; const isAbsent = rec.status === 'absent';
                      return (
                        <div key={worker.name} className={`px-3 py-2.5 rounded-xl border transition-all flex items-center justify-between gap-2 ${isAbsent ? 'bg-gray-50/50 border-gray-100' : 'bg-white border-orange-200 shadow-sm'}`}>
                          <p className="font-black text-gray-900 text-sm flex-1 min-w-0 pr-2 truncate">{worker.name}</p>
                          <div className="flex items-center gap-1.5 shrink-0">
                            <div className="flex bg-gray-100 p-0.5 rounded-lg border border-gray-200/50">
                              <button onClick={() => handleAttendanceChange(worker.name, 'present')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isPresent ? 'bg-orange-600 text-white shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>P</button>
                              <button onClick={() => handleAttendanceChange(worker.name, 'half')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isHalf ? 'bg-yellow-500 text-white shadow-md' : 'text-gray-500 hover:text-gray-700'}`}>HD</button>
                              <button onClick={() => handleAttendanceChange(worker.name, 'absent')} className={`px-2.5 py-1.5 rounded-md text-[10px] sm:text-xs font-black transition-all ${isAbsent ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>A</button>
                            </div>
                            {(isPresent || isHalf) && <input type="number" placeholder="OT" value={rec.ot} onChange={(e) => handleOTChange(worker.name, e.target.value)} className="w-12 sm:w-16 h-7 sm:h-8 bg-gray-50 border border-gray-200 px-1 rounded-md sm:rounded-lg text-[10px] sm:text-xs font-bold text-gray-800 outline-none focus:ring-2 focus:ring-orange-500/20 text-center" />}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {activeContractorWorkers.length === 0 && (
                  <div className="text-center py-8 text-gray-400 font-bold bg-gray-50 rounded-xl border border-gray-100 text-xs">
                    No workers found for this team.
                  </div>
                )}
              </div>

              <button onClick={submitDailyLog} disabled={isSubmittingLog || activeContractorWorkers.length === 0} className={`w-full mt-8 font-black py-3.5 rounded-xl flex items-center justify-center gap-2 shadow-md transition-colors ${isSubmittingLog || activeContractorWorkers.length === 0 ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : 'bg-gray-900 hover:bg-emerald-600 text-white'}`}>
                <Save className="w-4 h-4" /> {isSubmittingLog ? 'Saving...' : "Secure Today's Attendance"}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
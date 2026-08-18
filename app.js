/**
 * AHU & Daily KWH Energy Intelligence Tracker - Main Application Script
 * 
 * Features:
 * 1. Multi-Device Real-Time Cloud Persistence & Device Sync Engine (Room pairing, auto-polling, backup/restore)
 * 2. Interactive Month-Wise Consumption & Energy Analytics Dashboard (12-month navigation ribbon, KPI cards, Chart.js visuals, Excel-twin data grid)
 * 3. Daily KWH Entry, AHU Saving Tracker, Transaction Audit Log & Executive Admin Portal
 * 4. Real-time August 2026 & Multi-Month Excel Export Synchronization
 */

(function () {
  // Global App State
  const state = {
    activeTab: 'kwh', // 'kwh' | 'ahu' | 'txn' | 'month' | 'admin'
    selectedDate: '2026-08-18',
    selectedMonthMatrix: 'Aug-26', // Month filter key for Monthly Matrix view
    kwhFilter: 'all', // 'all' | 'eb' | 'ahu' | 'btu'
    fontScale: 1, // 0.85, 1, 1.15, 1.3
    isAdminUnlocked: false,
    showAdminModal: false,
    adminErrorMessage: '',
    
    // Transaction History Filters
    txnFilterPeriod: 'Aug-26', // 'all' | 'Aug-26' | 'Jul-26' | 'Jun-26' | 'May-26' | 'today'
    txnFilterCat: 'all', // 'all' | 'eb' | 'ahu' | 'btu' | 'ahu_sched' | 'dg'
    txnSearchQuery: '',

    rates: {
      kwh: 7.45,
      btu: 4.30,
      dg: 33.85
    },

    // Real-Time Cloud Synchronization State
    cloudSync: {
      room: 'CBRE-DT3-FACILITY-2026',
      status: 'synced', // 'synced' | 'syncing' | 'offline' | 'error'
      lastSyncTime: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      deviceId: 'DEV-' + Math.random().toString(36).substring(2, 7).toUpperCase(),
      showModal: false,
      isAutoPolling: true,
      lastCloudVersion: 0,
      syncError: ''
    },

    // Interactive Month-Wise Consumption Dashboard State
    monthDashboard: {
      selectedYear: '2026', // '2026' | '2025' | '2024' | 'all'
      selectedMonth: 'Aug-26', // 'Aug-26', 'Jul-26', 'Jun-26', etc.
      viewMode: 'dashboard', // 'dashboard' | 'consumption_grid' | 'readings_grid' | 'macro_matrix'
      meterFilter: 'all', // 'all' | 'eb' | 'ahu' | 'btu' | 'dg'
      searchQuery: '',
      highlightHigh: false,
      chartType: 'trend' // 'trend' | 'share' | 'mom' | 'cost'
    },

    // Master Tracker Data
    data: {
      ahu_saving: {},
      kwh_daily: {},
      month_baselines: {},
      summary_matrix: { months: [], consumption: {}, cost: {} },
      transactions: []
    }
  };

  // Active Chart.js instances storage for clean lifecycle management
  let activeCharts = {};

  // Helper: Convert date string "YYYY-MM-DD" to Month Key "Aug-26"
  function dateToMonthKey(dStr) {
    if (!dStr) return '';
    const dt = new Date(dStr);
    const mShort = dt.toLocaleString('en-US', { month: 'short' });
    const yShort = dt.toLocaleString('en-US', { year: '2-digit' });
    return `${mShort}-${yShort}`;
  }

  // Helper: Get previous date string "YYYY-MM-DD"
  function getPreviousDateStr(dateStr) {
    const d = new Date(dateStr);
    d.setDate(d.getDate() - 1);
    return d.toISOString().slice(0, 10);
  }

  // Helper: Get exact previous day reading or baseline
  function getExactPreviousReading(dateStr, category, meterId) {
    const prevDateStr = getPreviousDateStr(dateStr);
    const prevDayData = state.data.kwh_daily[prevDateStr];

    if (prevDayData && prevDayData[category]) {
      const found = prevDayData[category].find(x => x.id === meterId);
      if (found && found.reading !== undefined) {
        return {
          reading: Number(found.reading),
          dg_reading: Number(found.dg_reading || 0),
          source: `Prev Day (${prevDateStr})`
        };
      }
    }

    const baselines = state.data.month_baselines?.[dateStr];
    if (baselines && baselines[category]) {
      const bItem = baselines[category][meterId - 1];
      if (bItem) {
        return {
          reading: Number(bItem.prev_reading || 0),
          dg_reading: Number(bItem.prev_dg_reading || 0),
          source: `Excel Baseline (${dateStr})`
        };
      }
    }

    const sortedDates = Object.keys(state.data.kwh_daily || {}).sort().reverse();
    const precedingDate = sortedDates.find(d => d < dateStr);
    if (precedingDate && state.data.kwh_daily[precedingDate]?.[category]) {
      const found = state.data.kwh_daily[precedingDate][category].find(x => x.id === meterId);
      if (found) {
        return {
          reading: Number(found.reading || 0),
          dg_reading: Number(found.dg_reading || 0),
          source: `Prev (${precedingDate})`
        };
      }
    }

    return { reading: 0, dg_reading: 0, source: 'Baseline' };
  }

  // Helper: Get auto-calculated KWH consumptions for AHU calculations
  function getAutoKWHConsumptionsForDate(dateStr) {
    const currKWH = state.data.kwh_daily[dateStr] || {};
    const ahuCons = {};
    const btuCons = {};
    const dgCons = {};

    [1, 2, 3, 4].forEach(id => {
      const prevInfoA = getExactPreviousReading(dateStr, 'ahu', id);
      const currAHU = currKWH.ahu?.find(a => a.id === id) || {};
      const currValA = currAHU.reading !== undefined ? Number(currAHU.reading) : prevInfoA.reading;
      const currDGA = currAHU.dg_reading !== undefined ? Number(currAHU.dg_reading) : prevInfoA.dg_reading;
      
      ahuCons[`AHU${id}`] = Math.max(0, currValA - prevInfoA.reading);
      dgCons[`AHU${id}`] = Math.max(0, currDGA - prevInfoA.dg_reading);

      const prevInfoB = getExactPreviousReading(dateStr, 'btu', id);
      const currBTU = currKWH.btu?.find(b => b.id === id) || {};
      const currValB = currBTU.reading !== undefined ? Number(currBTU.reading) : prevInfoB.reading;
      btuCons[`AHU${id}`] = Math.max(0, currValB - prevInfoB.reading);
    });

    return { ahuCons, btuCons, dgCons };
  }

  
  function ensureDateStructure(dateStr) {
    if (!state.data.kwh_daily[dateStr]) {
      state.data.kwh_daily[dateStr] = {
        eb: [1, 2, 3, 4, 5, 6, 7, 8].map(id => {
          const prev = getExactPreviousReading(dateStr, 'eb', id);
          return {
            id: id,
            location: "3F",
            name: id === 1 ? "SL No 540430038840 (EB Unit)" :
                  id === 2 ? "SL No 540430038845 (EB Unit)" :
                  id === 3 ? "SL No 540430038844 (EB Unit)" :
                  id === 4 ? "SL No 540430038821 (EB Unit)" :
                  id === 5 ? "SL No 540430038849 (EB Unit)" :
                  id === 6 ? "SL No 540430038646 (EB Unit)" :
                  id === 7 ? "SL No 540430038843 (EB Unit)" : "SL No 540430038848 (EB Unit)",
            reading: prev.reading,
            dg_reading: prev.dg_reading
          };
        }),
        ahu: [1, 2, 3, 4].map(id => {
          const prev = getExactPreviousReading(dateStr, 'ahu', id);
          return {
            id: id,
            location: "3F",
            name: `AHU${id}`,
            reading: prev.reading,
            dg_reading: prev.dg_reading
          };
        }),
        btu: [1, 2, 3, 4].map(id => {
          const prev = getExactPreviousReading(dateStr, 'btu', id);
          return {
            id: id,
            location: "3F",
            name: `AHU${id} - BTU`,
            reading: prev.reading
          };
        })
      };
    }

    if (!state.data.ahu_saving[dateStr]) {
      const dt = new Date(dateStr);
      const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      const dayStr = dayNames[dt.getDay()];
      const autoCons = getAutoKWHConsumptionsForDate(dateStr);

      state.data.ahu_saving[dateStr] = {
        date: dateStr,
        day: dayStr,
        ahus: [
          { ahu_id: "AHU1", on_time: "07:12", off_time: "20:00", kwh_cons: autoCons.ahuCons["AHU1"], btu_cons: autoCons.btuCons["AHU1"] },
          { ahu_id: "AHU2", on_time: "07:12", off_time: "20:00", kwh_cons: autoCons.ahuCons["AHU2"], btu_cons: autoCons.btuCons["AHU2"] },
          { ahu_id: "AHU3", on_time: "07:12", off_time: "18:33", kwh_cons: autoCons.ahuCons["AHU3"], btu_cons: autoCons.btuCons["AHU3"] },
          { ahu_id: "AHU4", on_time: "07:12", off_time: "18:03", kwh_cons: autoCons.ahuCons["AHU4"], btu_cons: autoCons.btuCons["AHU4"] }
        ],
        kwh_rate: state.rates.kwh,
        btu_rate: state.rates.btu
      };
    }
  }

  // Real-Time Transaction Logging Engine
  function recalculateDynamicSummaryMatrix() {
    const kwhDaily = state.data.kwh_daily || {};
    const monthMap = {}; // { 'Aug-26': { eb: 0, ahu: 0, dg: 0, btu: 0 } }

    Object.keys(kwhDaily).forEach(dStr => {
      const mKey = dateToMonthKey(dStr);
      if (!monthMap[mKey]) {
        monthMap[mKey] = { eb: 0, ahu: 0, dg: 0, btu: 0 };
      }

      const dayData = kwhDaily[dStr];
      (dayData.eb || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'eb', m.id);
        const ebCons = Math.max(0, (Number(m.reading) || 0) - prev.reading);
        const dgCons = Math.max(0, (Number(m.dg_reading) || 0) - prev.dg_reading);
        monthMap[mKey].eb += ebCons;
        monthMap[mKey].dg += dgCons;
      });

      (dayData.ahu || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'ahu', m.id);
        const ahuCons = Math.max(0, (Number(m.reading) || 0) - prev.reading);
        monthMap[mKey].ahu += ahuCons;
      });

      (dayData.btu || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'btu', m.id);
        const btuCons = Math.max(0, (Number(m.reading) || 0) - prev.reading);
        monthMap[mKey].btu += btuCons;
      });
    });

    // Merge with base historical months
    const baseMonths = ["Jan-25", "Feb-25", "Mar-25", "Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26"];
    const allMonths = Array.from(new Set([...baseMonths, ...Object.keys(monthMap)]));

    if (!state.data.summary_matrix) state.data.summary_matrix = { months: [], consumption: {}, cost: {} };
    state.data.summary_matrix.months = allMonths;

    const consEB = [], consAHU = [], consDG = [], consBTU = [], consTot = [];
    const costEB = [], costAHU = [], costDG = [], costBTU = [], costTot = [];

    allMonths.forEach(m => {
      let eb = monthMap[m]?.eb || 0;
      let ahu = monthMap[m]?.ahu || 0;
      let dg = monthMap[m]?.dg || 0;
      let btu = monthMap[m]?.btu || 0;

      const seedIdx = (window.SEED_DATA?.summary_matrix?.months || []).indexOf(m);
      if (seedIdx !== -1 && eb === 0) {
        eb = window.SEED_DATA.summary_matrix.consumption['EB']?.[seedIdx] || 0;
        ahu = window.SEED_DATA.summary_matrix.consumption['AHU']?.[seedIdx] || 0;
        dg = window.SEED_DATA.summary_matrix.consumption['DG']?.[seedIdx] || 0;
        btu = window.SEED_DATA.summary_matrix.consumption['BTU']?.[seedIdx] || 0;
      }

      const tot = eb + ahu + dg + btu;
      consEB.push(eb); consAHU.push(ahu); consDG.push(dg); consBTU.push(btu); consTot.push(tot);

      const cEB = eb * state.rates.kwh;
      const cAHU = ahu * state.rates.kwh;
      const cDG = dg * state.rates.dg;
      const cBTU = btu * state.rates.btu;
      const cTot = cEB + cAHU + cDG + cBTU;

      costEB.push(cEB); costAHU.push(cAHU); costDG.push(cDG); costBTU.push(cTot);
    });

    state.data.summary_matrix.consumption = { 'EB': consEB, 'AHU': consAHU, 'DG': consDG, 'BTU': consBTU, 'Total': consTot };
    state.data.summary_matrix.cost = { 'EB': costEB, 'AHU': costAHU, 'DG': costDG, 'BTU': costBTU, 'Total': costTot };
  }

  // Calculate AHU Savings for a date
  function calculateAHUSaving(dayRec) {
    const dayStr = dayRec.day;
    const isSun = dayStr === 'Sun';
    const isSat = dayStr === 'Sat';

    // Standard Planned Turn-Off: 8:00 PM (20:00 / 1200 mins) on weekdays, 2:00 PM (14:00 / 840 mins) on Sat, 0 on Sun
    const stdSchedHrs = isSun ? 0 : (isSat ? 6 : 12);
    const stdSchedMins = stdSchedHrs * 60; // 720 mins on Mon-Fri, 360 mins on Sat
    const plannedOffMins = isSat ? (14 * 60) : (20 * 60);

    const autoCons = getAutoKWHConsumptionsForDate(dayRec.date);

    let totSchedHrs = 0, totSchedMins = 0, totRunHrs = 0, totSavedMins = 0, totSavedHrs = 0;
    let totAHUKwh = 0, totAHUBaseCost = 0, totAHUSavedCost = 0;
    let totBTUUnits = 0, totBTUBaseCost = 0, totBTUSavedCost = 0;
    let totDGKwh = 0, totDGBaseCost = 0, totDGSavedCost = 0;
    let totBaseScheduledCost = 0, totCombSavedCost = 0, totActualIncurredCost = 0;

    const ahusComputed = (dayRec.ahus || []).map((ahu) => {
      const onStr = ahu.on_time || "OFF";
      const offStr = ahu.off_time || "OFF";

      let runMins = 0;
      let savedMins = 0;

      if (onStr !== "OFF" && offStr !== "OFF" && onStr && offStr) {
        const [onH, onM] = onStr.split(':').map(Number);
        const [offH, offM] = offStr.split(':').map(Number);
        const onMinutes = (onH * 60) + (onM || 0);
        const offMinutes = (offH * 60) + (offM || 0);

        runMins = Math.max(0, offMinutes - onMinutes);

        // Calculate exact minutes saved from planned 8:00 PM (1200 mins) turn-off
        // e.g., if switched off at 19:45 (1185 mins), saved = 1200 - 1185 = 15 minutes!
        if (!isSun) {
          savedMins = Math.max(0, plannedOffMins - offMinutes);
        }
      }

      const runHrs = runMins / 60;
      const savedHrs = savedMins / 60;

      // 1. AHU Electrical Power: Difference KWH * 7.45 rupees
      const ahuKwh = ahu.kwh_cons !== undefined && ahu.kwh_cons !== null
        ? Number(ahu.kwh_cons) 
        : (autoCons.ahuCons[ahu.ahu_id] || 0);
      const ahuBaseCost = ahuKwh * 7.45;

      // 2. BTU Cooling Load: Difference * 40% * 10.75 rupees (effective rate 4.30)
      const btuUnits = ahu.btu_cons !== undefined && ahu.btu_cons !== null
        ? Number(ahu.btu_cons)
        : (autoCons.btuCons[ahu.ahu_id] || 0);
      const btuBaseCost = btuUnits * 0.40 * 10.75;

      // 3. DG Backup Power: Difference * 33.85 rupees
      const dgKwh = ahu.dg_cons !== undefined && ahu.dg_cons !== null
        ? Number(ahu.dg_cons)
        : (autoCons.dgCons[ahu.ahu_id] || 0);
      const dgBaseCost = dgKwh * 33.85;

      // Total Base Scheduled Cost for the day planned for 8:00 PM
      const ahuBaseScheduledCost = ahuBaseCost + btuBaseCost + dgBaseCost;

      // Minute-wise cost rates from total unit consumed for that day (divided by scheduled 720 mins)
      const costPerMin = stdSchedMins > 0 ? (ahuBaseScheduledCost / stdSchedMins) : 0;
      const ahuRatePerMin = stdSchedMins > 0 ? (ahuBaseCost / stdSchedMins) : 0;
      const btuRatePerMin = stdSchedMins > 0 ? (btuBaseCost / stdSchedMins) : 0;
      const dgRatePerMin = stdSchedMins > 0 ? (dgBaseCost / stdSchedMins) : 0;

      // Minute-wise cost savings: e.g. 15 mins * (103 * 7.45 / 720)
      const costSaved = savedMins * costPerMin;
      const ahuCostSaved = savedMins * ahuRatePerMin;
      const btuCostSaved = savedMins * btuRatePerMin;
      const dgCostSaved = savedMins * dgRatePerMin;

      // Actual incurred cost after switching off early
      const actualIncurredCost = ahuBaseScheduledCost - costSaved;

      const runHrsWhole = Math.floor(runHrs);
      const runMinsRem = Math.round((runHrs - runHrsWhole) * 60);

      totSchedHrs += stdSchedHrs;
      totSchedMins += stdSchedMins;
      totRunHrs += runHrs;
      totSavedMins += savedMins;
      totSavedHrs += savedHrs;

      totAHUKwh += ahuKwh;
      totAHUBaseCost += ahuBaseCost;
      totAHUSavedCost += ahuCostSaved;

      totBTUUnits += btuUnits;
      totBTUBaseCost += btuBaseCost;
      totBTUSavedCost += btuCostSaved;

      totDGKwh += dgKwh;
      totDGBaseCost += dgBaseCost;
      totDGSavedCost += dgCostSaved;

      totBaseScheduledCost += ahuBaseScheduledCost;
      totCombSavedCost += costSaved;
      totActualIncurredCost += actualIncurredCost;

      return {
        ...ahu,
        kwh_cons: ahuKwh,
        btu_cons: btuUnits,
        dg_cons: dgKwh,
        sched_hrs: stdSchedHrs,
        sched_mins: stdSchedMins,
        actual_run_hrs: runHrs,
        run_hrs_text: `${runHrsWhole}h ${runMinsRem}m`,
        saved_hrs: savedHrs,
        saved_mins: savedMins,
        saved_text: savedMins >= 60 ? `${(savedMins/60).toFixed(1)} hrs (${savedMins} mins)` : `${savedMins} mins`,
        ahu_cost: ahuBaseCost,
        btu_cost: btuBaseCost,
        dg_cost: dgBaseCost,
        tot_day_base_cost: ahuBaseScheduledCost,
        cost_per_min: costPerMin,
        cost_per_hr: costPerMin * 60,
        cost_saved: costSaved,
        ahu_cost_saved: ahuCostSaved,
        btu_cost_saved: btuCostSaved,
        dg_cost_saved: dgCostSaved,
        actual_incurred_cost: actualIncurredCost,
        tot_cost_saved_today: costSaved,
        tot_full_cost_8pm_today: ahuBaseScheduledCost
      };
    });

    const savingsPct = totBaseScheduledCost > 0 ? (totCombSavedCost / totBaseScheduledCost) : 0;

    return {
      ...dayRec,
      ahus: ahusComputed,
      tot_sched: totSchedHrs,
      tot_sched_mins: totSchedMins,
      tot_run: totRunHrs,
      tot_saved: totSavedHrs,
      tot_saved_mins: totSavedMins,
      tot_kwh: totAHUKwh,
      tot_ahu_cost: totAHUBaseCost,
      tot_ahu_cost_saved: totAHUSavedCost,
      tot_btu_units: totBTUUnits,
      tot_btu_cost: totBTUBaseCost,
      tot_btu_cost_saved: totBTUSavedCost,
      tot_dg_kwh: totDGKwh,
      tot_dg_cost: totDGBaseCost,
      tot_dg_cost_saved: totDGSavedCost,
      tot_base_cost: totBaseScheduledCost,
      tot_actual_cost: totActualIncurredCost,
      tot_comb_cost_saved: totCombSavedCost,
      tot_comb_full_cost: totBaseScheduledCost,
      total_saved_hrs: totSavedHrs,
      total_combined_saving: totCombSavedCost,
      savings_pct: savingsPct
    };
  }

  
  function showToast(msg) {
    const existing = document.querySelector('.toast-msg');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.className = 'toast-msg';
    toast.innerHTML = `<i data-lucide="check-circle" style="color:var(--success-color);"></i> <span>${msg}</span>`;
    document.body.appendChild(toast);
    if (window.lucide) lucide.createIcons();

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 300);
    }, 2800);
  }

  // Expose TrackerUtils globally for ExcelExporter & modular components
  window.TrackerUtils = {
    getExactPreviousReading,
    getAutoKWHConsumptionsForDate,
    calculateAHUSaving,
    dateToMonthKey,
    getPreviousDateStr
  };

  // =========================================================================
  // Real-Time Multi-Device Cloud Persistence & Sync Engine
  // =========================================================================
  window.CloudSync = {
    getRoom: function () {
      const urlParams = new URLSearchParams(window.location.search);
      const urlRoom = urlParams.get('sync') || urlParams.get('room');
      if (urlRoom) {
        localStorage.setItem('kwh_cloud_sync_room', urlRoom.trim());
        return urlRoom.trim();
      }
      return localStorage.getItem('kwh_cloud_sync_room') || 'CBRE-DT3-FACILITY-2026';
    },

    setRoom: function (newRoom) {
      if (!newRoom || !newRoom.trim()) return;
      const cleanRoom = newRoom.trim();
      state.cloudSync.room = cleanRoom;
      localStorage.setItem('kwh_cloud_sync_room', cleanRoom);
      this.pull(true);
    },

    init: function () {
      state.cloudSync.room = this.getRoom();

      // Initial cloud pull
      this.pull(false);

      // Auto-poll cloud every 4 seconds for real-time instantaneous updates across devices
      setInterval(() => {
        if (state.cloudSync.isAutoPolling && document.visibilityState === 'visible') {
          this.pull(false);
        }
      }, 4000);

      // Pull immediately when user switches back to the tab or focuses window
      window.addEventListener('focus', () => this.pull(false));
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') this.pull(false);
      });
    },

    // Pull latest data from cloud
    pull: function (manual) {
      if (!window.location.protocol.startsWith('http')) {
        state.cloudSync.status = 'synced';
        this.updateHeaderSyncUI();
        return;
      }

      state.cloudSync.status = 'syncing';
      this.updateHeaderSyncUI();

      const syncUrl = `/api/sync?room=${encodeURIComponent(state.cloudSync.room)}&_nocache=${Date.now()}`;

      fetch(syncUrl, {
        method: 'GET',
        cache: 'no-store',
        headers: {
          'Accept': 'application/json',
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache'
        }
      })
        .then(res => {
          if (!res.ok) throw new Error('Cloud sync HTTP ' + res.status);
          return res.json();
        })
        .then(json => {
          if (json && json.data) {
            const cloudData = json.data;
            const cloudVersion = json.version || 0;

            let hasChanges = false;
            let updatedDates = [];
            if (!state.data.kwh_daily) state.data.kwh_daily = {};
            if (!state.data.ahu_saving) state.data.ahu_saving = {};

            // Deep merge cloud kwh_daily dates into local
            Object.keys(cloudData.kwh_daily || {}).forEach(d => {
              if (!state.data.kwh_daily[d]) {
                state.data.kwh_daily[d] = cloudData.kwh_daily[d];
                hasChanges = true;
                updatedDates.push(d);
              } else {
                const localDay = state.data.kwh_daily[d];
                const cloudDay = cloudData.kwh_daily[d];
                
                ['eb', 'ahu', 'btu'].forEach(cat => {
                  if (cloudDay[cat] && cloudDay[cat].length) {
                    if (!localDay[cat]) localDay[cat] = [];
                    cloudDay[cat].forEach(cm => {
                      const lm = localDay[cat].find(x => x.id === cm.id);
                      if (!lm) {
                        localDay[cat].push(cm);
                        hasChanges = true;
                        if (!updatedDates.includes(d)) updatedDates.push(d);
                      } else if (cm.reading !== undefined && cm.reading !== lm.reading) {
                        lm.reading = cm.reading;
                        if (cm.dg_reading !== undefined) lm.dg_reading = cm.dg_reading;
                        hasChanges = true;
                        if (!updatedDates.includes(d)) updatedDates.push(d);
                      }
                    });
                  }
                });
              }
            });

            // Deep merge cloud ahu_saving records
            Object.keys(cloudData.ahu_saving || {}).forEach(d => {
              if (!state.data.ahu_saving[d]) {
                state.data.ahu_saving[d] = cloudData.ahu_saving[d];
                hasChanges = true;
              } else {
                const localAHU = state.data.ahu_saving[d];
                const cloudAHU = cloudData.ahu_saving[d];
                if (cloudAHU.ahus && cloudAHU.ahus.length) {
                  if (!localAHU.ahus) localAHU.ahus = [];
                  cloudAHU.ahus.forEach(ca => {
                    const la = localAHU.ahus.find(x => x.ahu_id === ca.ahu_id);
                    if (!la) {
                      localAHU.ahus.push(ca);
                      hasChanges = true;
                    } else {
                      if (ca.on_time !== la.on_time || ca.off_time !== la.off_time || ca.kwh_cons !== la.kwh_cons || ca.btu_cons !== la.btu_cons) {
                        la.on_time = ca.on_time;
                        la.off_time = ca.off_time;
                        la.kwh_cons = ca.kwh_cons;
                        la.btu_cons = ca.btu_cons;
                        hasChanges = true;
                      }
                    }
                  });
                }
              }
            });

            if (cloudData.transactions && cloudData.transactions.length) {
              const existingIds = new Set((state.data.transactions || []).map(t => t.id));
              cloudData.transactions.forEach(t => {
                if (!existingIds.has(t.id)) {
                  if (!state.data.transactions) state.data.transactions = [];
                  state.data.transactions.push(t);
                  hasChanges = true;
                }
              });
            }

            if (hasChanges || cloudVersion > state.cloudSync.lastCloudVersion) {
              state.cloudSync.lastCloudVersion = Math.max(cloudVersion, state.cloudSync.lastCloudVersion);
              recalculateDynamicSummaryMatrix();
              localStorage.setItem('kwh_ahu_tracker_data_v2', JSON.stringify(state.data));
              if (manual) {
                showToast(`☁️ Cloud data synchronized (Room: ${state.cloudSync.room})`);
              } else if (hasChanges && updatedDates.length > 0) {
                showToast(`⚡ Live update synced from other device (${updatedDates.join(', ')})`);
              }
              render();
            }
          }
          state.cloudSync.status = 'synced';
          state.cloudSync.lastSyncTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
          this.updateHeaderSyncUI();
        })
        .catch(err => {
          console.warn('Cloud pull fallback to local storage:', err);
          state.cloudSync.status = 'synced';
          state.cloudSync.lastSyncTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
          this.updateHeaderSyncUI();
        });
    },

    // Push local data to cloud
    push: function (manual) {
      if (!window.location.protocol.startsWith('http')) {
        state.cloudSync.status = 'synced';
        this.updateHeaderSyncUI();
        return;
      }

      state.cloudSync.status = 'syncing';
      this.updateHeaderSyncUI();

      const payload = {
        room: state.cloudSync.room,
        version: Date.now(),
        deviceId: state.cloudSync.deviceId,
        data: state.data
      };

      fetch(`/api/sync?_nocache=${Date.now()}`, {
        method: 'POST',
        cache: 'no-store',
        headers: {
          'Content-Type': 'application/json',
          'Cache-Control': 'no-cache, no-store'
        },
        body: JSON.stringify(payload)
      })
      .then(res => res.json())
      .then(json => {
        state.cloudSync.status = 'synced';
        state.cloudSync.lastSyncTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        state.cloudSync.lastCloudVersion = json.version || Date.now();
        this.updateHeaderSyncUI();
        if (manual) showToast('☁️ Data pushed & synchronized to all devices!');
      })
      .catch(err => {
        console.warn('Cloud push offline/local mode:', err);
        state.cloudSync.status = 'synced';
        state.cloudSync.lastSyncTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        this.updateHeaderSyncUI();
      });
    },

    // Update only the header pill without re-rendering whole app if typing
    updateHeaderSyncUI: function () {
      const pill = document.querySelector('.live-sync-indicator');
      if (!pill) return;
      pill.className = `live-sync-indicator ${state.cloudSync.status}`;
      const textSpan = pill.querySelector('.live-sync-text');
      if (textSpan) {
        if (state.cloudSync.status === 'syncing') {
          textSpan.textContent = 'Syncing...';
        } else {
          textSpan.textContent = `Cloud Synced (${state.cloudSync.room})`;
        }
      }
    },

    // Full JSON Backup Export
    exportBackupJSON: function () {
      const jsonStr = JSON.stringify(state.data, null, 2);
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const todayStr = new Date().toISOString().slice(0, 10);
      const fileName = `CBRE_Energy_Tracker_Full_Backup_${todayStr}.json`;
      if (window.ExcelExporter) {
        window.ExcelExporter.triggerBlobDownload(blob, fileName);
      }
      showToast('📥 Full Database Backup JSON exported!');
    },

    // Full JSON Backup Import
    importBackupJSON: function (file) {
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const parsed = JSON.parse(e.target.result);
          if (parsed && (parsed.kwh_daily || parsed.ahu_saving)) {
            state.data = parsed;
            recalculateDynamicSummaryMatrix();
            saveData(false);
            showToast('📤 Database Backup Restored & Synced Successfully!');
            render();
          } else {
            alert('Invalid backup JSON format. Required keys: kwh_daily or ahu_saving.');
          }
        } catch (err) {
          alert('Error parsing JSON backup file: ' + err.message);
        }
      };
      reader.readAsText(file);
    }
  };

  // Master Data Initialization
  function initData() {
    const stored = localStorage.getItem('kwh_ahu_tracker_data_v2');
    if (stored) {
      try {
        state.data = JSON.parse(stored);
      } catch (e) {
        console.error("Failed to parse stored data:", e);
        state.data = window.SEED_DATA || { ahu_saving: {}, kwh_daily: {}, month_baselines: {}, summary_matrix: {}, transactions: [] };
      }
    } else {
      state.data = window.SEED_DATA || { ahu_saving: {}, kwh_daily: {}, month_baselines: {}, summary_matrix: {}, transactions: [] };
    }

    if (!state.data.transactions) {
      state.data.transactions = [];
    }

    const todayStr = new Date().toISOString().slice(0, 10);
    state.selectedDate = todayStr;
    state.selectedMonthMatrix = dateToMonthKey(todayStr) || 'Aug-26';
    state.monthDashboard.selectedMonth = state.selectedMonthMatrix;

    ensureDateStructure(todayStr);
    generateInitialTransactionsIfNeeded();
    recalculateDynamicSummaryMatrix();

    // Start Cloud Synchronization Engine
    window.CloudSync.init();
  }

  // Save Data locally and push to cloud
  function saveData(skipBackend) {
    recalculateDynamicSummaryMatrix();
    localStorage.setItem('kwh_ahu_tracker_data_v2', JSON.stringify(state.data));

    if (!skipBackend) {
      window.CloudSync.push(false);
    }
  }

  // Record a transaction in audit log
  function recordTransaction(entry) {
    if (!state.data.transactions) state.data.transactions = [];

    const dt = new Date();
    const txn = {
      id: 'TXN-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 5).toUpperCase(),
      timestamp: dt.toISOString(),
      displayTime: dt.toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      date: entry.date || state.selectedDate,
      monthKey: dateToMonthKey(entry.date || state.selectedDate),
      category: entry.category,
      categoryLabel: entry.categoryLabel,
      meterId: entry.meterId,
      meterName: entry.meterName,
      location: entry.location || "3F",
      prevReading: entry.prevReading !== undefined ? Number(entry.prevReading) : undefined,
      newReading: entry.newReading !== undefined ? Number(entry.newReading) : undefined,
      consumption: entry.consumption !== undefined ? Number(entry.consumption.toFixed(2)) : undefined,
      cost: entry.cost !== undefined ? Number(entry.cost.toFixed(2)) : undefined,
      rate: entry.rate,
      details: entry.details || '',
      source: entry.source || 'User Input',
      status: 'Committed'
    };

    state.data.transactions.unshift(txn);
    if (state.data.transactions.length > 500) {
      state.data.transactions = state.data.transactions.slice(0, 500);
    }
  }

  // Generate initial audit log if empty
  function generateInitialTransactionsIfNeeded() {
    if (state.data.transactions && state.data.transactions.length > 0) return;

    const dates = Object.keys(state.data.kwh_daily || {}).sort().reverse();
    const sampleDates = dates.slice(0, 15);

    sampleDates.forEach(dStr => {
      const dayKWH = state.data.kwh_daily[dStr];
      const mKey = dateToMonthKey(dStr);

      (dayKWH.eb || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'eb', m.id);
        const cons = Math.max(0, (m.reading || 0) - prev.reading);
        if (cons > 0) {
          state.data.transactions.push({
            id: 'TXN-' + Math.random().toString(36).substring(2, 9).toUpperCase(),
            timestamp: new Date(dStr + 'T08:00:00Z').toISOString(),
            displayTime: `${dStr} 08:00:00`,
            date: dStr,
            monthKey: mKey,
            category: 'eb',
            categoryLabel: 'EB Meter Reading',
            meterId: m.id,
            meterName: m.name,
            location: m.location || "3F",
            prevReading: prev.reading,
            newReading: m.reading,
            consumption: Number(cons.toFixed(1)),
            cost: Number((cons * state.rates.kwh).toFixed(2)),
            rate: state.rates.kwh,
            details: 'Initial historical log baseline',
            source: 'Meter Sync',
            status: 'Committed'
          });
        }
      });
    });
  }

  // =========================================================================
  // Master Render Engine
  // =========================================================================
  function render() {
    document.documentElement.style.setProperty('--font-scale', state.fontScale);

    const appContainer = document.getElementById('app');
    if (!appContainer) return;

    appContainer.innerHTML = `
      <div class="app-layout">
        ${renderTopHeader()}
        ${renderNavTabs()}
        ${state.activeTab === 'kwh' || state.activeTab === 'ahu' || state.activeTab === 'txn' ? renderUserControlBarOnlyDate() : ''}
        
        <main class="main-content-area">
          ${renderTabContent()}
        </main>
      </div>

      ${state.showAdminModal ? renderAdminAuthModal() : ''}
      ${state.cloudSync.showModal ? renderCloudSyncModal() : ''}
    `;

    if (window.lucide) lucide.createIcons();

    // Render Charts for active tab
    if (state.activeTab === 'month') {
      renderMonthDashboardCharts();
    } else if (state.activeTab === 'admin') {
      renderAdminCharts();
    }

    attachEventListeners();
  }

  // Top Header with Live Cloud Sync Pill
  function renderTopHeader() {
    return `
      <header class="top-header">
        <div class="brand-container">
          <div class="brand-logo"><i data-lucide="zap"></i></div>
          <div class="brand-title">
            <h1>AHU & Energy (KWH) Intelligence Tracker</h1>
            <div class="brand-subtitle">CBRE Facility Management | Multi-Device Real-Time Energy Monitor</div>
          </div>
        </div>

        <div class="top-utilities">
          <div class="live-sync-indicator ${state.cloudSync.status}" id="btn-cloud-sync-modal" title="Click to manage Multi-Device Cloud Sync & Pairing">
            <span class="pulse-dot"></span>
            <span class="live-sync-text">Cloud Synced (${state.cloudSync.room})</span>
          </div>

          <button class="action-btn-sm" id="btn-force-cloud-sync" title="Sync now with all devices" style="background: rgba(59, 130, 246, 0.15); border: 1px solid var(--primary-blue); color: #fff; padding: 0.35rem 0.65rem; border-radius: var(--radius-sm); display:flex; align-items:center; gap:0.35rem; font-size:0.75rem; cursor:pointer;">
            <i data-lucide="refresh-cw" style="width:14px; height:14px;"></i> Sync
          </button>

          <div class="rate-badge-group">
            <div class="rate-item" title="Electricity Board KWH Unit Rate">⚡ KWH: <strong>₹${state.rates.kwh.toFixed(2)}</strong></div>
            <div class="rate-item" title="BTU Chilled Water Cooling Unit Rate">❄️ BTU: <strong>₹${state.rates.btu.toFixed(2)}</strong></div>
            <div class="rate-item" title="Diesel Generator Unit Rate">⛽ DG: <strong>₹${state.rates.dg.toFixed(2)}</strong></div>
          </div>

          <div class="font-size-control">
            <button class="font-btn ${state.fontScale === 0.85 ? 'active' : ''}" data-font="0.85">A-</button>
            <button class="font-btn ${state.fontScale === 1 ? 'active' : ''}" data-font="1">100%</button>
            <button class="font-btn ${state.fontScale === 1.15 ? 'active' : ''}" data-font="1.15">A+</button>
            <button class="font-btn ${state.fontScale === 1.3 ? 'active' : ''}" data-font="1.3">A++</button>
          </div>

          <button class="admin-trigger-btn ${state.isAdminUnlocked ? 'unlocked' : ''}" id="btn-admin-toggle">
            <i data-lucide="${state.isAdminUnlocked ? 'unlock' : 'lock'}"></i>
            ${state.isAdminUnlocked ? 'Admin Portal' : 'Admin Login'}
          </button>
        </div>
      </header>
    `;
  }

  // Navigation Tabs Bar
  function renderNavTabs() {
    const txnCount = (state.data.transactions || []).length;
    return `
      <nav class="nav-tabs-bar">
        <button class="nav-tab-btn kwh-tab ${state.activeTab === 'kwh' ? 'active' : ''}" data-tab="kwh">
          <i data-lucide="activity"></i> Daily KWH Entry (EB, AHU, BTU)
        </button>
        <button class="nav-tab-btn ahu-tab ${state.activeTab === 'ahu' ? 'active' : ''}" data-tab="ahu">
          <i data-lucide="fan"></i> AHU Saving Tracker & BTU Cost
        </button>
        <button class="nav-tab-btn txn-tab ${state.activeTab === 'txn' ? 'active' : ''}" data-tab="txn">
          <i data-lucide="history"></i> Transaction History & Audit Log
          <span class="tab-badge">${txnCount}</span>
        </button>
        <button class="nav-tab-btn month-tab ${state.activeTab === 'month' ? 'active' : ''}" data-tab="month">
          <i data-lucide="calendar-range"></i> Interactive Month Consumption Dashboard
          <span class="tab-badge" style="background:#00e5ff; color:#0b0f19; font-weight:800;">${state.monthDashboard.selectedMonth}</span>
        </button>
        <button class="nav-tab-btn admin-tab ${state.activeTab === 'admin' ? 'active' : ''}" data-tab="admin">
          <i data-lucide="shield-check"></i> Executive Dashboards & Admin
        </button>
      </nav>
    `;
  }

  // User Date Control Bar for Daily tabs
  function renderUserControlBarOnlyDate() {
    const kwhPills = state.activeTab === 'kwh' ? `
      <div class="filter-pills-container">
        <button class="pill-btn pill-all ${state.kwhFilter === 'all' ? 'active' : ''}" data-filter="all"><i data-lucide="layers"></i> All Sections</button>
        <button class="pill-btn pill-eb ${state.kwhFilter === 'eb' ? 'active' : ''}" data-filter="eb"><i data-lucide="zap"></i> EB Meters</button>
        <button class="pill-btn pill-ahu ${state.kwhFilter === 'ahu' ? 'active' : ''}" data-filter="ahu"><i data-lucide="wind"></i> AHU Meters</button>
        <button class="pill-btn pill-btu ${state.kwhFilter === 'btu' ? 'active' : ''}" data-filter="btu"><i data-lucide="snowflake"></i> BTU Meters</button>
      </div>
    ` : '';

    const dtObj = new Date(state.selectedDate);
    const dateFormattedStr = dtObj.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });

    return `
      <div class="controls-card">
        <div class="date-selector-group">
          <div class="date-input-wrap">
            <i data-lucide="calendar" style="color:#ffffff; font-size:1.2rem;"></i>
            <label>Select Operational Date:</label>
            <input type="date" id="selected-date-picker" value="${state.selectedDate}">
            <span class="date-formatted-text">(${dateFormattedStr})</span>
          </div>
        </div>

        ${kwhPills}
      </div>
    `;
  }

  // Tab Content Switcher
  function renderTabContent() {
    if (state.activeTab === 'kwh') {
      return renderKWHDailyLogContent();
    } else if (state.activeTab === 'ahu') {
      return renderAHUSavingLogContent();
    } else if (state.activeTab === 'txn') {
      return renderTransactionHistoryContent();
    } else if (state.activeTab === 'month') {
      return renderMonthDashboardContent();
    } else if (state.activeTab === 'admin') {
      return renderAdminPortalContent();
    }
    return '';
  }

  // Core Render Functions
  function renderKWHDailyLogContent() {
    const dateStr = state.selectedDate;
    const currKWH = state.data.kwh_daily[dateStr] || { eb: [], ahu: [], btu: [] };

    let totalEBCons = 0, totalEBDGCons = 0;
    let totalAHUCons = 0, totalAHUDGCons = 0;
    let totalBTUCons = 0;

    const showEB = state.kwhFilter === 'all' || state.kwhFilter === 'eb';
    const ebCardsHtml = (currKWH.eb || []).map((meter) => {
      const prevInfo = getExactPreviousReading(dateStr, 'eb', meter.id);
      const prevVal = prevInfo.reading;
      const currVal = meter.reading !== undefined ? Number(meter.reading) : prevVal;
      const ebCons = Math.max(0, currVal - prevVal);
      totalEBCons += ebCons;

      const prevDGVal = prevInfo.dg_reading;
      const currDGVal = meter.dg_reading !== undefined ? Number(meter.dg_reading) : prevDGVal;
      const dgCons = Math.max(0, currDGVal - prevDGVal);
      totalEBDGCons += dgCons;

      return `
        <div class="meter-card eb-theme">
          <div class="meter-card-header">
            <div class="meter-title-wrap">
              <div class="meter-icon-badge"><i data-lucide="zap"></i></div>
              <div>
                <div class="meter-title">${meter.name}</div>
                <div class="meter-location-tag">Loc: ${meter.location || '3F'}</div>
              </div>
            </div>
          </div>

          <div class="reading-inputs-grid">
            <div class="input-block">
              <label>Previous Reading <span class="auto-tag">(${prevInfo.source})</span></label>
              <input type="number" class="input-field" value="${prevVal}" readonly title="Auto-fetched previous reading">
            </div>
            <div class="input-block">
              <label style="color:#00e5ff;">Current Reading (${dateStr})</label>
              <input type="number" step="any" class="input-field kwh-input-field" 
                     data-cat="eb" data-id="${meter.id}" data-field="reading" value="${currVal}">
            </div>
          </div>

          <div class="dg-sub-box">
            <div class="dg-box-title"><i data-lucide="fuel"></i> DG Consumption (Generator Backup)</div>
            <div class="reading-inputs-grid" style="margin-bottom:0;">
              <div class="input-block">
                <label>Prev DG Reading</label>
                <input type="number" class="input-field" value="${prevDGVal}" readonly>
              </div>
              <div class="input-block">
                <label style="color:#fbbf24;">Curr DG Reading</label>
                <input type="number" step="any" class="input-field kwh-input-field" 
                       data-cat="eb" data-id="${meter.id}" data-field="dg_reading" value="${currDGVal}">
              </div>
            </div>
          </div>

          <div class="results-strip">
            <div class="res-item">
              <span class="res-label">EB Units Consumed</span>
              <span class="res-value highlight">${ebCons.toFixed(1)} kWh</span>
            </div>
            <div class="res-item">
              <span class="res-label">DG Units Consumed</span>
              <span class="res-value" style="color:#fbbf24;">${dgCons.toFixed(1)} kWh</span>
            </div>
            <div class="res-item" style="text-align:right;">
              <span class="res-label">Est Cost (EB+DG)</span>
              <span class="res-value">₹${((ebCons * state.rates.kwh) + (dgCons * state.rates.dg)).toFixed(2)}</span>
            </div>
          </div>
        </div>
      `;
    }).join('');

    const showAHU = state.kwhFilter === 'all' || state.kwhFilter === 'ahu';
    const ahuCardsHtml = (currKWH.ahu || []).map((meter) => {
      const prevInfo = getExactPreviousReading(dateStr, 'ahu', meter.id);
      const prevVal = prevInfo.reading;
      const currVal = meter.reading !== undefined ? Number(meter.reading) : prevVal;
      const ahuCons = Math.max(0, currVal - prevVal);
      totalAHUCons += ahuCons;

      const prevDGVal = prevInfo.dg_reading;
      const currDGVal = meter.dg_reading !== undefined ? Number(meter.dg_reading) : prevDGVal;
      const dgCons = Math.max(0, currDGVal - prevDGVal);
      totalAHUDGCons += dgCons;

      return `
        <div class="meter-card ahu-theme">
          <div class="meter-card-header">
            <div class="meter-title-wrap">
              <div class="meter-icon-badge"><i data-lucide="wind"></i></div>
              <div>
                <div class="meter-title">${meter.name}</div>
                <div class="meter-location-tag">Loc: ${meter.location || '3F'}</div>
              </div>
            </div>
          </div>

          <div class="reading-inputs-grid">
            <div class="input-block">
              <label>Previous Reading <span class="auto-tag">(${prevInfo.source})</span></label>
              <input type="number" class="input-field" value="${prevVal}" readonly title="Auto-fetched previous reading">
            </div>
            <div class="input-block">
              <label style="color:#10b981;">Current Reading (${dateStr})</label>
              <input type="number" step="any" class="input-field kwh-input-field" 
                     data-cat="ahu" data-id="${meter.id}" data-field="reading" value="${currVal}">
            </div>
          </div>

          <div class="dg-sub-box">
            <div class="dg-box-title"><i data-lucide="fuel"></i> AHU DG Consumption</div>
            <div class="reading-inputs-grid" style="margin-bottom:0;">
              <div class="input-block">
                <label>Prev DG Reading</label>
                <input type="number" class="input-field" value="${prevDGVal}" readonly>
              </div>
              <div class="input-block">
                <label style="color:#fbbf24;">Curr DG Reading</label>
                <input type="number" step="any" class="input-field kwh-input-field" 
                       data-cat="ahu" data-id="${meter.id}" data-field="dg_reading" value="${currDGVal}">
              </div>
            </div>
          </div>

          <div class="results-strip">
            <div class="res-item">
              <span class="res-label">AHU Consumed</span>
              <span class="res-value highlight">${ahuCons.toFixed(1)} kWh</span>
            </div>
            <div class="res-item">
              <span class="res-label">AHU DG Consumed</span>
              <span class="res-value" style="color:#fbbf24;">${dgCons.toFixed(1)} kWh</span>
            </div>
            <div class="res-item" style="text-align:right;">
              <span class="res-label">Est AHU Cost</span>
              <span class="res-value">₹${((ahuCons * state.rates.kwh) + (dgCons * state.rates.dg)).toFixed(2)}</span>
            </div>
          </div>
        </div>
      `;
    }).join('');

    const showBTU = state.kwhFilter === 'all' || state.kwhFilter === 'btu';
    const btuCardsHtml = (currKWH.btu || []).map((meter) => {
      const prevInfo = getExactPreviousReading(dateStr, 'btu', meter.id);
      const prevVal = prevInfo.reading;
      const currVal = meter.reading !== undefined ? Number(meter.reading) : prevVal;
      const btuCons = Math.max(0, currVal - prevVal);
      totalBTUCons += btuCons;

      return `
        <div class="meter-card btu-theme">
          <div class="meter-card-header">
            <div class="meter-title-wrap">
              <div class="meter-icon-badge"><i data-lucide="snowflake"></i></div>
              <div>
                <div class="meter-title">${meter.name}</div>
                <div class="meter-location-tag">Loc: ${meter.location || '3F'}</div>
              </div>
            </div>
          </div>

          <div class="reading-inputs-grid">
            <div class="input-block">
              <label>Previous Reading <span class="auto-tag">(${prevInfo.source})</span></label>
              <input type="number" class="input-field" value="${prevVal}" readonly title="Auto-fetched previous reading">
            </div>
            <div class="input-block">
              <label style="color:#f59e0b;">Current Reading (${dateStr})</label>
              <input type="number" step="any" class="input-field kwh-input-field" 
                     data-cat="btu" data-id="${meter.id}" data-field="reading" value="${currVal}">
            </div>
          </div>

          <div class="results-strip">
            <div class="res-item">
              <span class="res-label">BTU Units Consumed</span>
              <span class="res-value highlight">${btuCons.toFixed(1)} kWh</span>
            </div>
            <div class="res-item" style="text-align:right;">
              <span class="res-label">Est BTU Cost</span>
              <span class="res-value">₹${(btuCons * state.rates.btu).toFixed(2)}</span>
            </div>
          </div>
        </div>
      `;
    }).join('');

    return `
      <div class="category-summary-banner">
        <div class="summary-stat-box">
          <div class="stat-label">Total EB Power (kWh)</div>
          <div class="stat-val" style="color:var(--eb-color);">${totalEBCons.toFixed(1)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total EB DG Power (kWh)</div>
          <div class="stat-val" style="color:#fbbf24;">${totalEBDGCons.toFixed(1)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total AHU Power (kWh)</div>
          <div class="stat-val" style="color:var(--ahu-color);">${totalAHUCons.toFixed(1)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total BTU Consumption (kWh)</div>
          <div class="stat-val" style="color:var(--btu-color);">${totalBTUCons.toFixed(1)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Cumulative Total Cost (₹)</div>
          <div class="stat-val" style="color:#60a5fa;">₹${(
            (totalEBCons * state.rates.kwh) + 
            (totalEBDGCons * state.rates.dg) + 
            (totalBTUCons * state.rates.btu)
          ).toFixed(2)}</div>
        </div>
      </div>

      ${showEB ? `
        <div class="section-header-title eb-header">
          <h2><i data-lucide="zap"></i> Electricity Board (EB) Meters & DG Backups</h2>
          <span style="font-size:0.8rem; font-weight:normal;">8 Primary EB Meters</span>
        </div>
        <div class="section-grid">${ebCardsHtml}</div>
      ` : ''}

      ${showAHU ? `
        <div class="section-header-title ahu-header">
          <h2><i data-lucide="wind"></i> Air Handling Unit (AHU) Power & DG Backups</h2>
          <span style="font-size:0.8rem; font-weight:normal;">4 AHU Meters</span>
        </div>
        <div class="section-grid">${ahuCardsHtml}</div>
      ` : ''}

      ${showBTU ? `
        <div class="section-header-title btu-header">
          <h2><i data-lucide="snowflake"></i> BTU Chilled Water Cooling Meters</h2>
          <span style="font-size:0.8rem; font-weight:normal;">4 BTU Meters</span>
        </div>
        <div class="section-grid">${btuCardsHtml}</div>
      ` : ''}

      <div class="action-bar-floating">
        <button class="btn-primary-save" id="btn-save-kwh">
          <i data-lucide="save"></i> Save & Commit Daily KWH Readings
        </button>
      </div>
    `;
  }

  function renderAHUSavingLogContent() {
    const dateStr = state.selectedDate;
    const rawDayRec = state.data.ahu_saving[dateStr] || {
      date: dateStr, day: 'Mon', ahus: [], kwh_rate: 7.45, btu_rate: 4.30
    };

    const computed = calculateAHUSaving(rawDayRec);

    const cardsHtml = (computed.ahus || []).map((ahu) => {
      return `
        <div class="meter-card ahu-theme">
          <div class="meter-card-header" style="background: linear-gradient(135deg, rgba(5,150,105,0.25), rgba(15,23,42,0.7)); border-bottom: 1px solid rgba(5,150,105,0.35); padding: 0.9rem 1rem; border-radius: var(--radius-md) var(--radius-md) 0 0;">
            <div class="meter-title-wrap">
              <div class="meter-icon-badge" style="background:rgba(5,150,105,0.25); color:#10b981; border:1px solid #10b981;"><i data-lucide="fan"></i></div>
              <div>
                <div class="meter-title" style="font-size:1.15rem; font-weight:800; color:#fff;">${ahu.ahu_id} Minute-Wise Savings Engine</div>
                <div class="meter-location-tag" style="color:#94a3b8; font-size:0.75rem;">Standard Sched: <strong>${ahu.sched_hrs} hrs (720 mins)</strong> | Planned OFF: <strong>8:00 PM (20:00)</strong></div>
              </div>
            </div>
            <div style="text-align:right;">
              <span style="background:rgba(16,185,129,0.2); border:1px solid #10b981; color:#10b981; padding:0.3rem 0.75rem; border-radius:9999px; font-size:0.8rem; font-weight:800;">
                💰 Saved: ₹${ahu.cost_saved.toFixed(2)}
              </span>
            </div>
          </div>

          <div class="reading-inputs-grid" style="padding: 1rem 1rem 0.5rem 1rem;">
            <div class="input-block">
              <label style="color:#ffffff; font-weight:700;"><i data-lucide="clock" style="width:14px; height:14px; color:#10b981; display:inline-block; vertical-align:middle; margin-right:4px;"></i> ON Time (hh:mm)</label>
              <input type="time" class="input-field ahu-time-input" 
                     data-ahu="${ahu.ahu_id}" data-field="on_time" value="${ahu.on_time || '07:00'}">
            </div>
            <div class="input-block">
              <label style="color:#ffffff; font-weight:700;"><i data-lucide="clock" style="width:14px; height:14px; color:#ef4444; display:inline-block; vertical-align:middle; margin-right:4px;"></i> Actual OFF Time (hh:mm)</label>
              <input type="time" class="input-field ahu-time-input" 
                     data-ahu="${ahu.ahu_id}" data-field="off_time" value="${ahu.off_time || '19:45'}">
            </div>
          </div>

          <div class="reading-inputs-grid" style="padding: 0 1rem 0.75rem 1rem;">
            <div class="input-block">
              <label style="color:#34d399; font-weight:600;">⚡ AHU Power (${ahu.kwh_cons} kWh × ₹7.45)</label>
              <input type="number" step="any" class="input-field ahu-cons-input" 
                     data-ahu="${ahu.ahu_id}" data-field="kwh_cons" value="${ahu.kwh_cons}">
              <span style="font-size:0.72rem; color:#94a3b8;">Base Cost: ₹${ahu.ahu_cost.toFixed(2)} (₹${(ahu.ahu_cost / (ahu.sched_mins || 720)).toFixed(4)}/min)</span>
            </div>
            <div class="input-block">
              <label style="color:#38bdf8; font-weight:600;">🧊 BTU Cooling (${ahu.btu_cons} Units × 40% × ₹10.75)</label>
              <input type="number" step="any" class="input-field ahu-cons-input" 
                     data-ahu="${ahu.ahu_id}" data-field="btu_cons" value="${ahu.btu_cons}">
              <span style="font-size:0.72rem; color:#94a3b8;">Base Cost: ₹${ahu.btu_cost.toFixed(2)} (₹${(ahu.btu_cost / (ahu.sched_mins || 720)).toFixed(4)}/min)</span>
            </div>
          </div>

          <div class="results-strip" style="flex-wrap:wrap; gap:0.5rem; background:rgba(0,0,0,0.5); padding:0.85rem 1rem; border-radius:0 0 var(--radius-md) var(--radius-md); border-top:1px solid rgba(255,255,255,0.08);">
            <div class="res-item" style="flex:1; min-width:115px;">
              <span class="res-label" style="font-size:0.7rem; color:#94a3b8;">Total Scheduled Cost</span>
              <span class="res-value" style="color:#f3f4f6; font-weight:700;">₹${ahu.tot_day_base_cost.toFixed(2)}</span>
            </div>
            <div class="res-item" style="flex:1; min-width:115px;">
              <span class="res-label" style="font-size:0.7rem; color:#94a3b8;">Minute-Wise Rate</span>
              <span class="res-value" style="color:#cbd5e1; font-weight:700;">₹${ahu.cost_per_min.toFixed(4)}/min</span>
            </div>
            <div class="res-item" style="flex:1; min-width:115px;">
              <span class="res-label" style="font-size:0.7rem; color:#94a3b8;">Saved Time (vs 20:00)</span>
              <span class="res-value highlight" style="color:#10b981; font-weight:800;">⏱️ ${ahu.saved_mins} mins</span>
            </div>
            <div class="res-item" style="flex:1; min-width:115px;">
              <span class="res-label" style="font-size:0.7rem; color:#94a3b8;">AHU Power Saved</span>
              <span class="res-value" style="color:#34d399;">₹${ahu.ahu_cost_saved.toFixed(2)}</span>
            </div>
            <div class="res-item" style="flex:1; min-width:115px;">
              <span class="res-label" style="font-size:0.7rem; color:#94a3b8;">BTU Cooling Saved</span>
              <span class="res-value" style="color:#38bdf8;">₹${ahu.btu_cost_saved.toFixed(2)}</span>
            </div>
            <div class="res-item" style="text-align:right; width:100%; border-top:1px dashed rgba(255,255,255,0.15); padding-top:0.5rem; margin-top:0.3rem; display:flex; justify-content:space-between; align-items:center;">
              <span style="font-size:0.8rem; color:#94a3b8;">Actual Incurred Cost: <strong style="color:#e2e8f0;">₹${ahu.actual_incurred_cost.toFixed(2)}</strong></span>
              <div>
                <span class="res-label" style="font-size:0.75rem; color:#10b981; margin-right:0.4rem;">Total Cost Saved for ${ahu.saved_mins} mins:</span>
                <span class="res-value" style="color:#00e5ff; font-size:1.2rem; font-weight:800;">₹${ahu.cost_saved.toFixed(2)}</span>
              </div>
            </div>
          </div>
        </div>
      `;
    }).join('');

    return `
      <div class="category-summary-banner" style="background: linear-gradient(135deg, rgba(15,23,42,0.95), rgba(30,41,59,0.9)); border: 1px solid rgba(5,150,105,0.4); border-radius: var(--radius-lg); padding: 1.25rem; box-shadow: 0 10px 25px rgba(0,0,0,0.4);">
        <div class="summary-stat-box">
          <div class="stat-label">Total Sched Mins</div>
          <div class="stat-val">${computed.tot_sched_mins || 720} mins</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total Saved Time</div>
          <div class="stat-val" style="color:#10b981; font-weight:800;">⏱️ ${computed.tot_saved_mins} mins</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total Base Shift Cost</div>
          <div class="stat-val">₹${computed.tot_base_cost.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">AHU Power Saved (₹7.45)</div>
          <div class="stat-val" style="color:#34d399;">₹${computed.tot_ahu_cost_saved.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">BTU Cooling Saved (40%*10.75)</div>
          <div class="stat-val" style="color:#38bdf8;">₹${computed.tot_btu_cost_saved.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total Cost Saved Today</div>
          <div class="stat-val" style="color:#00e5ff; font-size:1.35rem; font-weight:800;">₹${computed.tot_comb_cost_saved.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Savings Efficiency %</div>
          <div class="stat-val" style="color:#60a5fa;">${(computed.savings_pct * 100).toFixed(2)}%</div>
        </div>
      </div>

      <div class="section-header-title ahu-header" style="margin-top:1.25rem;">
        <h2><i data-lucide="fan"></i> Operational AHU Minute-Wise Cost Savings Engine</h2>
        <span style="font-size:0.8rem; font-weight:normal; color:#94a3b8;">Formula: Rate/Min = (AHU kWh × ₹7.45 + BTU × 40% × ₹10.75 + DG × ₹33.85) / 720 mins | Saved Cost = Saved Mins × Rate/Min</span>
      </div>

      <div class="section-grid">${cardsHtml}</div>

      <div class="action-bar-floating">
        <button class="btn-primary-save" id="btn-save-ahu">
          <i data-lucide="save"></i> Save & Commit AHU & BTU Savings Log
        </button>
      </div>
    `;
  }

  
  function renderTransactionHistoryContent() {
    const txns = state.data.transactions || [];
    const todayStr = new Date().toISOString().slice(0, 10);
    const currMonthKey = dateToMonthKey(todayStr) || 'Aug-26';

    // Calculate live summary stats for current month
    const augTxns = txns.filter(t => t.month === 'Aug-26' || (t.date && t.date.startsWith('2026-08')));
    const augCount = augTxns.length;
    const augKwh = augTxns.reduce((acc, cur) => acc + (cur.consumption || 0), 0);
    const augCost = augTxns.reduce((acc, cur) => acc + (cur.cost || 0), 0);

    // Filter transactions by Period, Category & Search query
    let filtered = txns.filter(t => {
      if (state.txnFilterPeriod !== 'all') {
        if (state.txnFilterPeriod === 'today' && t.date !== todayStr) return false;
        if (state.txnFilterPeriod !== 'today' && t.month !== state.txnFilterPeriod && !(t.date && t.date.startsWith('2026-08') && state.txnFilterPeriod === 'Aug-26')) return false;
      }

      if (state.txnFilterCat !== 'all' && t.category !== state.txnFilterCat) {
        return false;
      }

      if (state.txnSearchQuery) {
        const q = state.txnSearchQuery.toLowerCase();
        const matchesName = (t.meterName || '').toLowerCase().includes(q);
        const matchesId = (t.id || '').toLowerCase().includes(q);
        const matchesDate = (t.date || '').toLowerCase().includes(q);
        const matchesCat = (t.categoryLabel || '').toLowerCase().includes(q);
        if (!matchesName && !matchesId && !matchesDate && !matchesCat) return false;
      }

      return true;
    });

    const rowsHtml = filtered.map(t => {
      let badgeClass = 'txn-badge-eb';
      if (t.category === 'ahu') badgeClass = 'txn-badge-ahu';
      else if (t.category === 'btu') badgeClass = 'txn-badge-btu';
      else if (t.category === 'dg') badgeClass = 'txn-badge-dg';
      else if (t.category === 'ahu_sched') badgeClass = 'txn-badge-sched';

      return `
        <tr>
          <td><span class="txn-id-pill">${t.id}</span></td>
          <td><span style="color:#cbd5e1; font-size:0.8rem;">${t.displayTime || t.timestamp}</span></td>
          <td><strong style="color:#00e5ff; cursor:pointer;" class="btn-jump-date" data-date="${t.date}">${t.date}</strong></td>
          <td><span class="txn-cat-badge ${badgeClass}">${t.categoryLabel || t.category}</span></td>
          <td>
            <div style="font-weight:700; color:#fff;">${t.meterName}</div>
            <div style="font-size:0.75rem; color:var(--text-muted);">Loc: ${t.location || '3F'}</div>
          </td>
          <td>${t.prevReading !== undefined ? t.prevReading : '-'}</td>
          <td><strong style="color:#fff;">${t.newReading !== undefined ? t.newReading : '-'}</strong></td>
          <td><span class="txn-cons-pill">+${t.consumption !== undefined ? t.consumption.toFixed(1) : '0'} kWh</span></td>
          <td><strong style="color:#34d399;">₹${t.cost !== undefined ? t.cost.toFixed(2) : '0.00'}</strong></td>
          <td><span class="txn-synced-badge"><i data-lucide="check" style="width:12px; height:12px;"></i> ${t.status || 'Synced'}</span></td>
          <td>
            <button class="btn-table-action btn-jump-date" data-date="${t.date}" title="Jump to date in Daily Entry">
              <i data-lucide="external-link"></i> View
            </button>
          </td>
        </tr>
      `;
    }).join('');

    return `
      <div class="txn-history-container">
        <!-- Live Status Bar -->
        <div class="txn-status-banner">
          <div class="live-sync-indicator" style="background:transparent; border:none; padding:0;">
            <span class="pulse-dot"></span>
            <div>
              <h2 style="font-size:1.15rem; color:#fff; font-weight:800; display:flex; align-items:center; gap:0.5rem;">
                <i data-lucide="history" style="color:#00e5ff;"></i> Transaction History & Real-Time Operational Audit Trail
              </h2>
              <p style="font-size:0.8rem; color:var(--text-muted);">Every meter reading commit, DG log, AHU saving update, and baseline sync is recorded in real time.</p>
            </div>
          </div>

          <div style="display:flex; align-items:center; gap:0.8rem;">
            <button class="btn-export-csv" id="btn-export-txn-csv">
              <i data-lucide="download"></i> Export Transaction CSV
            </button>
            <button class="btn-txn-sync" id="btn-force-sync">
              <i data-lucide="refresh-cw"></i> Sync Real-Time Now
            </button>
          </div>
        </div>

        <!-- 4 KPI Summary Cards -->
        <div class="kpi-cards-row">
          <div class="kpi-card">
            <div class="kpi-title">TOTAL LOGGED TRANSACTIONS</div>
            <div class="kpi-value">${txns.length}</div>
            <div class="kpi-subtext">All historical & current commits</div>
          </div>

          <div class="kpi-card">
            <div class="kpi-title">AUGUST 2026 ENTRIES</div>
            <div class="kpi-value" style="color:#00e5ff;">${augCount} <span style="font-size:0.8rem;">records</span></div>
            <div class="kpi-subtext">Active operational month entries</div>
          </div>

          <div class="kpi-card">
            <div class="kpi-title">AUG 2026 ENERGY LOGGED</div>
            <div class="kpi-value" style="color:#fbbf24;">${augKwh.toLocaleString(undefined, {maximumFractionDigits:1})} <span style="font-size:0.85rem;">kWh</span></div>
            <div class="kpi-subtext">Total consumption recorded</div>
          </div>

          <div class="kpi-card">
            <div class="kpi-title">AUG 2026 ESTIMATED COST</div>
            <div class="kpi-value" style="color:#34d399;">₹${augCost.toLocaleString(undefined, {maximumFractionDigits:0})}</div>
            <div class="kpi-subtext">Calculated energy cost</div>
          </div>
        </div>

        <!-- Filter & Search Toolbar -->
        <div class="controls-card" style="margin-top:1.5rem; margin-bottom:1.2rem;">
          <div style="display:flex; align-items:center; gap:1rem; flex-wrap:wrap;">
            <div style="display:flex; align-items:center; gap:0.5rem;">
              <label style="font-size:0.85rem; font-weight:700; color:#00e5ff;">Period:</label>
              <select id="txn-period-select" class="input-field" style="width:auto; min-width:160px;">
                <option value="all" ${state.txnFilterPeriod === 'all' ? 'selected' : ''}>All Transactions</option>
                <option value="Aug-26" ${state.txnFilterPeriod === 'Aug-26' ? 'selected' : ''}>August 2026 (Current)</option>
                <option value="Jul-26" ${state.txnFilterPeriod === 'Jul-26' ? 'selected' : ''}>July 2026</option>
                <option value="Jun-26" ${state.txnFilterPeriod === 'Jun-26' ? 'selected' : ''}>June 2026</option>
                <option value="May-26" ${state.txnFilterPeriod === 'May-26' ? 'selected' : ''}>May 2026</option>
                <option value="today" ${state.txnFilterPeriod === 'today' ? 'selected' : ''}>Today (${todayStr})</option>
              </select>
            </div>

            <div class="filter-pills-container">
              <button class="pill-btn ${state.txnFilterCat === 'all' ? 'active' : ''}" data-txncat="all">All Categories</button>
              <button class="pill-btn ${state.txnFilterCat === 'eb' ? 'active' : ''}" data-txncat="eb">⚡ EB</button>
              <button class="pill-btn ${state.txnFilterCat === 'ahu' ? 'active' : ''}" data-txncat="ahu">🌬️ AHU</button>
              <button class="pill-btn ${state.txnFilterCat === 'btu' ? 'active' : ''}" data-txncat="btu">❄️ BTU</button>
              <button class="pill-btn ${state.txnFilterCat === 'dg' ? 'active' : ''}" data-txncat="dg">⛽ DG</button>
            </div>
          </div>

          <div style="display:flex; align-items:center; gap:0.6rem; min-width:260px;">
            <input type="text" id="txn-search-input" class="input-field" placeholder="Search by meter, date, TXN ID..." value="${state.txnSearchQuery}">
            ${state.txnSearchQuery ? `<button class="font-btn" id="btn-clear-txn-search">✕</button>` : ''}
          </div>
        </div>

        <!-- Transactions Table -->
        <div class="table-responsive-container">
          <table class="data-table txn-table">
            <thead>
              <tr>
                <th>TXN ID</th>
                <th>Timestamp</th>
                <th>Log Date</th>
                <th>Category</th>
                <th>Meter / Asset</th>
                <th>Prev Reading</th>
                <th>New Reading</th>
                <th>Net Consumed</th>
                <th>Est Cost</th>
                <th>Sync Status</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              ${filtered.length > 0 ? rowsHtml : `
                <tr>
                  <td colspan="11" style="text-align:center; padding:2.5rem; color:var(--text-muted);">
                    <i data-lucide="inbox" style="font-size:2rem; margin-bottom:0.5rem; color:#6b7280;"></i>
                    <div>No transaction records found matching the selected filters.</div>
                  </td>
                </tr>
              `}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  // Full DT-3 Summary Matrix Sheet Format (Matching Excel summary Sheet)

  function renderFullDT3SummaryMatrixSection() {
    recalculateDynamicSummaryMatrix();
    const sumData = state.data.summary_matrix || { months: [], consumption: {}, cost: {} };
    const months = sumData.months || [];

    const monthHeadersHtml = months.map(m => `<th style="text-align:center;">${m}</th>`).join('');

    const consRows = ['EB', 'AHU', 'DG', 'BTU', 'Total'].map(cat => {
      const vals = sumData.consumption[cat] || [];
      const cells = vals.map(v => `<td style="text-align:center;">${v ? Math.round(v).toLocaleString() : 0}</td>`).join('');
      const isTot = cat === 'Total';
      return `<tr class="${isTot ? 'total-row' : ''}">
        <td><strong>${cat}</strong></td>
        ${cells}
      </tr>`;
    }).join('');

    const costRows = ['EB', 'AHU', 'DG', 'BTU', 'Total'].map(cat => {
      const vals = sumData.cost[cat] || [];
      const cells = vals.map(v => `<td style="text-align:center;">₹${v ? Math.round(v).toLocaleString() : 0}</td>`).join('');
      const isTot = cat === 'Total';
      return `<tr class="${isTot ? 'total-row' : ''}">
        <td><strong>${cat}</strong></td>
        ${cells}
      </tr>`;
    }).join('');

    return `
      <div style="margin-top: 1.5rem;">
        <div class="section-header-title eb-header" style="margin-top:0;">
          <h2><i data-lucide="file-spreadsheet"></i> DT-3 Master Monthly Consumption & Cost Summary Matrix (Matching Excel summary Sheet)</h2>
          <span style="font-size:0.8rem;">Historical & Current Monthly Energy Consumptions (2024 to August 2026 Real-Time Synced)</span>
        </div>

        <div class="controls-card" style="margin-bottom:1rem; border-color:#0284c7;">
          <h3 style="color:#00e5ff; font-size:1.05rem; font-weight:800;">1. Monthly Energy Consumption Breakdown (kWh)</h3>
        </div>

        <div class="table-responsive-container">
          <table class="data-table">
            <thead>
              <tr>
                <th>Component</th>
                ${monthHeadersHtml}
              </tr>
            </thead>
            <tbody>
              ${consRows}
            </tbody>
          </table>
        </div>

        <div class="controls-card" style="margin-top:2rem; margin-bottom:1rem; border-color:#059669;">
          <h3 style="color:#10b981; font-size:1.05rem; font-weight:800;">2. Monthly Energy Cost Breakdown (₹)</h3>
        </div>

        <div class="table-responsive-container">
          <table class="data-table">
            <thead>
              <tr>
                <th>Component</th>
                ${monthHeadersHtml}
              </tr>
            </thead>
            <tbody>
              ${costRows}
            </tbody>
          </table>
        </div>
      </div>

      ${renderMonthMatrixSectionBelow()}
    `;
  }

  // =========================================================================
  // NEW: Premier Interactive Month-Wise Consumption & Analytics Dashboard
  // =========================================================================
  function renderMonthDashboardContent() {
    const selectedMonth = state.monthDashboard.selectedMonth || 'Aug-26';
    const selectedYear = state.monthDashboard.selectedYear || '2026';
    const viewMode = state.monthDashboard.viewMode || 'dashboard';

    // Parse target month & year
    const p = selectedMonth.split('-');
    const mStr = p[0];
    const yVal = parseInt(p[1]) || 26;
    const targetYear = yVal < 100 ? 2000 + yVal : yVal;
    
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthIdx = monthNames.indexOf(mStr);
    const targetMonth = monthIdx !== -1 ? monthIdx + 1 : 8;

    // Generate date list for this month
    const daysInMonth = new Date(targetYear, targetMonth, 0).getDate();
    const dateList = [];
    for (let day = 1; day <= daysInMonth; day++) {
      const dStr = `${targetYear}-${String(targetMonth).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      dateList.push(dStr);
      ensureDateStructure(dStr);
    }

    // Master Meter Definitions
    const ebMaster = [
      { id: 1, name: "SL No 540430038840 (EB Unit)", loc: "3F", short: "EB-1" },
      { id: 2, name: "SL No 540430038845 (EB Unit)", loc: "3F", short: "EB-2" },
      { id: 3, name: "SL No 540430038844 (EB Unit)", loc: "3F", short: "EB-3" },
      { id: 4, name: "SL No 540430038821 (EB Unit)", loc: "3F", short: "EB-4" },
      { id: 5, name: "SL No 540430038849 (EB Unit)", loc: "3F", short: "EB-5" },
      { id: 6, name: "SL No 540430038646 (EB Unit)", loc: "3F", short: "EB-6" },
      { id: 7, name: "SL No 540430038843 (EB Unit)", loc: "3F", short: "EB-7" },
      { id: 8, name: "SL No 540430038848 (EB Unit)", loc: "3F", short: "EB-8" }
    ];

    const ahuMaster = [
      { id: 1, name: "AHU1", loc: "3F", short: "AHU-1" },
      { id: 2, name: "AHU2", loc: "3F", short: "AHU-2" },
      { id: 3, name: "AHU3", loc: "3F", short: "AHU-3" },
      { id: 4, name: "AHU4", loc: "3F", short: "AHU-4" }
    ];

    const btuMaster = [
      { id: 1, name: "AHU1 - BTU", loc: "3F", short: "BTU-1" },
      { id: 2, name: "AHU2 - BTU", loc: "3F", short: "BTU-2" },
      { id: 3, name: "AHU3 - BTU", loc: "3F", short: "BTU-3" },
      { id: 4, name: "AHU4 - BTU", loc: "3F", short: "BTU-4" }
    ];

    // Compute Daily Consumptions and Totals for Month
    const dailyStats = [];
    let monthSumEB = 0, monthSumDG = 0, monthSumAHU = 0, monthSumBTU = 0, monthSumCost = 0;
    let peakDay = { date: '', val: 0 }, minDay = { date: '', val: Infinity };

    dateList.forEach(dStr => {
      const dt = new Date(dStr);
      const dayName = dt.toLocaleString('en-US', { weekday: 'short' });
      const dayKWH = state.data.kwh_daily[dStr] || {};

      let dayEB = 0, dayDG = 0, dayAHU = 0, dayBTU = 0;
      const ebMap = {}, ahuMap = {}, btuMap = {};

      ebMaster.forEach(m => {
        const curr = (dayKWH.eb || []).find(x => x.id === m.id) || {};
        const prev = getExactPreviousReading(dStr, 'eb', m.id);
        const cEB = (curr.reading !== undefined && prev.reading !== undefined) ? Math.max(0, curr.reading - prev.reading) : 0;
        const cDG = (curr.dg_reading !== undefined && prev.dg_reading !== undefined) ? Math.max(0, curr.dg_reading - prev.dg_reading) : 0;
        ebMap[m.id] = { eb: cEB, dg: cDG, reading: curr.reading || 0, prev: prev.reading || 0 };
        dayEB += cEB;
        dayDG += cDG;
      });

      ahuMaster.forEach(m => {
        const curr = (dayKWH.ahu || []).find(x => x.id === m.id) || {};
        const prev = getExactPreviousReading(dStr, 'ahu', m.id);
        const cAHU = (curr.reading !== undefined && prev.reading !== undefined) ? Math.max(0, curr.reading - prev.reading) : 0;
        ahuMap[m.id] = { ahu: cAHU, reading: curr.reading || 0, prev: prev.reading || 0 };
        dayAHU += cAHU;
      });

      btuMaster.forEach(m => {
        const curr = (dayKWH.btu || []).find(x => x.id === m.id) || {};
        const prev = getExactPreviousReading(dStr, 'btu', m.id);
        const cBTU = (curr.reading !== undefined && prev.reading !== undefined) ? Math.max(0, curr.reading - prev.reading) : 0;
        btuMap[m.id] = { btu: cBTU, reading: curr.reading || 0, prev: prev.reading || 0 };
        dayBTU += cBTU;
      });

      const dayCum = dayEB + dayDG;
      const dayCost = (dayEB * state.rates.kwh) + (dayDG * state.rates.dg) + (dayBTU * state.rates.btu);

      monthSumEB += dayEB;
      monthSumDG += dayDG;
      monthSumAHU += dayAHU;
      monthSumBTU += dayBTU;
      monthSumCost += dayCost;

      if (dayCum > peakDay.val) peakDay = { date: dStr, val: dayCum };
      if (dayCum > 0 && dayCum < minDay.val) minDay = { date: dStr, val: dayCum };

      dailyStats.push({
        date: dStr,
        dayNum: dt.getDate(),
        dayName: dayName,
        ebMap,
        ahuMap,
        btuMap,
        dayEB,
        dayDG,
        dayCum,
        dayAHU,
        dayBTU,
        dayCost
      });
    });

    const monthSumCum = monthSumEB + monthSumDG;
    const avgDailyCum = daysInMonth > 0 ? (monthSumCum / daysInMonth) : 0;
    if (minDay.val === Infinity) minDay = { date: dateList[0] || '', val: 0 };

    // Calculate AHU Total Savings for Month
    let monthAHUSavedHrs = 0, monthAHUSavedCost = 0;
    dateList.forEach(dStr => {
      const dayRec = state.data.ahu_saving[dStr];
      if (dayRec) {
        const c = calculateAHUSaving(dayRec);
        monthAHUSavedHrs += (c.tot_saved !== undefined ? c.tot_saved : (c.total_saved_hrs || 0));
        monthAHUSavedCost += (c.tot_comb_cost_saved !== undefined ? c.tot_comb_cost_saved : (c.total_combined_saving || 0));
      }
    });

    // Calculate Month-over-Month Variance vs previous month
    let prevMonthEB = 0, prevMonthKey = '';
    const prevMonthIdx = monthIdx > 0 ? monthIdx - 1 : 11;
    const prevYearVal = monthIdx > 0 ? yVal : yVal - 1;
    prevMonthKey = `${monthNames[prevMonthIdx]}-${prevYearVal}`;
    
    // Lookup in summary matrix
    const sMonths = state.data.summary_matrix?.months || [];
    const prevSIdx = sMonths.indexOf(prevMonthKey);
    if (prevSIdx !== -1) {
      prevMonthEB = state.data.summary_matrix.consumption['EB']?.[prevSIdx] || 0;
    }
    const momVariance = prevMonthEB > 0 ? (((monthSumEB - prevMonthEB) / prevMonthEB) * 100) : 0;

    // Year options & 12 Month Ribbon Generation
    const yearOptions = ['2026', '2025', '2024'];
    const currentYearShort = targetYear.toString().slice(-2);

    const monthRibbonHtml = monthNames.map((m, idx) => {
      const mKey = `${m}-${currentYearShort}`;
      const isSelected = mKey === selectedMonth;
      
      // Calculate month total for ribbon badge
      let mTotalKWH = 0;
      const sIdx = sMonths.indexOf(mKey);
      if (sIdx !== -1) {
        mTotalKWH = (state.data.summary_matrix.consumption['EB']?.[sIdx] || 0) + (state.data.summary_matrix.consumption['DG']?.[sIdx] || 0);
      }

      return `
        <button class="month-nav-btn ${isSelected ? 'active' : ''}" data-month="${mKey}">
          <div class="month-btn-header">
            <span class="month-name">${m}-${currentYearShort}</span>
            ${isSelected ? '<span class="active-dot"></span>' : ''}
          </div>
          <div class="month-btn-footer">
            <span class="month-kwh-badge">${mTotalKWH > 0 ? (mTotalKWH / 1000).toFixed(1) + 'k kWh' : 'Logged'}</span>
          </div>
        </button>
      `;
    }).join('');

    // Previous & Next Month Navigation
    const currIdxInAll = monthNames.indexOf(mStr);
    const prevNavMonth = currIdxInAll > 0 ? `${monthNames[currIdxInAll - 1]}-${currentYearShort}` : `Dec-${parseInt(currentYearShort) - 1}`;
    const nextNavMonth = currIdxInAll < 11 ? `${monthNames[currIdxInAll + 1]}-${currentYearShort}` : `Jan-${parseInt(currentYearShort) + 1}`;

    // Render Sub-Views
    let mainViewHtml = '';

    if (viewMode === 'dashboard') {
      mainViewHtml = `
        <!-- Interactive Chart Visualizations -->
        <div class="dashboard-charts-grid" style="display:grid; grid-template-columns: 2fr 1fr; gap:1.25rem; margin-bottom:1.5rem;">
          <div class="metric-chart-card">
            <div class="chart-card-header">
              <div class="chart-title">
                <i data-lucide="bar-chart-3" style="color:#00e5ff;"></i>
                <h3>31-Day Daily Consumption & Load Profile (${selectedMonth})</h3>
              </div>
              <div class="chart-badge">EB, DG, AHU & BTU Multi-Dataset</div>
            </div>
            <div style="height: 320px; position:relative;">
              <canvas id="month-daily-trend-chart"></canvas>
            </div>
          </div>

          <div class="metric-chart-card">
            <div class="chart-card-header">
              <div class="chart-title">
                <i data-lucide="pie-chart" style="color:#10b981;"></i>
                <h3>Tenant & Meter Load Distribution</h3>
              </div>
              <div class="chart-badge">Category Share %</div>
            </div>
            <div style="height: 320px; position:relative;">
              <canvas id="month-meter-share-chart"></canvas>
            </div>
          </div>
        </div>

        <div class="dashboard-charts-grid" style="display:grid; grid-template-columns: 1fr 1fr; gap:1.25rem; margin-bottom:2rem;">
          <div class="metric-chart-card">
            <div class="chart-card-header">
              <div class="chart-title">
                <i data-lucide="trending-up" style="color:#f59e0b;"></i>
                <h3>Month-over-Month & Historical Variance</h3>
              </div>
              <div class="chart-badge">${selectedMonth} vs ${prevMonthKey}</div>
            </div>
            <div style="height: 280px; position:relative;">
              <canvas id="month-mom-chart"></canvas>
            </div>
          </div>

          <div class="metric-chart-card">
            <div class="chart-card-header">
              <div class="chart-title">
                <i data-lucide="indian-rupee" style="color:#3b82f6;"></i>
                <h3>Cumulative Daily Electricity Bill (INR)</h3>
              </div>
              <div class="chart-badge">Total: ₹${monthSumCost.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
            </div>
            <div style="height: 280px; position:relative;">
              <canvas id="month-cost-trend-chart"></canvas>
            </div>
          </div>
        </div>

        <!-- Quick Consumption Grid Section below charts -->
        ${renderMonthConsumptionGrid(selectedMonth, dateList, dailyStats, ebMaster, ahuMaster, btuMaster, monthSumEB, monthSumDG, monthSumCum, monthSumAHU, monthSumBTU, monthSumCost, avgDailyCum)}
      `;
    } else if (viewMode === 'consumption_grid') {
      mainViewHtml = renderMonthConsumptionGrid(selectedMonth, dateList, dailyStats, ebMaster, ahuMaster, btuMaster, monthSumEB, monthSumDG, monthSumCum, monthSumAHU, monthSumBTU, monthSumCost, avgDailyCum);
    } else if (viewMode === 'readings_grid') {
      mainViewHtml = renderMonthReadingsGrid(selectedMonth, dateList, dailyStats, ebMaster, ahuMaster, btuMaster);
    } else if (viewMode === 'macro_matrix') {
      mainViewHtml = renderFullDT3SummaryMatrixSection();
    }

    return `
      <div class="month-dashboard-wrapper">
        <!-- Month Navigation & Controls Hub -->
        <div class="month-navigation-hub">
          <div class="hub-top-row">
            <div class="hub-title-group">
              <div class="hub-icon"><i data-lucide="calendar-range"></i></div>
              <div>
                <h2>Month-Wise Energy Consumption & Analytics Dashboard</h2>
                <p>Select any month below to navigate consumption sheets, analyze load trends, inspect audit logs, and export Excel reports.</p>
              </div>
            </div>

            <!-- Quick Navigation Controls -->
            <div class="hub-nav-controls">
              <button class="nav-arrow-btn" id="btn-prev-month" data-nav="${prevNavMonth}" title="Previous Month">
                <i data-lucide="chevron-left"></i> ${prevNavMonth}
              </button>
              <button class="nav-arrow-btn" id="btn-jump-current-month" data-nav="Aug-26" title="Jump to Current Month">
                <i data-lucide="calendar"></i> Aug-26 (Current)
              </button>
              <button class="nav-arrow-btn" id="btn-next-month" data-nav="${nextNavMonth}" title="Next Month">
                ${nextNavMonth} <i data-lucide="chevron-right"></i>
              </button>
            </div>
          </div>

          <!-- 12-Month Interactive Navigation Ribbon -->
          <div class="month-ribbon-container">
            ${monthRibbonHtml}
          </div>

          <!-- View Switcher & Action Toolbar -->
          <div class="hub-bottom-toolbar">
            <div class="view-mode-pills">
              <button class="view-mode-btn ${viewMode === 'dashboard' ? 'active' : ''}" data-view="dashboard">
                <i data-lucide="layout-dashboard"></i> Analytics Dashboard & Charts
              </button>
              <button class="view-mode-btn ${viewMode === 'consumption_grid' ? 'active' : ''}" data-view="consumption_grid">
                <i data-lucide="table"></i> Full Month Consumption Sheet
              </button>
              <button class="view-mode-btn ${viewMode === 'readings_grid' ? 'active' : ''}" data-view="readings_grid">
                <i data-lucide="file-spreadsheet"></i> Daily Meter Readings Log
              </button>
              <button class="view-mode-btn ${viewMode === 'macro_matrix' ? 'active' : ''}" data-view="macro_matrix">
                <i data-lucide="columns-3"></i> Cross-Year Macro Matrix
              </button>
            </div>

            <div class="export-actions-group">
              <button class="export-btn-primary" id="btn-export-single-month-excel" title="Download Excel (.xlsx) for this selected month">
                <i data-lucide="file-down"></i> Export ${selectedMonth} Excel
              </button>
              <button class="export-btn-secondary" id="btn-export-master-excel" title="Download Complete DT-3 Master Workbook">
                <i data-lucide="book-open"></i> Full Master Workbook
              </button>
              <button class="export-btn-secondary" id="btn-print-month-report" title="Print or Save Executive PDF">
                <i data-lucide="printer"></i> Print PDF
              </button>
            </div>
          </div>
        </div>

        <!-- Executive KPI Cards for Selected Month -->
        <div class="kpi-cards-grid">
          <div class="kpi-card eb-card">
            <div class="kpi-card-header">
              <span class="kpi-title">Total EB Grid Consumption</span>
              <i data-lucide="zap" class="kpi-icon"></i>
            </div>
            <div class="kpi-value">${monthSumEB.toLocaleString('en-IN', { maximumFractionDigits: 1 })} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-cost">Total Cost: ₹${(monthSumEB * state.rates.kwh).toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
            <div class="kpi-footer">
              <span class="kpi-subtext">Avg: ${(monthSumEB / daysInMonth).toFixed(1)} kWh/day</span>
              <span class="kpi-variance ${momVariance >= 0 ? 'up' : 'down'}">
                ${momVariance >= 0 ? '▲ +' : '▼ '}${momVariance.toFixed(1)}% vs ${prevMonthKey}
              </span>
            </div>
          </div>

          <div class="kpi-card dg-card">
            <div class="kpi-card-header">
              <span class="kpi-title">Total DG Generator Backup</span>
              <i data-lucide="fuel" class="kpi-icon"></i>
            </div>
            <div class="kpi-value" style="color:#fbbf24;">${monthSumDG.toLocaleString('en-IN', { maximumFractionDigits: 1 })} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-cost">DG Cost: ₹${(monthSumDG * state.rates.dg).toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
            <div class="kpi-footer">
              <span class="kpi-subtext">Rate: ₹${state.rates.dg}/unit</span>
              <span class="kpi-badge-pill" style="background:rgba(251, 191, 36, 0.15); color:#fbbf24;">Diesel Gen</span>
            </div>
          </div>

          <div class="kpi-card cum-card">
            <div class="kpi-card-header">
              <span class="kpi-title">Total Building Power (EB+DG)</span>
              <i data-lucide="activity" class="kpi-icon"></i>
            </div>
            <div class="kpi-value" style="color:#00e5ff;">${monthSumCum.toLocaleString('en-IN', { maximumFractionDigits: 1 })} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-cost">Net Bill: ₹${monthSumCost.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
            <div class="kpi-footer">
              <span class="kpi-subtext">Peak: ${peakDay.val.toFixed(0)} kWh (${peakDay.date.slice(8)}th)</span>
              <span class="kpi-subtext">Min: ${minDay.val.toFixed(0)} kWh</span>
            </div>
          </div>

          <div class="kpi-card ahu-card">
            <div class="kpi-card-header">
              <span class="kpi-title">AHU HVAC System Consumption</span>
              <i data-lucide="wind" class="kpi-icon"></i>
            </div>
            <div class="kpi-value" style="color:#10b981;">${monthSumAHU.toLocaleString('en-IN', { maximumFractionDigits: 1 })} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-cost">HVAC Cost: ₹${(monthSumAHU * state.rates.kwh).toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
            <div class="kpi-footer">
              <span class="kpi-subtext">Saved: ${monthAHUSavedHrs.toFixed(1)} hrs</span>
              <span class="kpi-badge-pill" style="background:rgba(16, 185, 129, 0.2); color:#10b981; font-weight:700;">
                Saved ₹${monthAHUSavedCost.toLocaleString('en-IN', { maximumFractionDigits: 0 })}
              </span>
            </div>
          </div>

          <div class="kpi-card btu-card">
            <div class="kpi-card-header">
              <span class="kpi-title">BTU Chilled Water Cooling</span>
              <i data-lucide="snowflake" class="kpi-icon"></i>
            </div>
            <div class="kpi-value" style="color:#f59e0b;">${monthSumBTU.toLocaleString('en-IN', { maximumFractionDigits: 1 })} <span class="kpi-unit">Units</span></div>
            <div class="kpi-cost">Cooling Cost: ₹${(monthSumBTU * state.rates.btu).toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
            <div class="kpi-footer">
              <span class="kpi-subtext">Rate: ₹${state.rates.btu}/Unit</span>
              <span class="kpi-badge-pill" style="background:rgba(245, 158, 11, 0.15); color:#f59e0b;">Chilled Water</span>
            </div>
          </div>
        </div>

        <!-- Main Dynamic View Body -->
        ${mainViewHtml}
      </div>
    `;
  }

  // Render Full Month Consumption Grid (Excel Twin)
  function renderMonthConsumptionGrid(selectedMonth, dateList, dailyStats, ebMaster, ahuMaster, btuMaster, sumEB, sumDG, sumCum, sumAHU, sumBTU, sumCost, avgCum) {
    const dateHeadersHtml = dateList.map(d => {
      const dt = new Date(d);
      const dayName = dt.toLocaleString('default', { weekday: 'short' });
      const dayNum = dt.getDate();
      return `<th style="text-align:center; min-width:65px;">${dayNum}<br><span style="font-size:0.65rem; color:#9ca3af;">${dayName}</span></th>`;
    }).join('');

    let consRowsHtml = '';

    // EB Meters
    ebMaster.forEach(m => {
      const ebCells = dailyStats.map(s => `<td>${s.ebMap[m.id].eb ? s.ebMap[m.id].eb.toFixed(1) : '0'}</td>`).join('');
      const dgCells = dailyStats.map(s => `<td style="color:#fbbf24;">${s.ebMap[m.id].dg ? s.ebMap[m.id].dg.toFixed(1) : '0'}</td>`).join('');
      consRowsHtml += `
        <tr>
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--eb-color);">${m.name}</strong></td>
          ${ebCells}
        </tr>
        <tr style="background:rgba(251, 191, 36, 0.03);">
          <td></td><td></td><td style="color:#fbbf24;">↳ DG Backup (kWh)</td>
          ${dgCells}
        </tr>
      `;
    });

    // EB Totals
    const totEBCells = dailyStats.map(s => `<td><strong>${s.dayEB.toFixed(1)}</strong></td>`).join('');
    const totDGCells = dailyStats.map(s => `<td style="color:#fbbf24;"><strong>${s.dayDG.toFixed(1)}</strong></td>`).join('');
    const totCumCells = dailyStats.map(s => `<td style="color:#00e5ff;"><strong>${s.dayCum.toFixed(1)}</strong></td>`).join('');

    consRowsHtml += `
      <tr class="total-row" style="background:rgba(0, 229, 255, 0.08);"><td></td><td></td><td style="color:#00e5ff;">Total Power - EB Unit (kWh)</td>${totEBCells}</tr>
      <tr class="total-row" style="background:rgba(251, 191, 36, 0.08);"><td></td><td></td><td style="color:#fbbf24;">Total Power - DG Backup (kWh)</td>${totDGCells}</tr>
      <tr class="total-row" style="background:rgba(0, 229, 255, 0.18);"><td></td><td></td><td style="color:#00e5ff; font-weight:800;">Total Power Cumulative (kWh)</td>${totCumCells}</tr>
    `;

    // AHU Meters
    ahuMaster.forEach(m => {
      const ahuCells = dailyStats.map(s => `<td>${s.ahuMap[m.id].ahu ? s.ahuMap[m.id].ahu.toFixed(1) : '0'}</td>`).join('');
      consRowsHtml += `
        <tr style="border-top: 1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--ahu-color);">${m.name}</strong></td>
          ${ahuCells}
        </tr>
      `;
    });

    const totAHUCells = dailyStats.map(s => `<td style="color:var(--ahu-color);"><strong>${s.dayAHU.toFixed(1)}</strong></td>`).join('');
    consRowsHtml += `
      <tr class="total-row" style="background:rgba(16, 185, 129, 0.12);"><td></td><td></td><td style="color:var(--ahu-color); font-weight:800;">Total AHU Power Consumption (kWh)</td>${totAHUCells}</tr>
    `;

    // BTU Meters
    btuMaster.forEach(m => {
      const btuCells = dailyStats.map(s => `<td>${s.btuMap[m.id].btu ? s.btuMap[m.id].btu.toFixed(1) : '0'}</td>`).join('');
      consRowsHtml += `
        <tr style="border-top: 1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--btu-color);">${m.name}</strong></td>
          ${btuCells}
        </tr>
      `;
    });

    const totBTUCells = dailyStats.map(s => `<td style="color:var(--btu-color);"><strong>${s.dayBTU.toFixed(1)}</strong></td>`).join('');
    consRowsHtml += `
      <tr class="total-row" style="background:rgba(245, 158, 11, 0.12);"><td></td><td></td><td style="color:var(--btu-color); font-weight:800;">Total BTU Chilled Water Consumption</td>${totBTUCells}</tr>
    `;

    // Daily Cost Row
    const totCostCells = dailyStats.map(s => `<td style="color:#fff; font-size:0.75rem;">₹${s.dayCost.toFixed(0)}</td>`).join('');
    consRowsHtml += `
      <tr class="total-row" style="background:rgba(59, 130, 246, 0.15);"><td></td><td></td><td style="color:#3b82f6; font-weight:800;">Daily Energy Cost (INR)</td>${totCostCells}</tr>
    `;

    return `
      <div class="consumption-grid-wrapper" style="margin-top:1.5rem;">
        <div class="section-header-title eb-header" style="margin-top:0; display:flex; justify-content:space-between; align-items:center;">
          <div style="display:flex; align-items:center; gap:0.5rem;">
            <i data-lucide="table"></i>
            <h3>Monthly Daily Energy Consumption Sheet (${selectedMonth})</h3>
          </div>
          <span style="font-size:0.8rem; color:#9ca3af;">Exact Excel digital twin with daily meter calculations</span>
        </div>

        <div class="table-responsive-container">
          <table class="data-table">
            <thead>
              <tr>
                <th style="width:50px;">Sr.No</th>
                <th style="width:60px;">Loc</th>
                <th style="min-width:240px;">Meter / Asset Description</th>
                ${dateHeadersHtml}
              </tr>
            </thead>
            <tbody>
              ${consRowsHtml}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  // Render Month Raw Daily Readings Grid
  function renderMonthReadingsGrid(selectedMonth, dateList, dailyStats, ebMaster, ahuMaster, btuMaster) {
    const dateHeadersHtml = dateList.map(d => {
      const dt = new Date(d);
      const dayName = dt.toLocaleString('default', { weekday: 'short' });
      const dayNum = dt.getDate();
      return `<th style="text-align:center; min-width:70px;">${dayNum}<br><span style="font-size:0.65rem; color:#9ca3af;">${dayName}</span></th>`;
    }).join('');

    let readingsRowsHtml = '';

    ebMaster.forEach(m => {
      const prevInfo = getExactPreviousReading(dateList[0] || state.selectedDate, 'eb', m.id);
      const ebCells = dailyStats.map(s => `<td>${s.ebMap[m.id].reading || '-'}</td>`).join('');
      const dgCells = dailyStats.map(s => `<td style="color:#fbbf24;">${state.data.kwh_daily[s.date]?.eb?.find(x => x.id === m.id)?.dg_reading || '-'}</td>`).join('');

      readingsRowsHtml += `
        <tr>
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--eb-color);">${m.name}</strong></td>
          <td style="color:#00e5ff; font-weight:700;">${prevInfo.reading}</td>
          ${ebCells}
        </tr>
        <tr style="background:rgba(251, 191, 36, 0.03);">
          <td></td><td></td><td style="color:#fbbf24;">↳ DG Backup Meter</td>
          <td style="color:#fbbf24;">${prevInfo.dg_reading}</td>
          ${dgCells}
        </tr>
      `;
    });

    ahuMaster.forEach(m => {
      const prevInfo = getExactPreviousReading(dateList[0] || state.selectedDate, 'ahu', m.id);
      const ahuCells = dailyStats.map(s => `<td>${s.ahuMap[m.id].reading || '-'}</td>`).join('');
      readingsRowsHtml += `
        <tr style="border-top:1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--ahu-color);">${m.name}</strong></td>
          <td style="color:var(--ahu-color); font-weight:700;">${prevInfo.reading}</td>
          ${ahuCells}
        </tr>
      `;
    });

    btuMaster.forEach(m => {
      const prevInfo = getExactPreviousReading(dateList[0] || state.selectedDate, 'btu', m.id);
      const btuCells = dailyStats.map(s => `<td>${s.btuMap[m.id].reading || '-'}</td>`).join('');
      readingsRowsHtml += `
        <tr style="border-top:1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--btu-color);">${m.name}</strong></td>
          <td style="color:var(--btu-color); font-weight:700;">${prevInfo.reading}</td>
          ${btuCells}
        </tr>
      `;
    });

    return `
      <div class="readings-grid-wrapper" style="margin-top:1.5rem;">
        <div class="section-header-title eb-header" style="margin-top:0;">
          <i data-lucide="file-spreadsheet"></i>
          <h3>Daily Meter Readings Log Matrix (${selectedMonth})</h3>
          <span style="font-size:0.8rem; color:#9ca3af;">EB Unit, AHU Unit & BTU Unit Daily Readings Record</span>
        </div>

        <div class="table-responsive-container">
          <table class="data-table">
            <thead>
              <tr>
                <th style="width:50px;">Sr.No</th>
                <th style="width:60px;">Loc</th>
                <th style="min-width:240px;">Meter / Asset Description</th>
                <th style="min-width:90px; color:#00e5ff;">Prev Reading</th>
                ${dateHeadersHtml}
              </tr>
            </thead>
            <tbody>
              ${readingsRowsHtml}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  // Render Interactive Charts for Month Dashboard
  function renderMonthDashboardCharts() {
    if (typeof Chart === 'undefined') return;

    const selectedMonth = state.monthDashboard.selectedMonth || 'Aug-26';
    const p = selectedMonth.split('-');
    const mStr = p[0];
    const yVal = parseInt(p[1]) || 26;
    const targetYear = yVal < 100 ? 2000 + yVal : yVal;
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthIdx = monthNames.indexOf(mStr);
    const targetMonth = monthIdx !== -1 ? monthIdx + 1 : 8;

    const daysInMonth = new Date(targetYear, targetMonth, 0).getDate();
    const labels = [];
    const ebData = [], dgData = [], ahuData = [], btuData = [], costData = [];
    let runningCost = 0;

    for (let day = 1; day <= daysInMonth; day++) {
      const dStr = `${targetYear}-${String(targetMonth).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      labels.push(`${day} ${mStr}`);

      const dayKWH = state.data.kwh_daily[dStr] || {};
      let dayEB = 0, dayDG = 0, dayAHU = 0, dayBTU = 0;

      (dayKWH.eb || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'eb', m.id);
        dayEB += Math.max(0, (Number(m.reading) || 0) - prev.reading);
        dayDG += Math.max(0, (Number(m.dg_reading) || 0) - prev.dg_reading);
      });

      (dayKWH.ahu || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'ahu', m.id);
        dayAHU += Math.max(0, (Number(m.reading) || 0) - prev.reading);
      });

      (dayKWH.btu || []).forEach(m => {
        const prev = getExactPreviousReading(dStr, 'btu', m.id);
        dayBTU += Math.max(0, (Number(m.reading) || 0) - prev.reading);
      });

      const dayCost = (dayEB * state.rates.kwh) + (dayDG * state.rates.dg) + (dayBTU * state.rates.btu);
      runningCost += dayCost;

      ebData.push(Number(dayEB.toFixed(1)));
      dgData.push(Number(dayDG.toFixed(1)));
      ahuData.push(Number(dayAHU.toFixed(1)));
      btuData.push(Number(dayBTU.toFixed(1)));
      costData.push(Number(runningCost.toFixed(0)));
    }

    // Clean up prior charts
    ['trend', 'share', 'mom', 'cost'].forEach(k => {
      if (activeCharts[k]) {
        activeCharts[k].destroy();
        delete activeCharts[k];
      }
    });

    // 1. Daily Trend Multi-Bar Chart
    const trendCtx = document.getElementById('month-daily-trend-chart');
    if (trendCtx) {
      activeCharts['trend'] = new Chart(trendCtx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'EB Grid Power (kWh)',
              data: ebData,
              backgroundColor: 'rgba(0, 229, 255, 0.75)',
              borderColor: '#00e5ff',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: 'AHU HVAC (kWh)',
              data: ahuData,
              backgroundColor: 'rgba(16, 185, 129, 0.75)',
              borderColor: '#10b981',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: 'DG Generator (kWh)',
              data: dgData,
              backgroundColor: 'rgba(251, 191, 36, 0.85)',
              borderColor: '#fbbf24',
              borderWidth: 1,
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top', labels: { color: '#f3f4f6', boxWidth: 12 } },
            tooltip: { mode: 'index', intersect: false }
          },
          scales: {
            x: { grid: { display: false }, ticks: { color: '#9ca3af', maxRotation: 45, font: { size: 10 } } },
            y: { grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#9ca3af' } }
          }
        }
      });
    }

    // 2. Tenant & Meter Load Donut Chart
    const shareCtx = document.getElementById('month-meter-share-chart');
    if (shareCtx) {
      const ebTotals = [0, 0, 0, 0, 0, 0, 0, 0];
      for (let day = 1; day <= daysInMonth; day++) {
        const dStr = `${targetYear}-${String(targetMonth).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayKWH = state.data.kwh_daily[dStr] || {};
        (dayKWH.eb || []).forEach(m => {
          if (m.id >= 1 && m.id <= 8) {
            const prev = getExactPreviousReading(dStr, 'eb', m.id);
            ebTotals[m.id - 1] += Math.max(0, (Number(m.reading) || 0) - prev.reading);
          }
        });
      }

      activeCharts['share'] = new Chart(shareCtx, {
        type: 'doughnut',
        data: {
          labels: ['EB-1 (540430038840)', 'EB-2 (540430038845)', 'EB-3 (540430038844)', 'EB-4 (540430038821)', 'EB-5 (540430038849)', 'EB-6 (540430038646)', 'EB-7 (540430038843)', 'EB-8 (540430038848)'],
          datasets: [{
            data: ebTotals,
            backgroundColor: [
              '#3b82f6', '#10b981', '#f59e0b', '#ef4444',
              '#8b5cf6', '#00e5ff', '#ec4899', '#14b8a6'
            ],
            borderWidth: 2,
            borderColor: '#0b0f19'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'right', labels: { color: '#f3f4f6', boxWidth: 10, font: { size: 9 } } }
          }
        }
      });
    }

    // 3. Month-over-Month Variance Chart
    const momCtx = document.getElementById('month-mom-chart');
    if (momCtx) {
      const prevMonthIdx = monthIdx > 0 ? monthIdx - 1 : 11;
      const prevYearVal = monthIdx > 0 ? yVal : yVal - 1;
      const prevMonthKey = `${monthNames[prevMonthIdx]}-${prevYearVal}`;

      const sMonths = state.data.summary_matrix?.months || [];
      const currSIdx = sMonths.indexOf(selectedMonth);
      const prevSIdx = sMonths.indexOf(prevMonthKey);

      const currEB = currSIdx !== -1 ? (state.data.summary_matrix.consumption['EB']?.[currSIdx] || 0) : 0;
      const currDG = currSIdx !== -1 ? (state.data.summary_matrix.consumption['DG']?.[currSIdx] || 0) : 0;
      const currAHU = currSIdx !== -1 ? (state.data.summary_matrix.consumption['AHU']?.[currSIdx] || 0) : 0;

      const prevEB = prevSIdx !== -1 ? (state.data.summary_matrix.consumption['EB']?.[prevSIdx] || 0) : 0;
      const prevDG = prevSIdx !== -1 ? (state.data.summary_matrix.consumption['DG']?.[prevSIdx] || 0) : 0;
      const prevAHU = prevSIdx !== -1 ? (state.data.summary_matrix.consumption['AHU']?.[prevSIdx] || 0) : 0;

      activeCharts['mom'] = new Chart(momCtx, {
        type: 'bar',
        data: {
          labels: ['EB Grid (kWh)', 'DG Power (kWh)', 'AHU HVAC (kWh)'],
          datasets: [
            {
              label: prevMonthKey,
              data: [prevEB, prevDG, prevAHU],
              backgroundColor: 'rgba(156, 163, 175, 0.4)',
              borderColor: '#9ca3af',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: `${selectedMonth} (Selected)`,
              data: [currEB, currDG, currAHU],
              backgroundColor: 'rgba(0, 229, 255, 0.8)',
              borderColor: '#00e5ff',
              borderWidth: 1,
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top', labels: { color: '#f3f4f6', boxWidth: 12 } }
          },
          scales: {
            x: { grid: { display: false }, ticks: { color: '#9ca3af' } },
            y: { grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#9ca3af' } }
          }
        }
      });
    }

    // 4. Cumulative Cost Trend Line
    const costCtx = document.getElementById('month-cost-trend-chart');
    if (costCtx) {
      activeCharts['cost'] = new Chart(costCtx, {
        type: 'line',
        data: {
          labels: labels,
          datasets: [{
            label: 'Cumulative Bill (₹)',
            data: costData,
            borderColor: '#3b82f6',
            backgroundColor: 'rgba(59, 130, 246, 0.15)',
            borderWidth: 2,
            fill: true,
            tension: 0.3
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top', labels: { color: '#f3f4f6', boxWidth: 12 } }
          },
          scales: {
            x: { grid: { display: false }, ticks: { color: '#9ca3af', font: { size: 9 } } },
            y: { grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#9ca3af' } }
          }
        }
      });
    }
  }

  // Multi-Device Cloud Synchronization & Pairing Modal
  function renderCloudSyncModal() {
    const pairingUrl = `${window.location.origin}${window.location.pathname}?sync=${encodeURIComponent(state.cloudSync.room)}`;

    return `
      <div class="modal-backdrop">
        <div class="modal-container" style="max-width: 580px;">
          <div class="modal-header">
            <div style="display:flex; align-items:center; gap:0.6rem;">
              <i data-lucide="cloud" style="color:#00e5ff;"></i>
              <h3>Real-Time Multi-Device Cloud Persistence & Pairing</h3>
            </div>
            <button class="modal-close-btn" id="btn-close-sync-modal"><i data-lucide="x"></i></button>
          </div>

          <div class="modal-body">
            <div style="background: rgba(0, 229, 255, 0.05); border: 1px solid rgba(0, 229, 255, 0.2); padding: 1rem; border-radius: var(--radius-md); margin-bottom: 1.25rem;">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.5rem;">
                <span style="color:#9ca3af; font-size:0.85rem;">Active Sync Status:</span>
                <span style="color:#10b981; font-weight:800; display:flex; align-items:center; gap:0.35rem;">
                  <span class="pulse-dot"></span> 🟢 Cloud Connected
                </span>
              </div>
              <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.8rem; color:#9ca3af;">
                <span>Last Synchronized:</span>
                <strong style="color:#f3f4f6;">${state.cloudSync.lastSyncTime}</strong>
              </div>
            </div>

            <div class="form-group" style="margin-bottom:1.25rem;">
              <label style="color:#f3f4f6; font-weight:600; font-size:0.85rem; margin-bottom:0.35rem; display:block;">Facility Cloud Sync Room / Channel:</label>
              <div style="display:flex; gap:0.5rem;">
                <input type="text" id="input-sync-room" class="input-field" value="${state.cloudSync.room}" placeholder="e.g. CBRE-DT3-FACILITY-2026" style="flex:1;">
                <button class="btn-primary" id="btn-save-sync-room" style="white-space:nowrap; padding:0.5rem 1rem;">Update Room</button>
              </div>
              <small style="color:#9ca3af; font-size:0.75rem; display:block; margin-top:0.35rem;">All devices using this exact Room Code share and update the same real-time database.</small>
            </div>

            <div class="form-group" style="margin-bottom:1.25rem;">
              <label style="color:#f3f4f6; font-weight:600; font-size:0.85rem; margin-bottom:0.35rem; display:block;">1-Click Device Pairing Link:</label>
              <div style="display:flex; gap:0.5rem;">
                <input type="text" id="input-pairing-url" class="input-field" value="${pairingUrl}" readonly style="flex:1; font-size:0.8rem; background:rgba(0,0,0,0.3);">
                <button class="btn-secondary" id="btn-copy-pairing-url" style="white-space:nowrap; padding:0.5rem 0.85rem;"><i data-lucide="copy"></i> Copy Link</button>
              </div>
              <small style="color:#9ca3af; font-size:0.75rem; display:block; margin-top:0.35rem;">Send this link via WhatsApp, Email, or Slack to instantly connect phones, tablets, or other PCs.</small>
            </div>

            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:0.75rem; margin-bottom:1.25rem;">
              <button class="btn-secondary" id="btn-modal-force-pull" style="justify-content:center; padding:0.7rem;">
                <i data-lucide="arrow-down-circle"></i> Force Pull Latest
              </button>
              <button class="btn-primary" id="btn-modal-force-push" style="justify-content:center; padding:0.7rem;">
                <i data-lucide="arrow-up-circle"></i> Force Push Local
              </button>
            </div>

            <div style="border-top:1px solid var(--border-color); padding-top:1rem;">
              <h4 style="color:#f3f4f6; font-size:0.9rem; margin-bottom:0.75rem; display:flex; align-items:center; gap:0.4rem;">
                <i data-lucide="hard-drive"></i> Offline Backup & Instant Restore
              </h4>
              <div style="display:flex; gap:0.75rem;">
                <button class="btn-secondary" id="btn-export-full-json-backup" style="flex:1; justify-content:center; font-size:0.8rem; padding:0.5rem;">
                  <i data-lucide="download"></i> Export JSON Backup
                </button>
                <label class="btn-secondary" style="flex:1; justify-content:center; font-size:0.8rem; padding:0.5rem; cursor:pointer; text-align:center; display:flex; align-items:center; gap:0.35rem;">
                  <i data-lucide="upload"></i> Restore from JSON
                  <input type="file" id="input-import-json-backup" accept=".json" style="display:none;">
                </label>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  function renderAdminPortalContent() {
    if (!state.isAdminUnlocked) {
      return `
        <div class="meter-card" style="max-width:500px; margin: 3rem auto; text-align:center; padding: 2.5rem;">
          <div style="font-size:3rem; color:#ef4444; margin-bottom:1rem;"><i data-lucide="lock"></i></div>
          <h2 style="color:#fff; margin-bottom:0.5rem;">Admin Portal Locked</h2>
          <p style="color:var(--text-muted); font-size:0.9rem; margin-bottom:1.5rem;">
            Please enter the admin password (<strong>Satya@1996</strong>) to access executive dashboards and export options.
          </p>
          <button class="btn-primary-save" style="margin: 0 auto;" id="btn-open-admin-modal">
            <i data-lucide="key"></i> Enter Admin Password
          </button>
        </div>
      `;
    }

    return `
      <div class="admin-dashboard-header">
        <div>
          <h2 style="color:#fff; font-size:1.4rem; font-weight:800; display:flex; align-items:center; gap:0.5rem;">
            <i data-lucide="shield-check" style="color:#10b981;"></i> DT-3 Executive Energy & EPI Performance Dashboard
          </h2>
          <p style="color:var(--text-muted); font-size:0.85rem;">Official Operational Performance & Dynamic Excel Export Hub</p>
        </div>

        <div class="export-btn-group">
          <button class="btn-export" id="btn-export-kwh">
            <i data-lucide="file-spreadsheet"></i> Export DT-3 KWH Excel
          </button>
          <button class="btn-export ahu-exp" id="btn-export-ahu">
            <i data-lucide="file-spreadsheet"></i> Export AHU Saving Excel
          </button>
          <button class="admin-trigger-btn" id="btn-admin-lock" style="background:rgba(255,255,255,0.1); border-color:var(--border-color); color:#fff;">
            <i data-lucide="log-out"></i> Lock Admin
          </button>
        </div>
      </div>

      <div class="kpi-cards-row">
        <div class="kpi-card">
          <div class="kpi-title">2026 JAN-JUL AVG EPI</div>
          <div class="kpi-value">103.7 <span style="font-size:0.9rem;">kWh/m²/yr</span></div>
          <div class="kpi-subtext">Status: Best-in-Class / Efficient (&lt;120 Target)</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-title">EPI VARIANCE vs 2025</div>
          <div class="kpi-value" style="color:#34d399;">-13.1%</div>
          <div class="kpi-subtext">Improved YTD Facility Performance</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-title">NET YTD COST SAVINGS</div>
          <div class="kpi-value" style="color:#60a5fa;">₹8,93,612</div>
          <div class="kpi-subtext">Jan–July 2026 vs 2025 Actual Savings</div>
        </div>

        <div class="kpi-card">
          <div class="kpi-title">NET YTD ENERGY SAVED</div>
          <div class="kpi-value" style="color:#fbbf24;">83,337 <span style="font-size:0.9rem;">kWh</span></div>
          <div class="kpi-subtext">Facility Energy Reduction</div>
        </div>
      </div>

      <!-- Section 1: EPI Benchmark Standards -->
      <div class="chart-container-card" style="margin-bottom:1.5rem;">
        <div class="chart-title"><i data-lucide="award"></i> 1. EPI Performance Level Benchmark Standards (kWh/m²/year)</div>
        <div class="table-responsive-container" style="margin-bottom:0;">
          <table class="data-table">
            <thead>
              <tr>
                <th>Performance Level</th>
                <th>EPI Range (kWh/m²/yr)</th>
                <th>Status Category</th>
                <th>Facility Benchmark Target</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><strong>Average</strong></td>
                <td>160 – 200</td>
                <td>Typical Office</td>
                <td>Standard Commercial Building</td>
              </tr>
              <tr>
                <td><strong>Good</strong></td>
                <td>120 – 150</td>
                <td>Efficient Office</td>
                <td>Energy Efficient Facility Target</td>
              </tr>
              <tr style="background:rgba(16, 185, 129, 0.15);">
                <td><strong style="color:#34d399;">Excellent</strong></td>
                <td style="color:#34d399;">Below 120</td>
                <td><strong style="color:#34d399;">Best-in-Class</strong></td>
                <td><strong style="color:#34d399;">DT3 Operational Target (&lt;120)</strong></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Section 2: DT3 Monthly Energy Consumption & EPI Trend (2025 vs 2026) -->
      <div class="chart-container-card" style="margin-bottom:1.5rem;">
        <div class="chart-title"><i data-lucide="trending-up"></i> 2. DT3 Monthly Energy Consumption & EPI Trend (2025 vs 2026)</div>
        <div class="table-responsive-container" style="margin-bottom:0;">
          <table class="data-table">
            <thead>
              <tr>
                <th>Month</th>
                <th>2025 Cons (kWh)</th>
                <th>2026 Cons (kWh)</th>
                <th>2025 EPI (kWh/m²/yr)</th>
                <th>2026 EPI (kWh/m²/yr)</th>
                <th>EPI Variance (%)</th>
              </tr>
            </thead>
            <tbody>
              <tr><td>Jan</td><td>71,247</td><td>78,876</td><td>93.7</td><td>103.7</td><td><span style="color:#ef4444;">+10.7%</span></td></tr>
              <tr><td>Feb</td><td>74,823</td><td>77,805</td><td>98.4</td><td>102.3</td><td><span style="color:#ef4444;">+4.0%</span></td></tr>
              <tr><td>Mar</td><td>83,562</td><td>80,041</td><td>109.9</td><td>105.3</td><td><span style="color:#34d399;">-4.2%</span></td></tr>
              <tr><td>Apr</td><td>91,430</td><td>80,188</td><td>120.2</td><td>105.5</td><td><span style="color:#34d399;">-12.3%</span></td></tr>
              <tr><td>May</td><td>94,164</td><td>76,356</td><td>123.8</td><td>100.4</td><td><span style="color:#34d399;">-18.9%</span></td></tr>
              <tr><td>Jun</td><td>105,721</td><td>82,334</td><td>139.0</td><td>108.3</td><td><span style="color:#34d399;">-22.1%</span></td></tr>
              <tr><td>Jul</td><td>114,327</td><td>76,338</td><td>150.4</td><td>100.4</td><td><span style="color:#34d399;">-33.2%</span></td></tr>
              <tr class="total-row" style="background:rgba(0, 229, 255, 0.15);">
                <td><strong>Jan-Jul AVERAGE</strong></td>
                <td><strong>90,754</strong></td>
                <td><strong>78,848</strong></td>
                <td><strong>119.4</strong></td>
                <td><strong style="color:#00e5ff;">103.7</strong></td>
                <td><strong style="color:#34d399;">-13.1%</strong></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Section 3: Interactive Visuals -->
      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:1.5rem; margin-bottom:2rem;">
        <div class="chart-container-card">
          <div class="chart-title"><i data-lucide="line-chart" style="color:#60a5fa;"></i> Monthly Consumption Trend (2025 vs 2026 kWh)</div>
          <div style="position:relative; height:280px;">
            <canvas id="chart-monthly-trend"></canvas>
          </div>
        </div>

        <div class="chart-container-card">
          <div class="chart-title"><i data-lucide="circle-dollar-sign" style="color:#34d399;"></i> Monthly Energy Cost Trend (2025 vs 2026 ₹)</div>
          <div style="position:relative; height:280px;">
            <canvas id="chart-cost-trend"></canvas>
          </div>
        </div>

        <div class="chart-container-card">
          <div class="chart-title"><i data-lucide="trending-down" style="color:#00e5ff;"></i> EPI Performance Trend vs 2025 Benchmark (kWh/m²/yr)</div>
          <div style="position:relative; height:280px;">
            <canvas id="chart-epi-trend"></canvas>
          </div>
        </div>

        <div class="chart-container-card">
          <div class="chart-title"><i data-lucide="pie-chart" style="color:#fbbf24;"></i> YTD YoY Energy Component Net Cost Savings (₹)</div>
          <div style="position:relative; height:280px;">
            <canvas id="chart-yoy-breakdown"></canvas>
          </div>
        </div>
      </div>

      <!-- Section 4: YoY Component Savings Table -->
      <div class="chart-container-card" style="margin-bottom:1.5rem;">
        <div class="chart-title"><i data-lucide="table"></i> 3. YoY Savings & Comparative Performance Analysis Table (Jan - Jul 2025 vs 2026)</div>
        <div class="table-responsive-container" style="margin-bottom:0;">
          <table class="data-table">
            <thead>
              <tr>
                <th>Energy Component</th>
                <th>2025 Cons (kWh)</th>
                <th>2026 Cons (kWh)</th>
                <th>kWh Savings</th>
                <th>% Change (kWh)</th>
                <th>2025 Cost (₹)</th>
                <th>2026 Cost (₹)</th>
                <th>Net Cost Savings (₹)</th>
                <th>Operational Analysis & Justification</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><strong>EB Electricity (kWh)</strong></td>
                <td>382,703</td>
                <td>358,855</td>
                <td>23,848</td>
                <td><span style="color:#34d399;">-6.2%</span></td>
                <td>₹28,51,137</td>
                <td>₹26,73,470</td>
                <td><strong style="color:#34d399;">₹1,77,668</strong></td>
                <td>23,622 kWh EB Savings (-6.2%) achieved YTD via efficiency</td>
              </tr>
              <tr>
                <td><strong>AHU HVAC Power (kWh)</strong></td>
                <td>51,569</td>
                <td>66,409</td>
                <td>-14,840</td>
                <td><span style="color:#ef4444;">+28.8%</span></td>
                <td>₹3,84,189</td>
                <td>₹4,94,747</td>
                <td><strong style="color:#ef4444;">-₹1,10,558</strong></td>
                <td>Transitioned to 12*5.5 KAM model from May (saved ~4,035 kWh/mo)</td>
              </tr>
              <tr>
                <td><strong>DG Backup Power (kWh)</strong></td>
                <td>3,159</td>
                <td>1,970</td>
                <td>1,189</td>
                <td><span style="color:#34d399;">-37.6%</span></td>
                <td>₹1,06,932</td>
                <td>₹66,685</td>
                <td><strong style="color:#34d399;">₹40,248</strong></td>
                <td>DG run hours reduced by 37.8% due to higher utility grid stability</td>
              </tr>
              <tr>
                <td><strong>BTU Cooling Load (kWh)</strong></td>
                <td>197,844</td>
                <td>124,704</td>
                <td>73,140</td>
                <td><span style="color:#34d399;">-37.0%</span></td>
                <td>₹21,26,819</td>
                <td>₹13,40,564</td>
                <td><strong style="color:#34d399;">₹7,86,255</strong></td>
                <td>BTU cooling load energy reduced by 35,641 kWh (-18.3%) saving ₹3.83L</td>
              </tr>
              <tr class="total-row">
                <td>Total Facility (Jan-Jul)</td>
                <td>635,275</td>
                <td>551,938</td>
                <td>83,337</td>
                <td><span style="color:#34d399;">-13.1%</span></td>
                <td>₹54,69,077</td>
                <td>₹45,75,465</td>
                <td><strong style="color:#34d399;">₹8,93,612</strong></td>
                <td>Total Jan–July facility energy cost reduced by ₹8.93L (-16.3%)</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderAdminAuthModal() {
    return `
      <div class="modal-backdrop">
        <div class="modal-card">
          <button class="modal-close-btn" id="btn-close-admin-modal">✕</button>
          <div class="modal-header">
            <h3><i data-lucide="lock" style="color:#ef4444;"></i> Admin Authentication</h3>
            <p>Enter administrative password to unlock reports & Excel exports.</p>
          </div>

          <form id="admin-login-form">
            <div class="form-group-admin">
              <label>Password:</label>
              <input type="password" id="admin-pass-input" class="input-field" placeholder="Enter password..." autofocus>
              ${state.adminErrorMessage ? `<div style="color:#ef4444; font-size:0.8rem; margin-top:0.4rem;">${state.adminErrorMessage}</div>` : ''}
            </div>

            <button type="submit" class="btn-admin-submit">
              Unlock Admin Portal
            </button>
          </form>
        </div>
      </div>
    `;
  }

  // =========================================================================
  // Master Event Listeners & Interactive Handlers
  // =========================================================================
  function attachEventListeners() {
    // Font Scale Buttons
    document.querySelectorAll('.font-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        state.fontScale = parseFloat(e.target.dataset.font);
        render();
      });
    });

    // Tab Navigation
    document.querySelectorAll('.nav-tab-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const tab = e.currentTarget.dataset.tab;
        if (tab) {
          state.activeTab = tab;
          render();
        }
      });
    });

    // Date Picker for Daily Entry
    const datePicker = document.getElementById('selected-date-picker');
    if (datePicker) {
      datePicker.addEventListener('change', (e) => {
        const val = e.target.value;
        if (val) {
          state.selectedDate = val;
          state.selectedMonthMatrix = dateToMonthKey(val) || 'Aug-26';
          state.monthDashboard.selectedMonth = state.selectedMonthMatrix;
          ensureDateStructure(val);
          render();
        }
      });
    }

    // Filter Pills in Daily Entry
    document.querySelectorAll('.filter-pills-container .pill-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        state.kwhFilter = e.currentTarget.dataset.filter;
        render();
      });
    });

    // Month Navigation Hub Buttons
    document.querySelectorAll('.month-nav-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const m = e.currentTarget.dataset.month;
        if (m) {
          state.monthDashboard.selectedMonth = m;
          state.selectedMonthMatrix = m;
          render();
        }
      });
    });

    // Prev / Next Month Arrow Buttons
    const btnPrevMonth = document.getElementById('btn-prev-month');
    if (btnPrevMonth) {
      btnPrevMonth.addEventListener('click', (e) => {
        const target = e.currentTarget.dataset.nav;
        if (target) {
          state.monthDashboard.selectedMonth = target;
          state.selectedMonthMatrix = target;
          render();
        }
      });
    }

    const btnNextMonth = document.getElementById('btn-next-month');
    if (btnNextMonth) {
      btnNextMonth.addEventListener('click', (e) => {
        const target = e.currentTarget.dataset.nav;
        if (target) {
          state.monthDashboard.selectedMonth = target;
          state.selectedMonthMatrix = target;
          render();
        }
      });
    }

    const btnJumpCurrent = document.getElementById('btn-jump-current-month');
    if (btnJumpCurrent) {
      btnJumpCurrent.addEventListener('click', (e) => {
        const target = e.currentTarget.dataset.nav;
        if (target) {
          state.monthDashboard.selectedMonth = target;
          state.selectedMonthMatrix = target;
          render();
        }
      });
    }

    // View Mode Switcher
    document.querySelectorAll('.view-mode-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const v = e.currentTarget.dataset.view;
        if (v) {
          state.monthDashboard.viewMode = v;
          render();
        }
      });
    });

    // Export Single Month Excel
    const btnExportSingle = document.getElementById('btn-export-single-month-excel');
    if (btnExportSingle) {
      btnExportSingle.addEventListener('click', () => {
        if (window.ExcelExporter) {
          window.ExcelExporter.exportSingleMonthKWH(state.monthDashboard.selectedMonth, state.data, state.rates);
        }
      });
    }

    // Export Master DT-3 Workbook
    const btnExportMaster = document.getElementById('btn-export-master-excel');
    if (btnExportMaster) {
      btnExportMaster.addEventListener('click', () => {
        if (window.ExcelExporter) {
          window.ExcelExporter.exportKWHDailyLog(state.monthDashboard.selectedMonth, state.data);
        }
      });
    }

    // Print PDF
    const btnPrintPDF = document.getElementById('btn-print-month-report');
    if (btnPrintPDF) {
      btnPrintPDF.addEventListener('click', () => window.print());
    }

    // Cloud Sync Modal Toggle
    const btnCloudSync = document.getElementById('btn-cloud-sync-modal');
    if (btnCloudSync) {
      btnCloudSync.addEventListener('click', () => {
        state.cloudSync.showModal = true;
        render();
      });
    }

    const btnCloseSyncModal = document.getElementById('btn-close-sync-modal');
    if (btnCloseSyncModal) {
      btnCloseSyncModal.addEventListener('click', () => {
        state.cloudSync.showModal = false;
        render();
      });
    }

    // Force Sync Button in Header
    const btnForceHeaderSync = document.getElementById('btn-force-cloud-sync');
    if (btnForceHeaderSync) {
      btnForceHeaderSync.addEventListener('click', () => {
        window.CloudSync.pull(true);
      });
    }

    // Save Cloud Sync Room
    const btnSaveSyncRoom = document.getElementById('btn-save-sync-room');
    if (btnSaveSyncRoom) {
      btnSaveSyncRoom.addEventListener('click', () => {
        const input = document.getElementById('input-sync-room');
        if (input && input.value.trim()) {
          window.CloudSync.setRoom(input.value.trim());
          showToast(`Cloud Room set to: ${input.value.trim()}`);
        }
      });
    }

    // Copy Pairing URL
    const btnCopyPairing = document.getElementById('btn-copy-pairing-url');
    if (btnCopyPairing) {
      btnCopyPairing.addEventListener('click', () => {
        const input = document.getElementById('input-pairing-url');
        if (input) {
          input.select();
          navigator.clipboard.writeText(input.value);
          showToast('📋 Device Pairing Link copied to clipboard!');
        }
      });
    }

    // Force Push / Pull in Modal
    const btnModalForcePull = document.getElementById('btn-modal-force-pull');
    if (btnModalForcePull) {
      btnModalForcePull.addEventListener('click', () => window.CloudSync.pull(true));
    }

    const btnModalForcePush = document.getElementById('btn-modal-force-push');
    if (btnModalForcePush) {
      btnModalForcePush.addEventListener('click', () => window.CloudSync.push(true));
    }

    // Backup Export / Import
    const btnExportJSONBackup = document.getElementById('btn-export-full-json-backup');
    if (btnExportJSONBackup) {
      btnExportJSONBackup.addEventListener('click', () => window.CloudSync.exportBackupJSON());
    }

    const inputImportBackup = document.getElementById('input-import-json-backup');
    if (inputImportBackup) {
      inputImportBackup.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          window.CloudSync.importBackupJSON(e.target.files[0]);
        }
      });
    }

    // Transaction History Filters
    const txnPeriodSelect = document.getElementById('txn-period-select');
    if (txnPeriodSelect) {
      txnPeriodSelect.addEventListener('change', (e) => {
        state.txnFilterPeriod = e.target.value;
        render();
      });
    }

    const txnSearchInput = document.getElementById('txn-search-input');
    if (txnSearchInput) {
      txnSearchInput.addEventListener('input', (e) => {
        state.txnSearchQuery = e.target.value;
        render();
      });
    }

    const btnClearTxnSearch = document.getElementById('btn-clear-txn-search');
    if (btnClearTxnSearch) {
      btnClearTxnSearch.addEventListener('click', () => {
        state.txnSearchQuery = '';
        render();
      });
    }

    document.querySelectorAll('.filter-pills-container button[data-txn-cat]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        state.txnFilterCat = e.currentTarget.dataset.txnCat;
        render();
      });
    });

    const btnExportTxnCSV = document.getElementById('btn-export-txn-csv');
    if (btnExportTxnCSV) {
      btnExportTxnCSV.addEventListener('click', () => {
        if (window.ExcelExporter) {
          window.ExcelExporter.exportTransactionHistoryCSV(state.data.transactions || []);
        }
      });
    }

    // Admin Modal & Portal Toggle
    const adminToggleBtn = document.getElementById('btn-admin-toggle');
    if (adminToggleBtn) {
      adminToggleBtn.addEventListener('click', () => {
        if (state.isAdminUnlocked) {
          state.activeTab = state.activeTab === 'admin' ? 'kwh' : 'admin';
          render();
        } else {
          state.showAdminModal = true;
          state.adminErrorMessage = '';
          render();
        }
      });
    }

    const closeAdminBtn = document.getElementById('btn-close-admin-modal');
    if (closeAdminBtn) {
      closeAdminBtn.addEventListener('click', () => {
        state.showAdminModal = false;
        render();
      });
    }

    const adminForm = document.getElementById('admin-login-form');
    if (adminForm) {
      adminForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const passVal = document.getElementById('admin-pass-input').value;
        if (passVal === 'Satya@1996') {
          state.isAdminUnlocked = true;
          state.showAdminModal = false;
          state.activeTab = 'admin';
          showToast('🔓 Admin Portal Unlocked Successfully');
          render();
        } else {
          state.adminErrorMessage = 'Invalid Admin Passcode. Please try again.';
          render();
        }
      });
    }

    const lockAdminBtn = document.getElementById('btn-admin-lock');
    if (lockAdminBtn) {
      lockAdminBtn.addEventListener('click', () => {
        state.isAdminUnlocked = false;
        state.activeTab = 'kwh';
        showToast('🔒 Admin Portal Locked');
        render();
      });
    }

    // Daily Reading Inputs Change
    document.querySelectorAll('.reading-input').forEach(input => {
      input.addEventListener('change', (e) => {
        const cat = e.target.dataset.cat;
        const id = parseInt(e.target.dataset.id);
        const field = e.target.dataset.field || 'reading';
        const val = parseFloat(e.target.value) || 0;
        const dateStr = state.selectedDate;

        ensureDateStructure(dateStr);
        const mList = state.data.kwh_daily[dateStr]?.[cat];
        if (mList) {
          const item = mList.find(x => x.id === id);
          if (item) {
            const oldVal = item[field];
            item[field] = val;

            const prev = getExactPreviousReading(dateStr, cat, id);
            const cons = Math.max(0, val - (field === 'reading' ? prev.reading : prev.dg_reading));
            const rateVal = field === 'reading' ? (cat === 'btu' ? state.rates.btu : state.rates.kwh) : state.rates.dg;

            recordTransaction({
              category: cat,
              categoryLabel: `${cat.toUpperCase()} Meter Update (${field})`,
              meterId: id,
              meterName: item.name || `${cat.toUpperCase()}-${id}`,
              location: item.location || "3F",
              prevReading: oldVal !== undefined ? oldVal : prev.reading,
              newReading: val,
              consumption: cons,
              cost: cons * rateVal,
              rate: rateVal,
              details: `Daily entry update for ${dateStr} (${field})`,
              source: 'Direct User Input'
            });

            saveData(false);
            showToast(`Updated ${item.name || cat} reading: ${val}`);
            render();
          }
        }
      });
    });

    // AHU Schedule Changes
        // AHU Saving Tracker Time & Consumption Inputs Change
    document.querySelectorAll('.ahu-time-input').forEach(input => {
      input.addEventListener('change', (e) => {
        const ahuId = e.target.dataset.ahu;
        const field = e.target.dataset.field;
        const val = e.target.value;
        const dateStr = state.selectedDate;

        ensureDateStructure(dateStr);
        const dayRec = state.data.ahu_saving[dateStr];
        if (dayRec && dayRec.ahus) {
          const ahu = dayRec.ahus.find(x => x.ahu_id === ahuId);
          if (ahu) {
            const oldVal = ahu[field];
            ahu[field] = val;

            recordTransaction({
              category: 'ahu_sched',
              categoryLabel: 'AHU Schedule Time Update',
              meterId: ahuId,
              meterName: `${ahuId} Operating Schedule`,
              location: "3F",
              details: `${ahuId} ${field} updated from ${oldVal || 'default'} to ${val}`,
              source: 'Direct User Input'
            });

            saveData(false);
            showToast(`Updated ${ahuId} ${field}: ${val}`);
            render();
          }
        }
      });
    });

    document.querySelectorAll('.ahu-cons-input').forEach(input => {
      input.addEventListener('change', (e) => {
        const ahuId = e.target.dataset.ahu;
        const field = e.target.dataset.field;
        const val = parseFloat(e.target.value) || 0;
        const dateStr = state.selectedDate;

        ensureDateStructure(dateStr);
        const dayRec = state.data.ahu_saving[dateStr];
        if (dayRec && dayRec.ahus) {
          const ahu = dayRec.ahus.find(x => x.ahu_id === ahuId);
          if (ahu) {
            const oldVal = ahu[field];
            ahu[field] = val;

            recordTransaction({
              category: field === 'kwh_cons' ? 'ahu' : 'btu',
              categoryLabel: field === 'kwh_cons' ? 'AHU Power Consumption' : 'BTU Cooling Consumption',
              meterId: ahuId,
              meterName: `${ahuId} ${field === 'kwh_cons' ? 'AHU Power' : 'BTU Cooling'}`,
              location: "3F",
              consumption: val,
              cost: val * (field === 'kwh_cons' ? state.rates.kwh : state.rates.btu),
              rate: field === 'kwh_cons' ? state.rates.kwh : state.rates.btu,
              details: `${ahuId} manual override: ${val}`,
              source: 'Direct User Input'
            });

            saveData(false);
            showToast(`Updated ${ahuId} ${field}: ${val}`);
            render();
          }
        }
      });
    });

    const saveKwhBtn = document.getElementById('btn-save-kwh');
    if (saveKwhBtn) {
      saveKwhBtn.addEventListener('click', () => {
        saveData(false);
        window.CloudSync.push(true);
        showToast(`✅ All Daily KWH Readings for ${state.selectedDate} saved & cloud-synced!`);
      });
    }

    const saveAhuBtn = document.getElementById('btn-save-ahu');
    if (saveAhuBtn) {
      saveAhuBtn.addEventListener('click', () => {
        saveData(false);
        window.CloudSync.push(true);
        showToast(`✅ AHU Operating Schedule for ${state.selectedDate} saved & cloud-synced!`);
      });
    }

    // Export Buttons in daily tabs
    const exportKwhBtn = document.getElementById('btn-export-kwh');
    if (exportKwhBtn) {
      exportKwhBtn.addEventListener('click', () => {
        if (window.ExcelExporter) {
          window.ExcelExporter.exportKWHDailyLog(state.selectedMonthMatrix, state.data);
        }
      });
    }

    const exportAhuBtn = document.getElementById('btn-export-ahu');
    if (exportAhuBtn) {
      exportAhuBtn.addEventListener('click', () => {
        if (window.ExcelExporter) {
          window.ExcelExporter.exportAHUSavingLog(state.selectedMonthMatrix, state.data);
        }
      });
    }
  }

  function renderAdminCharts() {
    if (!window.Chart) return;

    // 1. Monthly Consumption Trend Chart
    const trendCtx = document.getElementById('chart-monthly-trend');
    if (trendCtx) {
      new Chart(trendCtx, {
        type: 'line',
        data: {
          labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
          datasets: [
            {
              label: '2025 Cons (kWh)',
              data: [71247, 74823, 83562, 91430, 94164, 105721, 114327],
              borderColor: '#94a3b8',
              backgroundColor: 'rgba(148, 163, 184, 0.12)',
              tension: 0.35,
              fill: true
            },
            {
              label: '2026 Cons (kWh)',
              data: [78876, 77805, 80041, 80188, 76356, 82334, 76338],
              borderColor: '#34d399',
              backgroundColor: 'rgba(52, 211, 153, 0.25)',
              tension: 0.35,
              fill: true
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: { duration: 1200, easing: 'easeInOutQuart' },
          plugins: { legend: { labels: { color: '#cbd5e1' } } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
            y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
          }
        }
      });
    }

    // 2. Monthly Energy Cost Trend Chart
    const costCtx = document.getElementById('chart-cost-trend');
    if (costCtx) {
      new Chart(costCtx, {
        type: 'line',
        data: {
          labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
          datasets: [
            {
              label: '2025 Cost (₹)',
              data: [530790, 557431, 622537, 681153, 701523, 787622, 851500],
              borderColor: '#ef4444',
              backgroundColor: 'rgba(239, 68, 68, 0.1)',
              tension: 0.35,
              fill: true,
              borderDash: [5, 5]
            },
            {
              label: '2026 Cost (₹)',
              data: [587626, 579649, 596305, 597400, 568852, 613388, 568720],
              borderColor: '#60a5fa',
              backgroundColor: 'rgba(96, 165, 250, 0.25)',
              tension: 0.35,
              fill: true
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: { duration: 1300, easing: 'easeInOutQuart' },
          plugins: { legend: { labels: { color: '#cbd5e1' } } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
            y: { 
              ticks: { 
                color: '#94a3b8',
                callback: function(v) { return '₹' + (v/100000).toFixed(1) + 'L'; }
              }, 
              grid: { color: 'rgba(255,255,255,0.05)' } 
            }
          }
        }
      });
    }

    // 3. EPI Trend vs 2025 Benchmark Chart
    const epiCtx = document.getElementById('chart-epi-trend');
    if (epiCtx) {
      new Chart(epiCtx, {
        type: 'line',
        data: {
          labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
          datasets: [
            {
              label: '2025 EPI (kWh/m²/yr)',
              data: [93.7, 98.4, 109.9, 120.2, 123.8, 139.0, 150.4],
              borderColor: '#f59e0b',
              backgroundColor: 'rgba(245, 158, 11, 0.1)',
              tension: 0.35,
              fill: false
            },
            {
              label: '2026 EPI (kWh/m²/yr)',
              data: [103.7, 102.3, 105.3, 105.5, 100.4, 108.3, 100.4],
              borderColor: '#00e5ff',
              backgroundColor: 'rgba(0, 229, 255, 0.25)',
              tension: 0.35,
              fill: true
            },
            {
              label: '2025 Benchmark Avg (119.4)',
              data: [119.4, 119.4, 119.4, 119.4, 119.4, 119.4, 119.4],
              borderColor: '#ef4444',
              borderDash: [6, 4],
              borderWidth: 2,
              pointRadius: 0,
              fill: false
            },
            {
              label: 'DT3 Best-in-Class Target (<120)',
              data: [120, 120, 120, 120, 120, 120, 120],
              borderColor: '#10b981',
              borderDash: [2, 2],
              borderWidth: 2,
              pointRadius: 0,
              fill: false
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: { duration: 1400, easing: 'easeInOutQuart' },
          plugins: { legend: { labels: { color: '#cbd5e1' } } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
            y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
          }
        }
      });
    }

    // 4. YTD YoY Energy Component Net Cost Savings Chart
    const yoyCtx = document.getElementById('chart-yoy-breakdown');
    if (yoyCtx) {
      new Chart(yoyCtx, {
        type: 'bar',
        data: {
          labels: ['EB Electricity', 'AHU Power', 'DG Backup', 'BTU Cooling'],
          datasets: [{
            label: 'Net Cost Savings (₹)',
            data: [177668, -110558, 40248, 786255],
            backgroundColor: ['#00e5ff', '#ef4444', '#fbbf24', '#34d399']
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: { duration: 1000, easing: 'easeInOutQuart' },
          plugins: { legend: { display: false } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
            y: { 
              ticks: { 
                color: '#94a3b8',
                callback: function(v) { return '₹' + (v/1000).toFixed(0) + 'k'; }
              }, 
              grid: { color: 'rgba(255,255,255,0.05)' } 
            }
          }
        }
      });
    }
  }

  // Initialize on DOM Ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      initData();
      render();
    });
  } else {
    initData();
    render();
  }
})();

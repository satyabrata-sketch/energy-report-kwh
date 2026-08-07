/**
 * AHU & Daily KWH Energy Tracker - Main Application Script
 * Render exact Excel Dashboard and Summary Sheet Formats
 */

(function () {
  // Global App State
  const state = {
    activeTab: 'kwh', // 'kwh' | 'ahu' | 'month' | 'admin'
    selectedDate: '2026-08-07',
    selectedMonthMatrix: 'Aug-25', // Month filter key for Monthly Matrix view
    kwhFilter: 'all', // 'all' | 'eb' | 'ahu' | 'btu'
    fontScale: 1, // 0.85, 1, 1.15, 1.3
    isAdminUnlocked: false,
    showAdminModal: false,
    adminErrorMessage: '',
    rates: {
      kwh: 7.45,
      btu: 4.30,
      dg: 33.85
    },

    // Master Tracker Data
    data: {
      ahu_saving: {},
      kwh_daily: {},
      month_baselines: {},
      summary_matrix: { months: [], consumption: {}, cost: {} }
    }
  };

  // Initialize Data from LocalStorage or Seed Data
  function initData() {
    const stored = localStorage.getItem('kwh_ahu_tracker_data_v2');
    if (stored) {
      try {
        state.data = JSON.parse(stored);
      } catch (e) {
        console.error("Failed to parse stored data:", e);
        state.data = window.SEED_DATA || { ahu_saving: {}, kwh_daily: {}, month_baselines: {}, summary_matrix: {} };
      }
    } else {
      state.data = window.SEED_DATA || { ahu_saving: {}, kwh_daily: {}, month_baselines: {}, summary_matrix: {} };
    }

    const todayStr = new Date().toISOString().slice(0, 10);
    state.selectedDate = todayStr;
    ensureDateStructure(todayStr);
  }

  function saveData() {
    localStorage.setItem('kwh_ahu_tracker_data_v2', JSON.stringify(state.data));
  }

  function getPreviousDateStr(dateStr) {
    const d = new Date(dateStr);
    d.setDate(d.getDate() - 1);
    return d.toISOString().slice(0, 10);
  }

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

  function getAutoKWHConsumptionsForDate(dateStr) {
    const currKWH = state.data.kwh_daily[dateStr] || {};
    const ahuCons = {};
    const btuCons = {};

    [1, 2, 3, 4].forEach(id => {
      const prevInfoA = getExactPreviousReading(dateStr, 'ahu', id);
      const currAHU = currKWH.ahu?.find(a => a.id === id) || {};
      const currValA = currAHU.reading !== undefined ? Number(currAHU.reading) : prevInfoA.reading;
      ahuCons[`AHU${id}`] = Math.max(0, currValA - prevInfoA.reading);

      const prevInfoB = getExactPreviousReading(dateStr, 'btu', id);
      const currBTU = currKWH.btu?.find(b => b.id === id) || {};
      const currValB = currBTU.reading !== undefined ? Number(currBTU.reading) : prevInfoB.reading;
      btuCons[`AHU${id}`] = Math.max(0, currValB - prevInfoB.reading);
    });

    return { ahuCons, btuCons };
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

  function calculateAHUSaving(dayRec) {
    const dayStr = dayRec.day;
    const isSun = dayStr === 'Sun';
    const isSat = dayStr === 'Sat';

    const stdSched = isSun ? 0 : (isSat ? 6 : 12);
    const stdOffHour = isSat ? 14 : 20;

    const autoCons = getAutoKWHConsumptionsForDate(dayRec.date);

    let totSched = 0, totRun = 0, totSaved = 0;
    let totAHUKwh = 0, totAHUCost = 0, totAHUSaved = 0;
    let totBTUKwh = 0, totBTUCost = 0, totBTUSaved = 0;
    let totCombSaved = 0, totCombFull = 0;

    const ahusComputed = (dayRec.ahus || []).map((ahu) => {
      const onStr = ahu.on_time || "OFF";
      const offStr = ahu.off_time || "OFF";

      let runHrs = 0;
      let savedHrs = 0;

      if (onStr !== "OFF" && offStr !== "OFF" && onStr && offStr) {
        const [onH, onM] = onStr.split(':').map(Number);
        const [offH, offM] = offStr.split(':').map(Number);
        const onVal = onH + (onM || 0) / 60;
        const offVal = offH + (offM || 0) / 60;

        runHrs = Math.max(0, offVal - onVal);

        if (!isSun) {
          savedHrs = Math.max(0, stdOffHour - offVal);
        }
      }

      const ahuKwh = ahu.kwh_cons !== undefined && ahu.kwh_cons !== null
        ? Number(ahu.kwh_cons) 
        : (autoCons.ahuCons[ahu.ahu_id] || 0);

      const kwhRate = Number(dayRec.kwh_rate || state.rates.kwh);
      const ahuCost = ahuKwh * kwhRate;
      const ahuRatePerHr = runHrs > 0 ? (ahuCost / runHrs) : 0;
      const ahuFullCost8pm = ahuCost + (savedHrs * ahuRatePerHr);
      const ahuCostSaved = savedHrs * ahuRatePerHr;

      const btuUnits = ahu.btu_cons !== undefined && ahu.btu_cons !== null
        ? Number(ahu.btu_cons)
        : (autoCons.btuCons[ahu.ahu_id] || 0);

      const btuRate = Number(dayRec.btu_rate || state.rates.btu);
      const btuCost = btuUnits * btuRate;
      const btuRatePerHr = runHrs > 0 ? (btuCost / runHrs) : 0;
      const btuFullCost8pm = btuCost + (savedHrs * btuRatePerHr);
      const btuCostSaved = savedHrs * btuRatePerHr;

      const totCostSavedToday = ahuCostSaved + btuCostSaved;
      const totFullCost8pmToday = ahuFullCost8pm + btuFullCost8pm;

      totSched += stdSched;
      totRun += runHrs;
      totSaved += savedHrs;

      totAHUKwh += ahuKwh;
      totAHUCost += ahuCost;
      totAHUSaved += ahuCostSaved;

      totBTUKwh += btuUnits;
      totBTUCost += btuCost;
      totBTUSaved += btuCostSaved;

      totCombSaved += totCostSavedToday;
      totCombFull += totFullCost8pmToday;

      return {
        ...ahu,
        kwh_cons: ahuKwh,
        btu_cons: btuUnits,
        sched_hrs: stdSched,
        actual_run_hrs: runHrs,
        saved_hrs: savedHrs,
        ahu_cost: ahuCost,
        ahu_rate_per_hr: ahuRatePerHr,
        ahu_full_cost_8pm: ahuFullCost8pm,
        ahu_cost_saved: ahuCostSaved,
        btu_cost: btuCost,
        btu_rate_per_hr: btuRatePerHr,
        btu_full_cost_8pm: btuFullCost8pm,
        btu_cost_saved: btuCostSaved,
        tot_cost_saved_today: totCostSavedToday,
        tot_full_cost_8pm_today: totFullCost8pmToday
      };
    });

    const savingsPct = totCombFull > 0 ? (totCombSaved / totCombFull) : 0;

    return {
      ...dayRec,
      ahus: ahusComputed,
      tot_sched: totSched,
      tot_run: totRun,
      tot_saved: totSaved,
      tot_kwh: totAHUKwh,
      tot_ahu_cost: totAHUCost,
      tot_ahu_cost_saved: totAHUSaved,
      tot_btu_kwh: totBTUKwh,
      tot_btu_cost: totBTUCost,
      tot_btu_cost_saved: totBTUSaved,
      tot_comb_cost_saved: totCombSaved,
      tot_comb_full_cost: totCombFull,
      savings_pct: savingsPct
    };
  }

  function showToast(msg) {
    const existing = document.querySelector('.toast-msg');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.className = 'toast-msg';
    toast.innerHTML = `<i data-lucide="check-circle"></i> ${msg}`;
    document.body.appendChild(toast);
    if (window.lucide) lucide.createIcons();

    setTimeout(() => {
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 300);
    }, 2500);
  }

  function render() {
    ensureDateStructure(state.selectedDate);

    document.documentElement.style.setProperty('--font-scale', state.fontScale);

    const appContainer = document.getElementById('app');
    if (!appContainer) return;

    appContainer.innerHTML = `
      ${renderTopHeader()}
      ${renderNavTabs()}
      ${renderUserControlBarOnlyDate()}
      ${renderTabContent()}
      ${state.activeTab === 'month' ? renderFullDT3SummaryMatrixSection() : renderMonthMatrixSectionBelow()}
      ${state.showAdminModal ? renderAdminAuthModal() : ''}
    `;

    if (window.lucide) lucide.createIcons();

    attachEventListeners();

    if (state.activeTab === 'admin' && state.isAdminUnlocked) {
      setTimeout(renderAdminCharts, 50);
    }
  }

  function renderTopHeader() {
    return `
      <header class="top-header">
        <div class="brand-container">
          <div class="brand-logo"><i data-lucide="zap"></i></div>
          <div class="brand-title">
            <h1>AHU & Energy (KWH) Tracker</h1>
            <div class="brand-subtitle">CBRE | Operational AHU Saving & Daily KWH Consumption Manager</div>
          </div>
        </div>

        <div class="top-utilities">
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

  function renderNavTabs() {
    return `
      <nav class="nav-tabs-bar">
        <button class="nav-tab-btn kwh-tab ${state.activeTab === 'kwh' ? 'active' : ''}" data-tab="kwh">
          <i data-lucide="activity"></i> Daily KWH Entry (EB, AHU, BTU)
        </button>
        <button class="nav-tab-btn ahu-tab ${state.activeTab === 'ahu' ? 'active' : ''}" data-tab="ahu">
          <i data-lucide="fan"></i> AHU Saving Tracker & BTU Cost
        </button>
        <button class="nav-tab-btn ${state.activeTab === 'month' ? 'active' : ''}" data-tab="month">
          <i data-lucide="calendar"></i> Monthly Excel Matrix View
        </button>
        <button class="nav-tab-btn admin-tab ${state.activeTab === 'admin' ? 'active' : ''}" data-tab="admin">
          <i data-lucide="shield-check"></i> Executive Dashboards & Admin
        </button>
      </nav>
    `;
  }

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
          <div class="date-input-wrap" style="background: linear-gradient(135deg, rgba(0,229,255,0.1), rgba(59,130,246,0.1)); border: 1px solid #00e5ff;">
            <i data-lucide="calendar" style="color:#00e5ff; font-size:1.2rem;"></i>
            <label style="font-size:0.85rem; color:#fff; font-weight:700;">Select Date from Calendar:</label>
            <input type="date" id="selected-date-picker" value="${state.selectedDate}" style="font-size:1rem; font-weight:800; color:#00e5ff; cursor:pointer;">
            <span style="font-size:0.75rem; color:var(--text-muted); margin-left:0.4rem;">(${dateFormattedStr})</span>
          </div>
        </div>

        ${kwhPills}
      </div>
    `;
  }

  function renderTabContent() {
    if (state.activeTab === 'kwh') {
      return renderKWHDailyLogContent();
    } else if (state.activeTab === 'ahu') {
      return renderAHUSavingLogContent();
    } else if (state.activeTab === 'month') {
      return '';
    } else if (state.activeTab === 'admin') {
      return renderAdminPortalContent();
    }
    return '';
  }

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
      date: dateStr, day: 'Mon', ahus: [], kwh_rate: state.rates.kwh, btu_rate: state.rates.btu
    };

    const computed = calculateAHUSaving(rawDayRec);

    const cardsHtml = (computed.ahus || []).map((ahu) => {
      return `
        <div class="meter-card ahu-theme">
          <div class="meter-card-header">
            <div class="meter-title-wrap">
              <div class="meter-icon-badge"><i data-lucide="fan"></i></div>
              <div>
                <div class="meter-title">${ahu.ahu_id} Saving & BTU Cooling Log</div>
                <div class="meter-location-tag">Standard Sched: ${ahu.sched_hrs} hrs</div>
              </div>
            </div>
          </div>

          <div class="reading-inputs-grid">
            <div class="input-block">
              <label>ON Time (hh:mm)</label>
              <input type="time" class="input-field ahu-time-input" 
                     data-ahu="${ahu.ahu_id}" data-field="on_time" value="${ahu.on_time || '07:12'}">
            </div>
            <div class="input-block">
              <label style="color:#ef4444;">OFF Time (hh:mm)</label>
              <input type="time" class="input-field ahu-time-input" 
                     data-ahu="${ahu.ahu_id}" data-field="off_time" value="${ahu.off_time || '20:00'}">
            </div>
          </div>

          <div class="reading-inputs-grid" style="margin-bottom:1rem;">
            <div class="input-block">
              <label style="color:#10b981;">AHU Power Consumed (kWh) <span class="auto-tag">(Auto/Edit)</span></label>
              <input type="number" step="any" class="input-field ahu-cons-input" 
                     data-ahu="${ahu.ahu_id}" data-field="kwh_cons" value="${ahu.kwh_cons}">
            </div>
            <div class="input-block">
              <label style="color:#f59e0b;">BTU Cooling Consumed (kWh) <span class="auto-tag">(Auto/Edit)</span></label>
              <input type="number" step="any" class="input-field ahu-cons-input" 
                     data-ahu="${ahu.ahu_id}" data-field="btu_cons" value="${ahu.btu_cons}">
            </div>
          </div>

          <div class="results-strip" style="flex-wrap:wrap; gap:0.5rem; background:rgba(0,0,0,0.4);">
            <div class="res-item">
              <span class="res-label">Actual Run Hrs</span>
              <span class="res-value">${ahu.actual_run_hrs.toFixed(1)} hrs</span>
            </div>
            <div class="res-item">
              <span class="res-label">Saved Hrs Today</span>
              <span class="res-value highlight" style="color:var(--ahu-color);">${ahu.saved_hrs.toFixed(1)} hrs</span>
            </div>
            <div class="res-item">
              <span class="res-label">AHU Cost Saved</span>
              <span class="res-value" style="color:#34d399;">₹${ahu.ahu_cost_saved.toFixed(0)}</span>
            </div>
            <div class="res-item">
              <span class="res-label">BTU Cost Saved</span>
              <span class="res-value" style="color:#fbbf24;">₹${ahu.btu_cost_saved.toFixed(0)}</span>
            </div>
            <div class="res-item" style="text-align:right; width:100%; border-top:1px dashed var(--border-color); padding-top:0.4rem; margin-top:0.2rem;">
              <span class="res-label">Total Cost Saved Today (AHU + BTU)</span>
              <span class="res-value" style="color:#00e5ff; font-size:1.1rem;">₹${ahu.tot_cost_saved_today.toFixed(0)}</span>
            </div>
          </div>
        </div>
      `;
    }).join('');

    return `
      <div class="category-summary-banner">
        <div class="summary-stat-box">
          <div class="stat-label">Total Scheduled Hours</div>
          <div class="stat-val">${computed.tot_sched} hrs</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total Actual Run Hours</div>
          <div class="stat-val">${computed.tot_run.toFixed(1)} hrs</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total Saved Hours Today</div>
          <div class="stat-val" style="color:var(--ahu-color);">${computed.tot_saved.toFixed(1)} hrs</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">AHU Power Saved (₹)</div>
          <div class="stat-val" style="color:#34d399;">₹${computed.tot_ahu_cost_saved.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">BTU Cooling Saved (₹)</div>
          <div class="stat-val" style="color:#fbbf24;">₹${computed.tot_btu_cost_saved.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Total Combined Savings (₹)</div>
          <div class="stat-val" style="color:#00e5ff;">₹${computed.tot_comb_cost_saved.toFixed(2)}</div>
        </div>
        <div class="summary-stat-box">
          <div class="stat-label">Combined Savings %</div>
          <div class="stat-val" style="color:#60a5fa;">${(computed.savings_pct * 100).toFixed(1)}%</div>
        </div>
      </div>

      <div class="section-header-title ahu-header">
        <h2><i data-lucide="fan"></i> Operational AHU & BTU Cooling Cost Savings Engine</h2>
        <span style="font-size:0.8rem; font-weight:normal;">Auto-synced from Daily KWH Tracker + Manual Edit Support</span>
      </div>

      <div class="section-grid">${cardsHtml}</div>

      <div class="action-bar-floating">
        <button class="btn-primary-save" id="btn-save-ahu">
          <i data-lucide="save"></i> Save & Commit AHU & BTU Savings Log
        </button>
      </div>
    `;
  }

  // Full DT-3 Summary Matrix Sheet Format (Matching Excel summary Sheet)
  function renderFullDT3SummaryMatrixSection() {
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
          <span style="font-size:0.8rem;">Historical & Current Monthly Energy Consumptions (2024 to 2026)</span>
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

  function renderMonthMatrixSectionBelow() {
    const selectedMonth = state.selectedMonthMatrix;
    const sortedDates = Object.keys(state.data.kwh_daily || {}).sort();

    const dateList = sortedDates.filter(dStr => {
      const dt = new Date(dStr);
      const mKey = dt.toLocaleString('default', { month: 'short', year: '2-digit' }).replace(' ', '-');
      const mLong = dt.toLocaleString('default', { month: 'long', year: 'numeric' });
      return mKey.toLowerCase() === selectedMonth.toLowerCase() || 
             mKey.replace('-', ' ').toLowerCase() === selectedMonth.toLowerCase() ||
             mLong.toLowerCase().includes(selectedMonth.toLowerCase().replace('-', ' '));
    });

    const monthOptions = [
      { key: "Aug-26", label: "August 2026" },
      { key: "July-2026", label: "July 2026" },
      { key: "Jun-2026", label: "June 2026" },
      { key: "May-2026", label: "May 2026" },
      { key: "Aug-25", label: "August 2025" },
      { key: "July-2025", label: "July 2025" },
      { key: "March-25", label: "March 2025" },
      { key: "Feb-2025", label: "February 2025" },
      { key: "Jan-2025", label: "January 2025" }
    ];

    const selectOptionsHtml = monthOptions.map(m => `
      <option value="${m.key}" ${state.selectedMonthMatrix === m.key ? 'selected' : ''}>${m.label}</option>
    `).join('');

    const ebMaster = [
      { id: 1, loc: "3F", name: "SL No 540430038840 (EB Unit)" },
      { id: 2, loc: "3F", name: "SL No 540430038845 (EB Unit)" },
      { id: 3, loc: "3F", name: "SL No 540430038844 (EB Unit)" },
      { id: 4, loc: "3F", name: "SL No 540430038821 (EB Unit)" },
      { id: 5, loc: "3F", name: "SL No 540430038849 (EB Unit)" },
      { id: 6, loc: "3F", name: "SL No 540430038646 (EB Unit)" },
      { id: 7, loc: "3F", name: "SL No 540430038843 (EB Unit)" },
      { id: 8, loc: "3F", name: "SL No 540430038848 (EB Unit)" }
    ];

    const ahuMaster = [
      { id: 1, loc: "3F", name: "AHU1" },
      { id: 2, loc: "3F", name: "AHU2" },
      { id: 3, loc: "3F", name: "AHU3" },
      { id: 4, loc: "3F", name: "AHU4" }
    ];

    const btuMaster = [
      { id: 1, loc: "3F", name: "AHU1 - BTU" },
      { id: 2, loc: "3F", name: "AHU2 - BTU" },
      { id: 3, loc: "3F", name: "AHU3 - BTU" },
      { id: 4, loc: "3F", name: "AHU4 - BTU" }
    ];

    let readingsRowsHtml = '';

    ebMaster.forEach(m => {
      const prevInfo = getExactPreviousReading(dateList[0] || state.selectedDate, 'eb', m.id);
      const prevVal = prevInfo.reading;
      const prevDGVal = prevInfo.dg_reading;

      let ebCells = dateList.map(d => {
        const val = state.data.kwh_daily[d]?.eb?.find(e => e.id === m.id)?.reading;
        return `<td>${val !== undefined ? val : ''}</td>`;
      }).join('');

      let dgCells = dateList.map(d => {
        const val = state.data.kwh_daily[d]?.eb?.find(e => e.id === m.id)?.dg_reading;
        return `<td style="color:#fbbf24;">${val !== undefined ? val : ''}</td>`;
      }).join('');

      readingsRowsHtml += `
        <tr>
          <td><strong>${m.id}</strong></td>
          <td>${m.loc}</td>
          <td><strong style="color:var(--eb-color);">${m.name}</strong></td>
          <td><strong>${prevVal}</strong></td>
          ${ebCells}
        </tr>
        <tr style="background:rgba(251, 191, 36, 0.03);">
          <td></td><td></td>
          <td style="color:#fbbf24; font-weight:bold;">↳ DG Backup Reading</td>
          <td style="color:#fbbf24;">${prevDGVal}</td>
          ${dgCells}
        </tr>
      `;
    });

    ahuMaster.forEach(m => {
      const prevInfo = getExactPreviousReading(dateList[0] || state.selectedDate, 'ahu', m.id);
      const prevVal = prevInfo.reading;
      const prevDGVal = prevInfo.dg_reading;

      let ahuCells = dateList.map(d => {
        const val = state.data.kwh_daily[d]?.ahu?.find(a => a.id === m.id)?.reading;
        return `<td>${val !== undefined ? val : ''}</td>`;
      }).join('');

      let dgCells = dateList.map(d => {
        const val = state.data.kwh_daily[d]?.ahu?.find(a => a.id === m.id)?.dg_reading;
        return `<td style="color:#fbbf24;">${val !== undefined ? val : ''}</td>`;
      }).join('');

      readingsRowsHtml += `
        <tr style="border-top: 1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td>
          <td>${m.loc}</td>
          <td><strong style="color:var(--ahu-color);">${m.name}</strong></td>
          <td><strong>${prevVal}</strong></td>
          ${ahuCells}
        </tr>
        <tr style="background:rgba(251, 191, 36, 0.03);">
          <td></td><td></td>
          <td style="color:#fbbf24; font-weight:bold;">↳ DG Backup Reading</td>
          <td style="color:#fbbf24;">${prevDGVal}</td>
          ${dgCells}
        </tr>
      `;
    });

    btuMaster.forEach(m => {
      const prevInfo = getExactPreviousReading(dateList[0] || state.selectedDate, 'btu', m.id);
      const prevVal = prevInfo.reading;

      let btuCells = dateList.map(d => {
        const val = state.data.kwh_daily[d]?.btu?.find(b => b.id === m.id)?.reading;
        return `<td>${val !== undefined ? val : ''}</td>`;
      }).join('');

      readingsRowsHtml += `
        <tr style="border-top: 1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td>
          <td>${m.loc}</td>
          <td><strong style="color:var(--btu-color);">${m.name}</strong></td>
          <td><strong>${prevVal}</strong></td>
          ${btuCells}
        </tr>
      `;
    });

    let consRowsHtml = '';
    const totEBByDate = {};
    const totDGByDate = {};
    const totAHUByDate = {};
    const totAHUDGByDate = {};
    const totBTUByDate = {};

    dateList.forEach(dStr => {
      totEBByDate[dStr] = 0;
      totDGByDate[dStr] = 0;
      totAHUByDate[dStr] = 0;
      totAHUDGByDate[dStr] = 0;
      totBTUByDate[dStr] = 0;
    });

    ebMaster.forEach(m => {
      let ebConsCells = '';
      let dgConsCells = '';

      for (let idx = 0; idx < dateList.length; idx++) {
        const currDate = dateList[idx];
        const currData = state.data.kwh_daily[currDate]?.eb?.find(e => e.id === m.id) || {};
        const prevInfo = getExactPreviousReading(currDate, 'eb', m.id);

        const ebCons = (currData.reading && prevInfo.reading) ? Math.max(0, currData.reading - prevInfo.reading) : 0;
        const dgCons = (currData.dg_reading && prevInfo.dg_reading) ? Math.max(0, currData.dg_reading - prevInfo.dg_reading) : 0;

        totEBByDate[currDate] += ebCons;
        totDGByDate[currDate] += dgCons;

        ebConsCells += `<td>${ebCons}</td>`;
        dgConsCells += `<td style="color:#fbbf24;">${dgCons}</td>`;
      }

      consRowsHtml += `
        <tr>
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--eb-color);">${m.name}</strong></td>
          ${ebConsCells}
        </tr>
        <tr style="background:rgba(251, 191, 36, 0.03);">
          <td></td><td></td><td style="color:#fbbf24;">↳ DG Backup Cons (kWh)</td>
          ${dgConsCells}
        </tr>
      `;
    });

    const totEBCells = dateList.map(d => `<td><strong>${totEBByDate[d]}</strong></td>`).join('');
    const totDGCells = dateList.map(d => `<td style="color:#fbbf24;"><strong>${totDGByDate[d]}</strong></td>`).join('');
    const totCumCells = dateList.map(d => `<td style="color:#00e5ff;"><strong>${totEBByDate[d] + totDGByDate[d]}</strong></td>`).join('');

    consRowsHtml += `
      <tr class="total-row"><td></td><td></td><td>Total Power - EB Unit</td>${totEBCells}</tr>
      <tr class="total-row"><td></td><td></td><td style="color:#fbbf24;">Total Power - DG Backup</td>${totDGCells}</tr>
      <tr class="total-row" style="background:rgba(0, 229, 255, 0.15);"><td></td><td></td><td style="color:#00e5ff;">Total Power Cumulative</td>${totCumCells}</tr>
    `;

    ahuMaster.forEach(m => {
      let ahuConsCells = '';
      let dgConsCells = '';

      for (let idx = 0; idx < dateList.length; idx++) {
        const currDate = dateList[idx];
        const currData = state.data.kwh_daily[currDate]?.ahu?.find(a => a.id === m.id) || {};
        const prevInfo = getExactPreviousReading(currDate, 'ahu', m.id);

        const ahuCons = (currData.reading && prevInfo.reading) ? Math.max(0, currData.reading - prevInfo.reading) : 0;
        const dgCons = (currData.dg_reading && prevInfo.dg_reading) ? Math.max(0, currData.dg_reading - prevInfo.dg_reading) : 0;

        totAHUByDate[currDate] += ahuCons;
        totAHUDGByDate[currDate] += dgCons;

        ahuConsCells += `<td>${ahuCons}</td>`;
        dgConsCells += `<td style="color:#fbbf24;">${dgCons}</td>`;
      }

      consRowsHtml += `
        <tr style="border-top: 1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--ahu-color);">${m.name}</strong></td>
          ${ahuConsCells}
        </tr>
        <tr style="background:rgba(251, 191, 36, 0.03);">
          <td></td><td></td><td style="color:#fbbf24;">↳ AHU DG Cons (kWh)</td>
          ${dgConsCells}
        </tr>
      `;
    });

    const totAHUCells = dateList.map(d => `<td style="color:var(--ahu-color);"><strong>${totAHUByDate[d]}</strong></td>`).join('');
    const totAHUDGCells = dateList.map(d => `<td style="color:#fbbf24;"><strong>${totAHUDGByDate[d]}</strong></td>`).join('');

    consRowsHtml += `
      <tr class="total-row"><td></td><td></td><td style="color:var(--ahu-color);">Total AHU Power Consumption</td>${totAHUCells}</tr>
      <tr class="total-row"><td></td><td></td><td style="color:#fbbf24;">Total AHU DG consumption</td>${totAHUDGCells}</tr>
    `;

    btuMaster.forEach(m => {
      let btuConsCells = '';

      for (let idx = 0; idx < dateList.length; idx++) {
        const currDate = dateList[idx];
        const currData = state.data.kwh_daily[currDate]?.btu?.find(b => b.id === m.id) || {};
        const prevInfo = getExactPreviousReading(currDate, 'btu', m.id);

        const btuCons = (currData.reading && prevInfo.reading) ? Math.max(0, currData.reading - prevInfo.reading) : 0;
        totBTUByDate[currDate] += btuCons;
        btuConsCells += `<td>${btuCons}</td>`;
      }

      consRowsHtml += `
        <tr style="border-top: 1px solid var(--border-highlight);">
          <td><strong>${m.id}</strong></td><td>${m.loc}</td>
          <td><strong style="color:var(--btu-color);">${m.name}</strong></td>
          ${btuConsCells}
        </tr>
      `;
    });

    const totBTUCells = dateList.map(d => `<td style="color:var(--btu-color);"><strong>${totBTUByDate[d]}</strong></td>`).join('');
    consRowsHtml += `
      <tr class="total-row" style="background:rgba(245, 158, 11, 0.15);"><td></td><td></td><td style="color:var(--btu-color);">Total BTU Consumption</td>${totBTUCells}</tr>
    `;

    const dateHeadersHtml = dateList.map(d => {
      const dt = new Date(d);
      const dayName = dt.toLocaleString('default', { weekday: 'short' });
      const dayNum = dt.getDate();
      return `<th style="text-align:center; min-width:65px;">${dayNum}<br><span style="font-size:0.65rem; color:#9ca3af;">${dayName}</span></th>`;
    }).join('');

    return `
      <div style="margin-top: 3rem; padding-top: 2rem; border-top: 2px dashed var(--border-highlight);">
        <div class="controls-card" style="margin-bottom:1.5rem; background: linear-gradient(135deg, rgba(15, 23, 42, 0.9), rgba(30, 41, 59, 0.9)); border: 1px solid #3b82f6;">
          <div style="display:flex; align-items:center; gap:1rem; flex-wrap:wrap;">
            <i data-lucide="file-spreadsheet" style="font-size:1.8rem; color:#00e5ff;"></i>
            <div>
              <h2 style="color:#fff; font-size:1.2rem; font-weight:800;">Monthly Excel Readings & Daily Energy Consumption Matrix</h2>
              <p style="color:var(--text-muted); font-size:0.8rem;">Select month below to display complete monthly readings and daily consumption grid exactly as per source Excel file.</p>
            </div>
          </div>

          <div style="display:flex; align-items:center; gap:0.8rem;">
            <label style="color:#00e5ff; font-weight:700; font-size:0.85rem;">Select Month:</label>
            <select id="month-matrix-select" class="input-field" style="width: auto; min-width: 180px;">
              ${selectOptionsHtml}
            </select>
          </div>
        </div>

        <div class="section-header-title eb-header" style="margin-top:0;">
          <h3><i data-lucide="table"></i> Section 1: Monthly Meter Readings Matrix (${selectedMonth})</h3>
          <span style="font-size:0.8rem;">EB Unit, AHU Unit & BTU Unit Daily Readings</span>
        </div>

        <div class="table-responsive-container">
          <table class="data-table">
            <thead>
              <tr>
                <th>Sr.No</th>
                <th>Location</th>
                <th>Meter Name</th>
                <th>Prev Reading</th>
                ${dateHeadersHtml}
              </tr>
            </thead>
            <tbody>
              ${readingsRowsHtml}
            </tbody>
          </table>
        </div>

        <div class="section-header-title ahu-header" style="margin-top:2rem;">
          <h3><i data-lucide="activity"></i> Section 2: Daily Energy Consumption Breakdown (${selectedMonth})</h3>
          <span style="font-size:0.8rem;">Daily Calculated Consumptions & Category Totals</span>
        </div>

        <div class="table-responsive-container">
          <table class="data-table">
            <thead>
              <tr>
                <th>Sr.No</th>
                <th>Location</th>
                <th>Meter Name</th>
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

  // 4. Render Admin Portal & Executive Dashboard Content (Matching Dashboard Sheets)
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
            <i data-lucide="shield-check" style="color:#10b981;"></i> DT-3 Executive Energy & EPI Performance Dashboard (Matching Excel)
          </h2>
          <p style="color:var(--text-muted); font-size:0.85rem;">Official Operational Performance & Excel Export Hub</p>
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
          <div class="kpi-subtext">Status: Best-in-Class / Efficient (<120 Target)</div>
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

      <!-- Section 1: EPI Performance Benchmark Standards -->
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
                <td><strong style="color:#34d399;">DT3 Operational Target (<120)</strong></td>
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

      <!-- Section 3: Interactive Animated Visuals (Consumption, Cost, EPI & Savings) -->
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

      <!-- Section 5: Key Operational Insights & Savings Justification Panel -->
      <div class="chart-container-card">
        <div class="chart-title"><i data-lucide="info"></i> 4. Key Operational Insights & Savings Justification (DT3 Facility)</div>
        <div style="padding:1rem; font-size:0.9rem; line-height:1.6; color:#d1d5db;">
          <p style="margin-bottom:0.8rem;"><strong style="color:#00e5ff;">• Operational Model Transition:</strong> Prior to May, DT3 AHUs operated on a 24*7 running schedule. From May onwards, DT3 transitioned to a 12*5.5 KAM model (12 hrs/day, 5.5 days/week), discontinuing night operations and weekend shifts.</p>
          <p style="margin-bottom:0.8rem;"><strong style="color:#10b981;">• Hours & Energy Saved:</strong> Switching to the 12*5.5 KAM model saved ~12 hrs/day on weekdays and full weekends (~434 operating hrs/mo per AHU, totaling ~1,736 AHU hrs/mo saved). Monthly AHU consumption dropped from 10,301 kWh in April to 6,266 kWh in May (saved 4,035 kWh / -39.2%), 7,682 kWh in June, and 8,308 kWh in July.</p>
          <p style="margin-bottom:0.8rem;"><strong style="color:#fbbf24;">• EPI Trend Justification (vs 2025):</strong> Annualized EPI dropped significantly YTD (May EPI: 100.42 vs 123.85 in 2025, -18.9%; June EPI: 108.28 vs 139.05 in 2025, -22.1%; July EPI: 100.4 vs 150.36 in 2025). Overall Jan–Jul average EPI improved from 119.36 to 103.7 kWh/m²/yr (-13.1%), maintaining Best-in-Class status (&lt;120 kWh/m²/yr).</p>
          <p><strong style="color:#34d399;">• YTD Cost Savings:</strong> Total Jan–July energy cost fell from ₹54.69L in 2025 to ₹45.75L in 2026, delivering ₹8,93,612 (-16.3%) cost savings YTD.</p>
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

  function attachEventListeners() {
    document.querySelectorAll('.nav-tab-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        state.activeTab = e.currentTarget.dataset.tab;
        render();
      });
    });

    document.querySelectorAll('.font-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        state.fontScale = parseFloat(e.currentTarget.dataset.font);
        render();
      });
    });

    const datePicker = document.getElementById('selected-date-picker');
    if (datePicker) {
      const handleDateChange = (e) => {
        const newDate = e.target.value;
        if (newDate && newDate !== state.selectedDate) {
          state.selectedDate = newDate;
          render();
        }
      };
      datePicker.addEventListener('change', handleDateChange);
      datePicker.addEventListener('input', handleDateChange);
    }

    const monthMatrixSelect = document.getElementById('month-matrix-select');
    if (monthMatrixSelect) {
      monthMatrixSelect.addEventListener('change', (e) => {
        state.selectedMonthMatrix = e.target.value;
        render();
      });
    }

    document.querySelectorAll('.pill-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        state.kwhFilter = e.currentTarget.dataset.filter;
        render();
      });
    });

    const adminToggleBtn = document.getElementById('btn-admin-toggle');
    if (adminToggleBtn) {
      adminToggleBtn.addEventListener('click', () => {
        if (state.isAdminUnlocked) {
          state.activeTab = 'admin';
          render();
        } else {
          state.showAdminModal = true;
          state.adminErrorMessage = '';
          render();
        }
      });
    }

    const openAdminBtn = document.getElementById('btn-open-admin-modal');
    if (openAdminBtn) {
      openAdminBtn.addEventListener('click', () => {
        state.showAdminModal = true;
        render();
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
          showToast('Admin Portal Unlocked Successfully!');
          render();
        } else {
          state.adminErrorMessage = 'Invalid Password. Please try again.';
          render();
        }
      });
    }

    const lockAdminBtn = document.getElementById('btn-admin-lock');
    if (lockAdminBtn) {
      lockAdminBtn.addEventListener('click', () => {
        state.isAdminUnlocked = false;
        state.activeTab = 'kwh';
        showToast('Admin Logged Out.');
        render();
      });
    }

    document.querySelectorAll('.kwh-input-field').forEach(input => {
      const handleReadingChange = (e) => {
        const cat = e.target.dataset.cat;
        const id = parseInt(e.target.dataset.id);
        const field = e.target.dataset.field;
        const val = parseFloat(e.target.value) || 0;

        const dateStr = state.selectedDate;
        ensureDateStructure(dateStr);
        const targetArr = state.data.kwh_daily[dateStr]?.[cat];
        if (targetArr) {
          const item = targetArr.find(x => x.id === id);
          if (item) {
            item[field] = val;
            saveData();
          }
        }
      };

      input.addEventListener('change', handleReadingChange);
      input.addEventListener('input', handleReadingChange);
    });

    document.querySelectorAll('.ahu-time-input').forEach(input => {
      const handleTimeChange = (e) => {
        const ahuId = e.target.dataset.ahu;
        const field = e.target.dataset.field;
        const val = e.target.value;

        const dateStr = state.selectedDate;
        ensureDateStructure(dateStr);
        const ahus = state.data.ahu_saving[dateStr]?.ahus;
        if (ahus) {
          const item = ahus.find(x => x.ahu_id === ahuId);
          if (item) {
            item[field] = val;
            saveData();
          }
        }
      };

      input.addEventListener('change', handleTimeChange);
      input.addEventListener('input', handleTimeChange);
    });

    document.querySelectorAll('.ahu-cons-input').forEach(input => {
      const handleConsChange = (e) => {
        const ahuId = e.target.dataset.ahu;
        const field = e.target.dataset.field;
        const val = parseFloat(e.target.value) || 0;

        const dateStr = state.selectedDate;
        ensureDateStructure(dateStr);
        const ahus = state.data.ahu_saving[dateStr]?.ahus;
        if (ahus) {
          const item = ahus.find(x => x.ahu_id === ahuId);
          if (item) {
            item[field] = val;
            saveData();
          }
        }
      };

      input.addEventListener('change', handleConsChange);
      input.addEventListener('input', handleConsChange);
    });

    const saveKwhBtn = document.getElementById('btn-save-kwh');
    if (saveKwhBtn) {
      saveKwhBtn.addEventListener('click', () => {
        saveData();
        showToast(`Daily KWH Readings for ${state.selectedDate} Saved!`);
        render();
      });
    }

    const saveAhuBtn = document.getElementById('btn-save-ahu');
    if (saveAhuBtn) {
      saveAhuBtn.addEventListener('click', () => {
        saveData();
        showToast(`AHU & BTU Savings Log for ${state.selectedDate} Saved!`);
        render();
      });
    }

    const exportKwhBtn = document.getElementById('btn-export-kwh');
    if (exportKwhBtn) {
      exportKwhBtn.addEventListener('click', () => {
        try {
          if (window.ExcelExporter) {
            window.ExcelExporter.exportKWHDailyLog(state.selectedMonthMatrix, state.data);
            showToast('KWH Daily Log Excel Download Triggered!');
          }
        } catch (err) {
          console.error('Export Error:', err);
        }
      });
    }

    const exportAhuBtn = document.getElementById('btn-export-ahu');
    if (exportAhuBtn) {
      exportAhuBtn.addEventListener('click', () => {
        try {
          if (window.ExcelExporter) {
            window.ExcelExporter.exportAHUSavingLog(state.selectedMonthMatrix, state.data);
            showToast('AHU & BTU Savings Excel Download Triggered!');
          }
        } catch (err) {
          console.error('Export Error:', err);
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

    # 2. Monthly Energy Cost Trend Chart
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

    # 3. EPI Trend vs 2025 Benchmark Chart
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

    # 4. YTD YoY Energy Component Net Cost Savings Chart
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

  document.addEventListener('DOMContentLoaded', () => {
    initData();
    render();
  });
})();

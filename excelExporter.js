/**
 * Enhanced Real-Time Excel Exporter Engine for AHU Saving Log, Daily KWH Energy Report & Single Month Dashboards
 * Supports both Server-Side (OpenPyXL) and Client-Side (SheetJS / XLSX) with full August 2026 & multi-month synchronization
 */

window.ExcelExporter = {
  // Master DT-3 KWH Excel Export (all sheets synchronized)
  exportKWHDailyLog: function (selectedMonth, trackerData) {
    trackerData = trackerData || {};
    
    // 1. Try Server-Side API first if hosted on Python server
    if (window.location.protocol.startsWith('http') && !window.location.hostname.includes('vercel.app')) {
      fetch('/api/export/kwh', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(trackerData)
      })
      .then(response => {
        if (!response.ok) throw new Error("Server status " + response.status);
        return response.blob();
      })
      .then(blob => {
        this.triggerBlobDownload(blob, "DT-3 Daily energy Consumption  Updated.xlsx");
      })
      .catch(err => {
        console.warn("Server API export unavailable, using client-side dynamic Excel engine:", err);
        this.exportKWHDailyLogClientSide(selectedMonth, trackerData);
      });
    } else {
      this.exportKWHDailyLogClientSide(selectedMonth, trackerData);
    }
  },

  // Dynamic Client-Side DT-3 KWH Excel Generator
  exportKWHDailyLogClientSide: function(selectedMonth, trackerData) {
    const fileName = "DT-3 Daily energy Consumption  Updated.xlsx";
    fetch("./" + encodeURIComponent(fileName))
      .then(res => {
        if (!res.ok) throw new Error("Could not fetch template file");
        return res.arrayBuffer();
      })
      .then(buffer => {
        if (typeof XLSX === 'undefined') {
          this.downloadFolderFile(fileName);
          return;
        }

        const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellStyles: true });
        const kwhDaily = trackerData.kwh_daily || {};

        // Update sheets with user entered readings
        wb.SheetNames.forEach(sheetName => {
          if (sheetName === 'Dashboard' || sheetName === 'summary') return;
          const ws = wb.Sheets[sheetName];
          if (!ws) return;

          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:AJ100');
          const colDateMap = {};

          // Intelligent Date Column Locator: Scan header rows (Row 2, 4, 40, 84 - 0-indexed: 1, 3, 39, 83)
          const checkRows = [1, 3, 39, 83];
          for (let C = 4; C <= range.e.c; C++) {
            for (let rIdx of checkRows) {
              const cellAddress = XLSX.utils.encode_cell({ r: rIdx, c: C });
              const cell = ws[cellAddress];
              if (cell && cell.v) {
                let dateStr = '';
                if (cell.v instanceof Date) {
                  dateStr = cell.v.toISOString().slice(0, 10);
                } else if (typeof cell.v === 'number' && cell.v > 40000 && cell.v < 60000) {
                  // Excel serial date to JS Date
                  const dateObj = new Date(Math.round((cell.v - 25569) * 86400 * 1000));
                  dateStr = dateObj.toISOString().slice(0, 10);
                } else if (typeof cell.v === 'string' && cell.v.includes('-') && cell.v.length >= 10) {
                  dateStr = cell.v.slice(0, 10);
                }
                if (dateStr && !colDateMap[dateStr]) {
                  colDateMap[dateStr] = C;
                }
              }
            }
          }

          // If sheet is Aug-2026, ensure header banner and August dates are written
          if (sheetName.toLowerCase().includes('aug')) {
            const a1Addr = XLSX.utils.encode_cell({ r: 0, c: 0 });
            ws[a1Addr] = { t: 's', v: 'AUGUST -  2026' };
          }

          // Populate user entered readings for dates that match
          Object.keys(kwhDaily).forEach(dStr => {
            if (colDateMap[dStr] !== undefined) {
              const colIdx = colDateMap[dStr];
              const dayData = kwhDaily[dStr];

              // EB meters (rows 6, 8, 10, 12, 14, 16, 18, 20 -> 0-indexed: 5, 7, 9, 11, 13, 15, 17, 19)
              const ebRows = [5, 7, 9, 11, 13, 15, 17, 19];
              (dayData.eb || []).forEach((ebItem, idx) => {
                if (idx < ebRows.length) {
                  const r = ebRows[idx];
                  const addrEB = XLSX.utils.encode_cell({ r: r, c: colIdx });
                  const addrDG = XLSX.utils.encode_cell({ r: r + 1, c: colIdx });
                  ws[addrEB] = { t: 'n', v: Number(ebItem.reading || 0) };
                  if (ebItem.dg_reading !== undefined) {
                    ws[addrDG] = { t: 'n', v: Number(ebItem.dg_reading || 0) };
                  }
                }
              });

              // AHU meters (rows 23, 25, 27, 29 -> 0-indexed: 22, 24, 26, 28)
              const ahuRows = [22, 24, 26, 28];
              (dayData.ahu || []).forEach((ahuItem, idx) => {
                if (idx < ahuRows.length) {
                  const r = ahuRows[idx];
                  const addrAHU = XLSX.utils.encode_cell({ r: r, c: colIdx });
                  const addrDG = XLSX.utils.encode_cell({ r: r + 1, c: colIdx });
                  ws[addrAHU] = { t: 'n', v: Number(ahuItem.reading || 0) };
                  if (ahuItem.dg_reading !== undefined) {
                    ws[addrDG] = { t: 'n', v: Number(ahuItem.dg_reading || 0) };
                  }
                }
              });

              // BTU meters (rows 32, 33, 34, 35 -> 0-indexed: 31, 32, 33, 34)
              const btuRows = [31, 32, 33, 34];
              (dayData.btu || []).forEach((btuItem, idx) => {
                if (idx < btuRows.length) {
                  const r = btuRows[idx];
                  const addrBTU = XLSX.utils.encode_cell({ r: r, c: colIdx });
                  ws[addrBTU] = { t: 'n', v: Number(btuItem.reading || 0) };
                }
              });
            }
          });
        });

        // Write out modified workbook
        const outArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        this.triggerBlobDownload(new Blob([outArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
      })
      .catch(err => {
        console.error("Client side export fallback failed:", err);
        this.downloadFolderFile(fileName);
      });
  },

  // Dedicated Single Month Excel Exporter (High fidelity, standalone clean workbook)
  exportSingleMonthKWH: function(monthKey, trackerData, rates) {
    rates = rates || { kwh: 7.45, btu: 4.30, dg: 33.85 };
    trackerData = trackerData || {};
    const kwhDaily = trackerData.kwh_daily || {};

    if (typeof XLSX === 'undefined') {
      alert("Excel library is loading, please try again in a second.");
      return;
    }

    // Filter dates for this month
    const matchingDates = Object.keys(kwhDaily)
      .filter(d => {
        const dt = new Date(d);
        const mStr = dt.toLocaleString('en-US', { month: 'short' });
        const yStr = dt.toLocaleString('en-US', { year: '2-digit' });
        return `${mStr}-${yStr}` === monthKey;
      })
      .sort();

    if (matchingDates.length === 0) {
      alert(`No daily readings recorded for ${monthKey} yet.`);
      return;
    }

    const wb = XLSX.utils.book_new();

    // 1. Sheet 1: Daily Energy Consumption Breakdown
    const consData = [
      [`CBRE FACILITY MANAGEMENT - DAILY ENERGY CONSUMPTION REPORT (${monthKey})`],
      [`Generated: ${new Date().toLocaleString('en-US')} | KWH Rate: Rs. ${rates.kwh} | DG Rate: Rs. ${rates.dg} | BTU Rate: Rs. ${rates.btu}`],
      [],
      ["Date", "Day", "Total EB (kWh)", "Total DG (kWh)", "Total Power Cum. (kWh)", "Total AHU (kWh)", "Total BTU (Units)", "Daily Cost (INR)"]
    ];

    let sumEB = 0, sumDG = 0, sumCum = 0, sumAHU = 0, sumBTU = 0, sumCost = 0;

    matchingDates.forEach(dStr => {
      const dt = new Date(dStr);
      const dayName = dt.toLocaleString('en-US', { weekday: 'short' });
      const dayData = kwhDaily[dStr];

      let dayEB = 0, dayDG = 0, dayAHU = 0, dayBTU = 0;

      // EB
      (dayData.eb || []).forEach(m => {
        // Calculate consumption from prev reading
        const prev = (window.TrackerUtils?.getExactPreviousReading ? window.TrackerUtils.getExactPreviousReading(dStr, 'eb', m.id) : { reading: m.reading, dg_reading: m.dg_reading });
        const cEB = Math.max(0, (Number(m.reading) || 0) - (Number(prev.reading) || 0));
        const cDG = Math.max(0, (Number(m.dg_reading) || 0) - (Number(prev.dg_reading) || 0));
        dayEB += cEB;
        dayDG += cDG;
      });

      // AHU
      (dayData.ahu || []).forEach(m => {
        const prev = (window.TrackerUtils?.getExactPreviousReading ? window.TrackerUtils.getExactPreviousReading(dStr, 'ahu', m.id) : { reading: m.reading });
        const cAHU = Math.max(0, (Number(m.reading) || 0) - (Number(prev.reading) || 0));
        dayAHU += cAHU;
      });

      // BTU
      (dayData.btu || []).forEach(m => {
        const prev = (window.TrackerUtils?.getExactPreviousReading ? window.TrackerUtils.getExactPreviousReading(dStr, 'btu', m.id) : { reading: m.reading });
        const cBTU = Math.max(0, (Number(m.reading) || 0) - (Number(prev.reading) || 0));
        dayBTU += cBTU;
      });

      const dayCum = dayEB + dayDG;
      const dayCost = (dayEB * rates.kwh) + (dayDG * rates.dg) + (dayBTU * rates.btu);

      sumEB += dayEB;
      sumDG += dayDG;
      sumCum += dayCum;
      sumAHU += dayAHU;
      sumBTU += dayBTU;
      sumCost += dayCost;

      consData.push([
        dStr,
        dayName,
        Number(dayEB.toFixed(1)),
        Number(dayDG.toFixed(1)),
        Number(dayCum.toFixed(1)),
        Number(dayAHU.toFixed(1)),
        Number(dayBTU.toFixed(1)),
        Number(dayCost.toFixed(2))
      ]);
    });

    // Add Totals & Averages
    const count = matchingDates.length;
    consData.push([]);
    consData.push([
      "MONTH TOTAL",
      `${count} Days`,
      Number(sumEB.toFixed(1)),
      Number(sumDG.toFixed(1)),
      Number(sumCum.toFixed(1)),
      Number(sumAHU.toFixed(1)),
      Number(sumBTU.toFixed(1)),
      Number(sumCost.toFixed(2))
    ]);
    consData.push([
      "DAILY AVERAGE",
      "-",
      Number((sumEB / count).toFixed(1)),
      Number((sumDG / count).toFixed(1)),
      Number((sumCum / count).toFixed(1)),
      Number((sumAHU / count).toFixed(1)),
      Number((sumBTU / count).toFixed(1)),
      Number((sumCost / count).toFixed(2))
    ]);

    // Sized columns and gridlines for Consumption Sheet
    wsCons['!cols'] = [
      { wch: 15 }, // Date
      { wch: 10 }, // Day
      { wch: 20 }, // Total EB
      { wch: 20 }, // Total DG
      { wch: 24 }, // Total Power Cum.
      { wch: 20 }, // Total AHU
      { wch: 20 }, // Total BTU
      { wch: 22 }  // Daily Cost (INR)
    ];
    wsCons['!views'] = [{ showGridLines: true }];
    XLSX.utils.book_append_sheet(wb, wsCons, `${monthKey} Consumption`);

    // 2. Sheet 2: Raw Daily Meter Readings
    const meterHeaders = ["Date", "Day"];
    const ebNames = ["EB-1 (540430038840)", "EB-2 (540430038845)", "EB-3 (540430038844)", "EB-4 (540430038821)", "EB-5 (540430038849)", "EB-6 (540430038646)", "EB-7 (540430038843)", "EB-8 (540430038848)"];
    ebNames.forEach(n => meterHeaders.push(`${n} Reading`, `${n} DG`));
    for (let i = 1; i <= 4; i++) meterHeaders.push(`AHU-${i} Reading`, `AHU-${i} DG`);
    for (let i = 1; i <= 4; i++) meterHeaders.push(`BTU-${i} Reading`);

    const readingsData = [
      [`CBRE FACILITY MANAGEMENT - DAILY METER READINGS LOG (${monthKey})`],
      meterHeaders
    ];

    matchingDates.forEach(dStr => {
      const dt = new Date(dStr);
      const dayName = dt.toLocaleString('en-US', { weekday: 'short' });
      const dayData = kwhDaily[dStr];
      const row = [dStr, dayName];

      (dayData.eb || []).forEach(m => {
        row.push(Number(m.reading || 0), Number(m.dg_reading || 0));
      });
      (dayData.ahu || []).forEach(m => {
        row.push(Number(m.reading || 0), Number(m.dg_reading || 0));
      });
      (dayData.btu || []).forEach(m => {
        row.push(Number(m.reading || 0));
      });

      readingsData.push(row);
    });

    const wsReadings = XLSX.utils.aoa_to_sheet(readingsData);
    const readCols = [{ wch: 15 }, { wch: 10 }];
    for (let i = 0; i < 30; i++) readCols.push({ wch: 18 });
    wsReadings['!cols'] = readCols;
    wsReadings['!views'] = [{ showGridLines: true }];
    XLSX.utils.book_append_sheet(wb, wsReadings, `${monthKey} Readings`);

    // Write file
    const outArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const outFileName = `DT3_Energy_Report_${monthKey}.xlsx`;
    this.triggerBlobDownload(new Blob([outArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), outFileName);
  },

  // Export AHU Saving Excel file
  exportAHUSavingLog: function (selectedMonth, trackerData) {
    trackerData = trackerData || {};

    if (window.location.protocol.startsWith('http') && !window.location.hostname.includes('vercel.app')) {
      fetch('/api/export/ahu', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(trackerData)
      })
      .then(response => {
        if (!response.ok) throw new Error("Server status " + response.status);
        return response.blob();
      })
      .then(blob => {
        this.triggerBlobDownload(blob, "AHU_Saved_Hours_and_Energy_Cost_Report.xlsx");
      })
      .catch(err => {
        console.warn("Server API export unavailable, serving AHU file directly:", err);
        this.downloadFolderFile("AHU_Saved_Hours_and_Energy_Cost_Report.xlsx");
      });
    } else {
      this.downloadFolderFile("AHU_Saved_Hours_and_Energy_Cost_Report.xlsx");
    }
  },

  // Export Live Transaction History to CSV
  exportTransactionHistoryCSV: function(transactions) {
    if (!transactions || !transactions.length) {
      alert("No transaction records to export.");
      return;
    }

    const headers = ["TXN ID", "Timestamp", "Operational Date", "Category", "Meter / Asset", "Location", "Prev Reading", "New Reading", "Consumption (kWh)", "Cost (INR)", "Status", "Source"];
    const rows = transactions.map(t => [
      `"${t.id || ''}"`,
      `"${t.displayTime || t.timestamp || ''}"`,
      `"${t.date || ''}"`,
      `"${t.categoryLabel || t.category || ''}"`,
      `"${t.meterName || ''}"`,
      `"${t.location || '3F'}"`,
      t.prevReading !== undefined ? t.prevReading : '',
      t.newReading !== undefined ? t.newReading : '',
      t.consumption !== undefined ? t.consumption : '',
      t.cost !== undefined ? t.cost : '',
      `"${t.status || 'Committed'}"`,
      `"${t.source || 'User Input'}"`
    ]);

    const csvContent = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const todayStr = new Date().toISOString().slice(0, 10);
    this.triggerBlobDownload(blob, `CBRE_KWH_Transaction_Audit_History_${todayStr}.csv`);
  },

  // Helper to trigger blob download in browser
  triggerBlobDownload: function(blob, fileName) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => window.URL.revokeObjectURL(url), 1000);
  },

  // Direct folder file downloader
  downloadFolderFile: function(fileName) {
    const a = document.createElement('a');
    a.href = "./" + encodeURIComponent(fileName);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }
};

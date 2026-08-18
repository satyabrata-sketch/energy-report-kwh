/**
 * Dynamic Real-Time Excel Exporter Engine for AHU Saving Log, Daily KWH Energy Report & Transactions
 * Supports both Server-Side (OpenPyXL) and Client-Side (SheetJS / XLSX) with full August 2026 & dynamic month synchronization
 */

window.ExcelExporter = {
  // Export DT-3 KWH Excel file with real-time August 2026 & all dynamic data
  exportKWHDailyLog: function (selectedMonth, trackerData) {
    trackerData = trackerData || {};
    
    // 1. Try Server-Side API first
    if (window.location.protocol.startsWith('http')) {
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
          // If XLSX library isn't loaded, fallback to raw template download
          this.triggerBlobDownload(new Blob([buffer]), fileName);
          return;
        }

        const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellStyles: true });
        const kwhDaily = trackerData.kwh_daily || {};

        // Update sheets with user entered readings
        wb.SheetNames.forEach(sheetName => {
          if (sheetName === 'Dashboard' || sheetName === 'summary') return;
          const ws = wb.Sheets[sheetName];
          if (!ws) return;

          // Locate date columns in row 4 (1-indexed row 4 is index 3)
          // Scan row 4 cells
          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z100');
          const colDateMap = {};

          for (let C = 4; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 3, c: C }); // Row 4 (index 3)
            const cell = ws[cellAddress];
            if (cell && cell.v) {
              let dateStr = '';
              if (cell.v instanceof Date) {
                dateStr = cell.v.toISOString().slice(0, 10);
              } else if (typeof cell.v === 'number') {
                // Excel serial date to JS Date
                const dateObj = new Date(Math.round((cell.v - 25569) * 86400 * 1000));
                dateStr = dateObj.toISOString().slice(0, 10);
              } else if (typeof cell.v === 'string' && cell.v.includes('-')) {
                dateStr = cell.v.slice(0, 10);
              }
              if (dateStr) {
                colDateMap[dateStr] = C;
              }
            }
          }

          // If dates match in kwhDaily, update cell values
          Object.keys(kwhDaily).forEach(dStr => {
            if (colDateMap[dStr] !== undefined) {
              const colIdx = colDateMap[dStr];
              const dayData = kwhDaily[dStr];

              // EB meters (rows 6, 8, 10, 12, 14, 16, 18, 20)
              const ebRows = [5, 7, 9, 11, 13, 15, 17, 19]; // 0-indexed
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

              // AHU meters (rows 23, 25, 27, 29) -> 0-indexed: 22, 24, 26, 28
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

              // BTU meters (rows 32, 33, 34, 35) -> 0-indexed: 31, 32, 33, 34
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

  // Export AHU Saving Excel file
  exportAHUSavingLog: function (selectedMonth, trackerData) {
    trackerData = trackerData || {};

    if (window.location.protocol.startsWith('http')) {
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

/**
 * Excel Exporter Engine for AHU Saving Log and Daily KWH Energy Report
 * Directly downloads the EXACT original Excel files from the folder:
 * - DT-3 Daily energy Consumption  Updated.xlsx
 * - AHU_Saved_Hours_and_Energy_Cost_Report.xlsx
 */

window.ExcelExporter = {
  // Export DT-3 KWH Excel file (Exact file from folder)
  exportKWHDailyLog: function (selectedMonth, trackerData) {
    if (window.location.protocol.startsWith('http')) {
      fetch('/api/export/kwh', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(trackerData || {})
      })
      .then(response => {
        if (!response.ok) throw new Error("Server status " + response.status);
        return response.blob();
      })
      .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = "DT-3 Daily energy Consumption  Updated.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      })
      .catch(err => {
        console.warn("Server API export failed, serving folder file directly:", err);
        this.downloadFolderFile("DT-3 Daily energy Consumption  Updated.xlsx");
      });
    } else {
      this.downloadFolderFile("DT-3 Daily energy Consumption  Updated.xlsx");
    }
  },

  // Export AHU Saving Excel file (Exact file from folder)
  exportAHUSavingLog: function (selectedMonth, trackerData) {
    if (window.location.protocol.startsWith('http')) {
      fetch('/api/export/ahu', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(trackerData || {})
      })
      .then(response => {
        if (!response.ok) throw new Error("Server status " + response.status);
        return response.blob();
      })
      .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = "AHU_Saved_Hours_and_Energy_Cost_Report.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
      })
      .catch(err => {
        console.warn("Server API export failed, serving folder file directly:", err);
        this.downloadFolderFile("AHU_Saved_Hours_and_Energy_Cost_Report.xlsx");
      });
    } else {
      this.downloadFolderFile("AHU_Saved_Hours_and_Energy_Cost_Report.xlsx");
    }
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

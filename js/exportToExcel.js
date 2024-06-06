function exportToExcel(tableData) {
    // Create a new workbook
    var wb = XLSX.utils.book_new();
    
    // Convert the TableData object to a 2D array
    var ws_data = [
        ["Id", "Text", "Rows", "Columns", "Total Cells"],
    ];

    if (tableData.length > 0) {
        for(let i = 0; i < tableData.length; i++)
        {
            ws_data.push([i + 1, tableData[i].range, tableData[i].rows, tableData[i].columns, tableData[i].totalCells]);
        }
    }
    
    // Create a worksheet from the data array
    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    
    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    // Generate the Excel file and trigger the download
    XLSX.writeFile(wb, "export.xlsx");
}
/**
 * Forces columns to re-parse as dates and applies yyyy-mm-dd format.
 */
function forceDateFormatData1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Data' not found!");
    return;
  }

  const rangesToFormat = [
    "I:I", "U:AA", "AF:AF", "AH:AK","AU:AY", "BC:BF", "BH:BH"
  ];

  rangesToFormat.forEach(rangeString => {
    const range = sheet.getRange(rangeString);
    
    // 1. Clear existing formatting to "reset" the cells
    range.setNumberFormat("@"); // Set to plain text temporarily
    range.setNumberFormat("yyyy-mm-dd"); // Set to desired date format
    
    // 2. Force the sheet to "re-type" the data
    // This trick of setting values to themselves forces Google Sheets to re-parse strings as dates
    const values = range.getValues();
    range.setValues(values);
  });
}

/**
 * 即使是文本格式的时间戳也能强制转换为 yyyy-mm-dd
 */
function absoluteBruteForceDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Data_1' not found!");
    return;
  }

  const rangesToFormat = ["I:I", "U:AA", "AF:AF", "AH:AK","AU:AY", "BC:BF", "BH:BH"];

  rangesToFormat.forEach(rangeString => {
    const range = sheet.getRange(rangeString);
    
    // CRITICAL: getDisplayValues reads EXACTLY what is on your screen ("2/17/2026 23:00:00")
    // This bypasses all of Google Sheets' stubborn data type rules.
    const displayValues = range.getDisplayValues();
    
    const newValues = displayValues.map(row => {
      return row.map(cell => {
        if (!cell || cell.trim() === "") return "";
        
        // 1. Split by space and keep only the date part -> "2/17/2026"
        const datePart = cell.trim().split(" ")[0]; 
        
        // 2. Split the date by "/" or "-"
        const parts = datePart.split(/[\/\-]/); 
        
        if (parts.length === 3) {
          let year, month, day;
          
          // If the year is at the end (M/D/YYYY) -> e.g., 2/17/2026
          if (parts[2].length === 4) { 
            month = parts[0].padStart(2, '0');
            day = parts[1].padStart(2, '0');
            year = parts[2];
          } 
          // If the year is at the beginning (YYYY/M/D)
          else if (parts[0].length === 4) {
            year = parts[0];
            month = parts[1].padStart(2, '0');
            day = parts[2].padStart(2, '0');
          } else {
            return cell; // Return original if it doesn't match a known pattern
          }
          
          // 3. Rebuild it as a strict string
          return `${year}-${month}-${day}`;
        }
        
        return cell; // Return original if not a date (like headers)
      });
    });
    
    // 4. Force the column to be PLAIN TEXT so Sheets CANNOT auto-add the time back
    range.setNumberFormat("@");
    
    // 5. Write the static YYYY-MM-DD text back to the sheet
    range.setValues(newValues);
  });
}

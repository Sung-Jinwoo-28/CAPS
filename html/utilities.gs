function getWorkshopCount(workshopName) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheetName = workshopName;

        // Ensure exact sheet names are used
        // "Digital Marketing", "From Data To Decisions", "Financial Literacy for Professionals"
        // This mapping is optional if inputs are exact, but good for safety.
        if (workshopName === "Digital Marketing") sheetName = "Digital Marketing";
        else if (workshopName === "From Data To Decisions") sheetName = "From Data To Decisions";
        else if (workshopName === "Financial Literacy for Professionals") sheetName = "Financial Literacy for Professionals";
        else return 0; // If not one of the tracked workshops, return 0 (or count existing sheet?)

        const sheet = ss.getSheetByName(sheetName);
        
        // If sheet exists, count is total rows - 1 (header). 
        // If sheet doesn't exist, count is 0.
        if (sheet) {
            return Math.max(0, sheet.getLastRow() - 1);
        }
        return 0;
    } catch (e) {
        console.error("Count check failed for " + workshopName, e);
        return 0; // Fail open
    }
}

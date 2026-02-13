
/**
 * counts.gs - Standalone Script for Fetching Workshop Counts
 * 
 * deployed as a web app, this script returns the current participant count
 * for each workshop by counting rows in their respective sheets.
 */

const WORKSHOP_NAMES = [
    "Financial Literacy for Professionals",
    "Digital Marketing",
    "From Data To Decisions"
];

const MAX_CAPACITY = 400; // Global event capacity

function doGet(e) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let totalEntries = 0;
        const workshopCounts = {};

        WORKSHOP_NAMES.forEach(name => {
            const sheet = ss.getSheetByName(name);
            // If sheet exists, count rows minus header. If not, 0.
            const count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
            workshopCounts[name] = count;
            totalEntries += count;
        });

        const seatsLeft = Math.max(0, MAX_CAPACITY - totalEntries);

        return ContentService.createTextOutput(JSON.stringify({
            success: true,
            totalEntries: totalEntries,
            seatsLeft: seatsLeft,
            workshopCounts: workshopCounts
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            success: false,
            error: err.message
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function main(
    workbook: ExcelScript.Workbook,
    rawBatchText: string,
    receivedDate: string
): number {
    // 1. Get active worksheet or fallback to "Unit Register"
    let sheet: ExcelScript.Worksheet = workbook.getWorksheet("Unit Register");
    if (!sheet) {
        sheet = workbook.getActiveWorksheet();
    }

    // 2. Check if the table "Unit Register" exists; if not, look for any table or create one
    let table: ExcelScript.Table = sheet.getTable("Unit Register");

    if (!table) {
        // Fallback: grab the first table on the sheet if the name differs
        let tables = sheet.getTables();
        if (tables.length > 0) {
            table = tables[0];
        } else {
            throw new Error("No Excel Table found on the sheet 'Unit Register'. Please select your data and click Insert > Table first.");
        }
    }

    // 3. Split raw text by newlines to process line by line
    let lines: string[] = rawBatchText.split(/\r?\n/);

    // 4. Generate a unique batch ID for this submission
    let batchID: string = "BATCH-" + Math.floor(100000 + Math.random() * 900000);

    // 5. Safely calculate starting row sequence number
    let existingRowsCount: number = 0;
    let rows = table.getRows();
    if (rows) {
        existingRowsCount = rows.length;
    }

    let newRows: (string | number)[][] = [];

    lines.forEach((line: string) => {
        let trimmedLine: string = line.trim();
        if (trimmedLine.length > 0) {
            // Split each line by spaces into tokens: [Order ID, Serial Number, Model]
            let tokens: string[] = trimmedLine.split(/\s+/);

            if (tokens.length >= 3) {
                let rowNumber: number = existingRowsCount + newRows.length + 1;
                newRows.push([
                    rowNumber,     // Column 1: ROW
                    tokens[0],     // Column 2: Order ID
                    tokens[1],     // Column 3: Serial Number
                    tokens[2],     // Column 4: Product Model
                    receivedDate,  // Column 5: Received Date
                    batchID        // Column 6: Batch ID
                ]);
            }
        }
    });

    // 6. Bulk add the rows to the table
    if (newRows.length > 0) {
        table.addRows(undefined, newRows as string[][]);
    }

    return newRows.length;
}

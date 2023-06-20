/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * Basic function to show how to insert a value into cell A1 on the selected Excel worksheet.
 */
export function helloButton() {

    return Excel.run(context => {

        // Insert text 'Hello world!' into cell A1.
        context.workbook.worksheets.getActiveWorksheet().getRange("A1").values = [['Hello world!']];

        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}

export function sheetValues() {
    return Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load("values");

        return context.sync().then(function () {
            return usedRange.values;
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

export function formulaOfSelected() {
    return Excel.run(function (context) {
        const selectedRange = context.workbook.getSelectedRange();
        const cellToAnalyse = selectedRange.getCell(0, 0);

        // Load the formula of the selected cell
        cellToAnalyse.load("formulas");

        return context.sync().then(function () {
            return cellToAnalyse.formulas;
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

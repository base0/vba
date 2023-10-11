async function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    while (true) {
        await sheet.calculate(true);
        for (let i = 0; i < 100000000; i++) { }
    }
}

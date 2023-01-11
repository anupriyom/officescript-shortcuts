function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();

    const retailers = sheet.getRange('e2')
        .getExtendedRange(ExcelScript.KeyboardDirection.down)
        .getValues()
        .map(([x]: [string]) => x);

    const retailerset = new Set(retailers);

    const finyear = sheet.getRange('r2')
        .getExtendedRange(ExcelScript.KeyboardDirection.down)
        .getValues()
        .map(([x]: [string]) => x);

    console.log(retailerset);

    sheet.getRange('w2')
        .getExtendedRange(ExcelScript.KeyboardDirection.down)
        .getOffsetRange(0, 1)
        .setValues(retailers.reduce((arr = [], x, i) => {
            const v = (finyear[i] == '2023' && retailerset.delete(x)) ? 1 : 0
            arr.push(v)
            return arr
        }, []).map((x)=>[x]));
  
}

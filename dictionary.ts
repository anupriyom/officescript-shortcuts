
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet()
    const extenddown = (cell: string) => sheet.getRange(cell).getExtendedRange(ExcelScript.KeyboardDirection.down).getValues().map((x) => x[0])
    const z  = extenddown('b4') ! as string[]
    const qty = extenddown('h4') ! as number[]
    console.log(z.reduce((zones, zone, i) => {
        return zones[zone] = zone in zones ? zones[zone] + qty[i] : 0, zones
      }, {}))
}

async function checkMissedClasses(sheet, start, column, totalClasses, studentCount) {
    let j = 0
    while (sheet.getCell(start, column + j).value !== 'Situação') {
        j += 1
    }
    for (let i = 0; i < studentCount; i += 1) {
        if (sheet.getCell(start + i + 1, column).value > totalClasses / 4) {
            sheet.getCell(start + i + 1, column + j).value = 'Reprovado por Falta'
            if (sheet.getCell(start, column + j + 1).value === 'Nota para Aprovação Final') {
                sheet.getCell(start + i + 1, column + j + 1).value = 0
            }
        }
        await sheet.saveUpdatedCells()
    }
    console.log('Failed by Missed Classes checked')
}

module.exports = { checkMissedClasses }

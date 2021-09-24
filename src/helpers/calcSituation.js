async function calcSituation(sheet, P1, P2, P3, line, column, studentCount) {
    for (let i = 1; i <= studentCount; i += 1) {
        if (sheet.getCell(line + i, column).value === null) {
            let total = 0
            total += sheet.getCell(line + i, P1).value
            total += sheet.getCell(line + i, P2).value
            total += sheet.getCell(line + i, P3).value
            const average = total / 3
            if (average >= 70) {
                sheet.getCell(line + i, column).value = 'Aprovado'
                if (sheet.getCell(line, column + 1).value === 'Nota para Aprovação Final') {
                    sheet.getCell(line + i, column + 1).value = 0
                }
            } else if (average >= 50) {
                sheet.getCell(line + i, column).value = 'Exame Final'
                if (sheet.getCell(line, column + 1).value === 'Nota para Aprovação Final') {
                    sheet.getCell(line + i, column + 1).value = Math.ceil(100 - average)
                }
            } else {
                sheet.getCell(line + i, column).value = 'Reprovado'
                if (sheet.getCell(line, column + 1).value === 'Nota para Aprovação Final') {
                    sheet.getCell(line + i, column + 1).value = 0
                }
            }
            await sheet.saveUpdatedCells()
        }
    }
    console.log('Situtations Calculated')
}

module.exports = { calcSituation }

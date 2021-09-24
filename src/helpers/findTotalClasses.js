function findTotalClasses(sheet, line, column) {
    const aux = sheet.getCell(line, column).value.match(/(\d+)/)
    console.log('Found Total Classes')
    return aux[0]
}

module.exports = { findTotalClasses }

function getStudentCount(sheet, start, column) {
    let studentCount = 0
    while (sheet.getCell(studentCount + start, column).value != null) {
        studentCount += 1
    }
    console.log('Got Student Count')
    return studentCount
}

module.exports = { getStudentCount }

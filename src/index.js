const { GoogleSpreadsheet } = require('google-spreadsheet')

const credentials = require('../credentials.json')

const doc = new GoogleSpreadsheet('17ICUzzaVTnV9IOIp8pH2wn3o2KMx8YfNWHfBcbeYtxw')

const { calcSituation } = require('./helpers/calcSituation')
const { checkMissedClasses } = require('./helpers/checkMissedClasses')
const { findTotalClasses } = require('./helpers/findTotalClasses')
const { getStudentCount } = require('./helpers/getStudentCount')

async function accessSpreadsheet() {
    await doc.useServiceAccountAuth(credentials)
    await doc.loadInfo()
    const sheet = doc.sheetsByIndex[0]

    let P1Index
    let P2Index
    let P3Index
    let studentCount
    let totalClasses
    let i = 0
    let j = 0

    await sheet.loadCells()

    while (i < sheet.rowCount && j < sheet.columnCount) {
        if (sheet.getCell(i, j).value == null) {
            i += 1
            j = 0
            if (sheet.getCell(i, j).value == null) {
                break
            }
        }
        if (typeof (sheet.getCell(i, j).value) === 'string') {
            if (sheet.getCell(i, j).value.includes('Total de aulas no semestre:')) {
                console.log('Finding Total Classes...')
                totalClasses = findTotalClasses(sheet, i, j)
            }
        }
        if (sheet.getCell(i, j).value === 'Matricula') {
            console.log('Getting Student Count...')
            studentCount = getStudentCount(sheet, i + 1, j)
        }
        if (sheet.getCell(i, j).value === 'Faltas') {
            console.log('Checking Failed by Missed Classes...')
            await checkMissedClasses(sheet, i, j, totalClasses, studentCount)
        }
        if (sheet.getCell(i, j).value === 'P1') {
            P1Index = j
            console.log('Found P1 Index')
        }
        if (sheet.getCell(i, j).value === 'P2') {
            P2Index = j
            console.log('Found P2 Index')
        }
        if (sheet.getCell(i, j).value === 'P3') {
            P3Index = j
            console.log('Found P3 Index')
        }
        if (sheet.getCell(i, j).value === 'Situação') {
            console.log('Calculating Situations')
            await calcSituation(sheet, P1Index, P2Index, P3Index, i, j, studentCount)
        }
        j += 1
    }
}

accessSpreadsheet()

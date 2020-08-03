const express = require('express')
const bodyParser = require('body-parser')
const {google} = require('googleapis')
const keys = require('../keys.json')

const sheetRouter = express.Router()
sheetRouter.use(bodyParser.json())



sheetRouter.route('/')
.all((req, res, next) => {
    res.statusCode = 200
    res.setHeader('Content-Type', 'text/plain')
    next()
})
.get((req, res, next) => {
    res.statusCode = 403
    res.end('GET operation not supported on /sheet')
})
.post(async (req, res) => {
    const client = new google.auth.JWT(
        req.body.client_email,
        null,
        req.body.private_key,
        ['https://www.googleapis.com/auth/spreadsheets']
    )
    const gsapi = google.sheets({version: "v4", auth: client})
    const dataObject = req.body.data
    const dataArray = []

    for (let row = 0; row < dataObject.length; row++) {
        dataArray[row] = [dataObject[row].value1, dataObject[row].value2]
    }

    const request = {
        spreadsheetId: req.body.spreadsheetId,
        range: 'A2',
        valueInputOption: 'USER_ENTERED',
        resource: {values: dataArray}
    }
    let response = await gsapi.spreadsheets.values.update(optValue)
    // let response = await gsapi.spreadsheets.sheets.copyTo(request)

    res.statusCode = response.status
    res.end(response.statusText)
})
.put((req, res, next) => {
    res.statusCode = 403
    res.end('PUT operation not supported on /sheet')
})
.delete((req, res, next) => {
    res.statusCode = 403
    res.end('DELETE operation not supported on /sheet')
})


module.exports = sheetRouter
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

    const date = new Date()
    const month = date.getMonth() + 1

    const sheetVal = {
        spreadsheetId: req.body.spreadsheetId,
        range: req.body.date+'!A2',
        valueInputOption: 'USER_ENTERED',
        resource: {values: dataArray}
    }

    const newSheet = {
        spreadsheetId: req.body.spreadsheetId,
        resource: {
            requests: [{
                duplicateSheet: {
                    "sourceSheetId": 0,
                    "insertSheetIndex": 1,
                    // "newSheetId": req.body.date,
                    "newSheetName": req.body.date
                }
            }]
        }
    }
    await gsapi.spreadsheets.batchUpdate(newSheet, async (err, response) => {
        if (err) {
            console.log(err.message)
            return res.json({
                statusCode: 400,
                message: err.message
            })
        } else {
            console.log('New Sheet has been Created')
            await gsapi.spreadsheets.values.update(sheetVal, (err, response) => {
                if (err) {
                    console.log(err.message)
                    return res.json({
                        statusCode: 400,
                        message: err.message
                    })
                } else {
                    return res.json({
                        "statusCode": 200,
                        "message": 'New Data has been added to new sheet!',
                        "spreadsheetId": req.body.spreadsheetId
                    })
                }
            })
        }
    })
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
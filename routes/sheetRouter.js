const express = require('express')
const bodyParser = require('body-parser')
const {
    google
} = require('googleapis')

const sheetRouter = express.Router()
sheetRouter.use(bodyParser.json())

sheetRouter.route('/')
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
        const gsapi = google.sheets({
            version: "v4",
            auth: client
        })
        const dataObject = req.body.data
        const dataArray = []

        for (let row = 0; row < dataObject.length; row++) {
            dataArray[row] = [dataObject[row].value1, dataObject[row].value2]
        }

        const sheetVal = {
            spreadsheetId: req.body.spreadsheetId,
            range: req.body.date + '!A8',
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: dataArray
            }
        }

        const newSheet = {
            spreadsheetId: req.body.spreadsheetId,
            resource: {
                requests: [{
                    duplicateSheet: {
                        "sourceSheetId": 0,
                        "insertSheetIndex": 1,
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

sheetRouter.route('/newSpreadsheet')
    .post(async (req, res, next) => {
        const client = new google.auth.JWT(
            req.body.client_email,
            null,
            req.body.private_key,
            ['https://www.googleapis.com/auth/spreadsheets']
        )

        const gsapi = google.sheets({
            version: "v4",
            auth: client
        })
        const dataObject = req.body.data
        const dataArray = []

        dataObject.forEach(data => {
            dataArray.push([data.coa, data.type, data.description, data.current_month, data.budget, data.yeartodate])
        });

        const sheetVal = {
            spreadsheetId: req.body.spreadsheetId,
            range: 'A8',
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: dataArray
            }
        }

        await gsapi.spreadsheets.values.update(sheetVal).then(async (response) => {
                const requestData = {
                    spreadsheetId: req.body.spreadsheetId,
                    ranges: [],
                    includeGridData: true,
                };
                return await gsapi.spreadsheets.get(requestData)
            }, (err) => next(err))
            .then(async (sheetData) => {

                const copyRequest = []

                // Get all sheet ID
                for (let sheet = 0; sheet < sheetData.data.sheets.length; sheet++) {
                    copyRequest[sheet] = {
                        spreadsheetId: req.body.spreadsheetId,
                        sheetId: sheetData.data.sheets[sheet].properties.sheetId,
                        resource: {
                            destinationSpreadsheetId: req.body.destination_spreadsheetId,
                        }
                    }
                }
                const resultStatus = []
                copyRequest.forEach(async (request) => {
                    await delay()
                    await gsapi.spreadsheets.sheets.copyTo(request)
                });

                res.statusCode = 200
                res.json({
                    message: 'Copy Data Success',
                    copySheetStatus: resultStatus
                })

            }, (err) => next(err))
            .catch((err) => next(err))
    })

function delay() {
    // `delay` returns a promise
    return new Promise(function (resolve, reject) {
        // Only `delay` is able to resolve or reject the promise
        setTimeout(function () {
            resolve(); // After 3 seconds, resolve the promise with value 42
        }, 2000);
    });
}

module.exports = sheetRouter
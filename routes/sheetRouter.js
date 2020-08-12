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
        const date = new Date()
        const dateNow = date.getFullYear() + '-' + date.getMonth() + '-' + date.getDay()
        const dataBodyArray = []
        const dataHeaderArray = [[req.body.propertyName], [req.body.date], [req.body.printedBy], [dateNow]]
        
        dataObject.forEach(data => {
            dataBodyArray.push([data.coa, data.type, data.description, data.current_month, data.budget, data.yeartodate])
        });

        const data = [
            {
                range: 'A8',
                values: dataBodyArray
            },
            {
                range: 'masterData!C1:C4',
                values: dataHeaderArray
            }
        ]

        const sheetVal = {
            spreadsheetId: req.body.spreadsheetId,
            resource: {
                data: data,
                valueInputOption: 'USER_ENTERED'
            }
        }

        await gsapi.spreadsheets.values.batchUpdate(sheetVal).then(async (response) => {
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
            var processCopyRequest = async (x) => {
                console.log(x)
                if (x < copyRequest.length) {
                    await gsapi.spreadsheets.sheets.copyTo(copyRequest[x]).then(async () => {
                        const requestDestinationData = {
                            spreadsheetId: req.body.destination_spreadsheetId,
                            ranges: [],
                            includeGridData: false
                        }
                        await delay()
                        return await gsapi.spreadsheets.get(requestDestinationData)
                    }).then(async (sheetData) => {
                        const updateRequest = {
                            spreadsheetId: req.body.destination_spreadsheetId,
                            resource: {
                                requests: [
                                    
                                ],
                            },
                        }

                        var processCondition = async (i) => {
                            if (i < sheetData.data.sheets.length) {
                                var strSheet = sheetData.data.sheets[i].properties.title
                                console.log(sheetData.data.sheets[i].properties.title)
                                if(strSheet == 'Sheet1') {
                                    updateRequest.resource.requests.push(
                                        {"deleteSheet": {
                                            sheetId: sheetData.data.sheets[i].properties.sheetId,
                                        }}
                                    )
                                }
                                if(strSheet.slice(0,8) == 'Copy of ') {
                                    updateRequest.resource.requests.push(
                                        {"updateSheetProperties": {
                                            properties: {
                                                sheetId: sheetData.data.sheets[i].properties.sheetId,
                                                title: strSheet.slice(8),
                                                index: i
                                            },
                                            fields: "index,title"
                                        }}
                                    )
                                }
                                await processCondition(i+1)
                            }
                        }

                        await processCondition(0)

                        return await gsapi.spreadsheets.batchUpdate(updateRequest)

                    })

                    await processCopyRequest(x+1)
                }
            }

            processCopyRequest(0).then(((result) => {
                res.statusCode = 200
                res.json({
                    message: 'Copy Data Success',
                    destinationSpreadsheet: `https://docs.google.com/spreadsheets/d/${req.body.destination_spreadsheetId}/edit#gid=0`,
                    spreadsheetId: req.body.destination_spreadsheetId,
                    propertyName: req.body.propertyName,
                    createdAt: dateNow
                })
            }))
        }, (err) => next(err))
        .catch((err) => next(err))
    })

function delay() {
    return new Promise((resolve) => {
        setTimeout(() => resolve(), 1000);
    });
}

module.exports = sheetRouter
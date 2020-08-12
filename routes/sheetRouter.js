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
                    await gsapi.spreadsheets.sheets.copyTo(copyRequest[x])    
                    await processCopyRequest(x+1)
                }
            }

            processCopyRequest(0)

            // copyRequest.forEach(async (request) => {
            //     // await delay()
            //     await gsapi.spreadsheets.sheets.copyTo(request)
            // });

            const requestDestinationData = {
                spreadsheetId: req.body.destination_spreadsheetId,
                ranges: [],
                includeGridData: false
            }
            await delay()
            return await gsapi.spreadsheets.get(requestDestinationData)
        }, (err) => next(err))
        .then(async (sheetData) => {
            const updateRequest = {
                spreadsheetId: req.body.destination_spreadsheetId,
                resource: {
                    requests: [
                        
                    ],
                },
            }
            var processCondition = async (x) => {
                if (x < sheetData.data.sheets.length) {
                    console.log(sheetData.data.sheets[x].properties.title)
                    if(sheetData.data.sheets[x].properties.title == 'tmp') {
                        updateRequest.resource.requests.push(
                            {"deleteSheet": {
                                sheetId: sheetData.data.sheets[x].properties.sheetId,
                            }}
                        )
                    }
                    else if(sheetData.data.sheets[x].properties.title == 'Copy of MasterData') {
                        updateRequest.resource.requests.push(
                            {"updateSheetProperties": {
                                properties: {
                                    sheetId: sheetData.data.sheets[x].properties.sheetId,
                                    title: 'MasterData',
                                    index: 1
                                },
                                fields: "index,title"
                            }}
                        )
                    }
                    await processCondition(x+1)
                }
            }
            await processCondition(0)
            await gsapi.spreadsheets.batchUpdate(updateRequest).then((result) => {
                res.statusCode = 200
                res.json({
                    message: 'Copy Data Success',
                    destinationSpreadsheet: `https://docs.google.com/spreadsheets/d/${req.body.destination_spreadsheetId}/edit#gid=0`,
                    spreadsheetId: req.body.destination_spreadsheetId,
                    propertyName: req.body.propertyName,
                    createdAt: dateNow,
                    replies: result.replies
                })
            })
        })
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

// function delay() {
//     return new Promise(resolve => setTimeout(resolve, 300));
// }

module.exports = sheetRouter
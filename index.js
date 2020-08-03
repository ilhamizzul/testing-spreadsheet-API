const express = require('express')
const http = require('http')
const bodyParser = require('body-parser')

const sheetRouter = require('./routes/sheetRouter')
const hostname = 'localhost'
const port = 3000

const app = express()
app.use(bodyParser.json())
app.use('/sheet', sheetRouter)

app.use((req, res, next) => {
    res.statusCode = 200
    res.setHeader('Content-Type', 'text/html')
    res.end('<html><body><h1>Server Running!</h1></body></html>')
})

const server = http.createServer(app)
server.listen(port, hostname, () => {
    console.log(`Server running on http://${hostname}:${port}/`)
})
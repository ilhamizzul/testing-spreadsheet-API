const express = require('express')
const http = require('http')
const bodyParser = require('body-parser')
const morgan = require('morgan')
require('dotenv').config()

const sheetRouter = require('./routes/sheetRouter')
const hostname = 'localhost'
const port = 3000

const app = express()
app.use(bodyParser.json())
app.use(morgan('dev'))
app.use('/sheet', sheetRouter)

// catch 404 and forward to error handler
app.use(function (req, res, next) {
    next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
    // set locals, only providing error in development
    res.locals.message = err.message;
    res.locals.error = req.app.get('env') === 'development' ? err : {};

    // render the error page
    res.status(err.status || 500);
    res.json({
        message: err.message,
        error: err
    });
});

const server = http.createServer(app)
server.listen(port, hostname, () => {
    console.log(`Server running on http://${hostname}:${port}/`)
})
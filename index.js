var url = require('url');
var microsoftGraph = require("@microsoft/microsoft-graph-client");
var authHelper = require('./authHelper');
var router = require('./router');
var server = require('./server');

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/insights'] = getLastUsed;

server.start(router.route, handle);

function home(response, request) {
    console.log('Request handler \'home\' was called.');
    response.writeHead(200, { 'Content-Type': 'text/html' });
    response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 account.</p>');
    response.end();
}

function authorize(response, request) {
    console.log('Request handler \'authorize\' was called.');
    // the authorization code is passed as a query parameter
    var url_parts = url.parse(request.url, true);
    var code = url_parts.query.code;
    console.log('Code: ' + code);
    authHelper.getTokenFromCode(code, tokenReceived, response);
}

function getValueFromCookie(valueName, cookie) {
    if (cookie.indexOf(valueName) !== -1) {
        var start = cookie.indexOf(valueName) + valueName.length + 1;
        var end = cookie.indexOf(';', start);
        end = end === -1 ? cookie.length : end;
        return cookie.substring(start, end);
    }
}

function getAccessToken(request, response, callback) {
    var expiration = new Date(parseFloat(getValueFromCookie('msgraph-insights-demo-token-expires', request.headers.cookie)));

    if (expiration <= new Date()) {
        // refresh token
        console.log('TOKEN EXPIRED, REFRESHING');
        var refresh_token = getValueFromCookie('msgraph-insights-demo-refresh-token', request.headers.cookie);
        authHelper.refreshAccessToken(refresh_token, function (error, newToken) {
            if (error) {
                callback(error, null);
            } else if (newToken) {
                var cookies = ['msgraph-insights-demo-token=' + newToken.token.access_token + ';Max-Age=4000',
                'msgraph-insights-demo-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                'msgraph-insights-demo-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
                response.setHeader('Set-Cookie', cookies);
                callback(null, newToken.token.access_token);
            }
        });
    } else {
        // return cached token
        var access_token = getValueFromCookie('msgraph-insights-demo-token', request.headers.cookie);
        callback(null, access_token);
    }
}

function tokenReceived(response, error, token) {
    if (error) {
        console.log('Access token error: ', error.message);
        response.writeHead(200, { 'Content-Type': 'text/html' });
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
    } else {
        getUserEmail(token.token.access_token, function (error, email) {
            if (error) {
                console.log('Getting insights returned an error: ' + error);
                response.write('<p>ERROR: ' + error + '</p>');
                response.end();
            } else if (email) {
                var cookies = ['msgraph-insights-demo-token=' + token.token.access_token + ';Max-Age=4000',
                'msgraph-insights-demo-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                'msgraph-insights-demo-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                'msgraph-insights-demo-email=' + email + ';Max-Age=4000'];
                response.setHeader('Set-Cookie', cookies);
                response.writeHead(302, { 'Location': 'http://localhost:8000/insights' });
                response.end();
            }
        });
    }
}

function getUserEmail(token, callback) {
    // create an msgraph client
    var client = microsoftGraph.Client.init({
        authProvider: (done) => {
            // just return the token
            done(null, token);
        }
    });

    // get the Graph /me endpoint to get user email address
    client
        .api('/me')
        .get()
        .then((res) => {
            callback(null, res.mail);
        })
        .catch((err) => {
            callback(err, null);
        });
}

function getLastUsed(response, request) {
    console.log('Request handler \'insights\' was called.');

    var token = getValueFromCookie('msgraph-insights-demo-token', request.headers.cookie);
    console.log('Token found in cookie: ', token);
    var email = getValueFromCookie('msgraph-insights-demo-email', request.headers.cookie);
    console.log('Email address found in cookie: ', email);

    if (token) {

        response.writeHead(200, { 'Content-Type': 'text/html' });
        response.write('<div><h1>Last used files</h1></div>');

        // create an ms msgraph client
        var client = microsoftGraph.Client.init({
            defaultVersion: 'beta',
            debugLogging: true,
            authProvider: (done) => {
                // just return the token
                done(null, token);
            }
        });

        // get list of used files
        client
            .api('/me/insights/used')
            .header('X-AnchorMailbox', email)
            .get()
            .then((res) => {
                console.log('lastused returned ' + res.value.length + ' files');
                console.log('lastused returned ' + JSON.stringify(res.value));
                response.write('<table><tr><th>title</th><th>last accessed</th><th>last modified</th></tr>');
                res.value.forEach(function (file) {
                    console.log('  Id: ' + file.id);
                    response.write('<tr><td>' + file.resourceVisualization.title +
                        '</td><td>' + file.resourceVisualization.type +
                        '</td><td>' + formatDate(new Date(file.lastUsed.lastModifiedDateTime)) +
                        '</td><td>' + formatDate(new Date(file.lastUsed.lastAccessedDateTime)) + '</td></tr>');
                });
                response.write('</table>');
                response.end();
            })
            .catch((err) => {
                console.log('lastused returned an error: ' + err);
                console.error(err);
                response.write('<p>ERROR: ' + err + '</p>');
                response.end();
            });
    } else {
        response.writeHead(200, { 'Content-Type': 'text/html' });
        response.write('<p>No token found in cookie!</p>');
        response.end();
    }
}

function formatDate(date) {
    var monthNames = [
        "January", "February", "March",
        "April", "May", "June", "July",
        "August", "September", "October",
        "November", "December"
    ];

    var day = date.getDate();
    var monthIndex = date.getMonth();
    var year = date.getFullYear();

    return day + ' ' + monthNames[monthIndex] + ' ' + year;
}
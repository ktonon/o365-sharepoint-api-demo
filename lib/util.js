var url = require('url'),
	request = require('request'),
	q = require('q'),
	querystring = require('querystring'),
	config = require('./config.json');

exports.haveSamePath = function(aUrl, bUrl) {
	var a = url.parse(aUrl),
		b = url.parse(bUrl),
		n = a.path.length;
	return a.path.slice(0, n) == b.path.slice(0, n);
};

exports.getAuthCode = function(redirectUri) {
	return querystring.parse(url.parse(redirectUri).query).code;
};

exports.chainableCall = function(name, requestDataFunc) {
	return function(data) {
		var requestData = requestDataFunc(data);
		console.log(name, requestData.method.toUpperCase(), requestData.url);
		var deferred = q.defer();
		request[requestData.method.toLowerCase()](requestData, function(err, resp, body) {
			if (err) {
				deferred.reject(err);
			}
			else if (resp.statusCode != 200) {
				deferred.reject('' + resp.statusCode + ': ' + body);
			}
			else {
				deferred.resolve(JSON.parse(body));
			}
		});
		return deferred.promise;
	}
};

exports.redirectUrl = function() {
	return 'http://localhost:' + config.port + '/authenticatereply';
};

exports.sharepointUrl = function() {
	return 'https://' + config.sharepointHost + '/';
};

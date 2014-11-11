var http = require('http'),
	open = require('open'),
	querystring = require('querystring'),
	_ = require('underscore'),

	util = require('./util.js'),
	config = require('./config.json');

var apiVersion = /^v?1(\.0)?$/.test(process.argv[2]) ? 'v1.0/' : ''

function authorize() {
	open('https://login.windows.net/common/oauth2/authorize?' +
		querystring.stringify({
			response_type: 'code',
			client_id: config.clientId,
			resource: 'Microsoft.SharePoint',
			redirect_uri: util.redirectUrl()
		}));
}

http.createServer(function(req, res) {
	
	var bucket = {};

	if (util.haveSamePath(util.redirectUrl(), req.url)) {
		util.chainableCall('Get token', function() {
			return {
				method: 'post',
				url: 'https://login.windows.net/common/oauth2/token',
				form: {
					client_id: config.clientId,
					client_secret: config.clientSecret,
					grant_type: 'authorization_code',
					code: util.getAuthCode(req.url),
					redirect_uri: util.redirectUrl()
				}
			};
		})()
		.then(util.chainableCall('Discover endpoints', function(data) {
			bucket.refresh_token = data.refresh_token;
			return {
				method: 'get',
				url: 'https://api.office.com/discovery/' + apiVersion + 'me/Services',
				secureOptions: require('constants').SSL_OP_NO_TLSv1_2,
				headers: {
					'Accept': 'application/json',
					'Authorization': 'Bearer ' + data.access_token,
					'Host': 'api.office.com',
				}
			}
		}))
		.then(util.chainableCall('Refresh token', function(data) {
			var service = _.find(data.value, function(s) {
				return (s.serviceId || s.ServiceId) == 'O365_SHAREPOINT';
			});
			bucket.endpoint = service.ServiceEndpointUri || service.serviceEndpointUri;
			return {
				method: 'post',
				url: 'https://login.windows.net/common/oauth2/token',
				form: {
					client_id: config.clientId,
					client_secret: config.clientSecret,
					grant_type: 'refresh_token',
					refresh_token: bucket.refresh_token,
					resource: util.sharepointUrl()
				}
			};
		}))
		.then(util.chainableCall('List files', function(data) {
			return {
				method: 'get',
				url: bucket.endpoint + '/files',
				secureOptions: require('constants').SSL_OP_NO_TLSv1_2,
				headers: {
					'Accept': 'application/json',
					'Authorization': 'Bearer ' + data.access_token,
					'Host': config.sharepointHost,
				}
			};
		}))
		.then(function(data) {
			res.writeHead(200, {'Content-Type': 'text/plain'});
			for (var i=0, n=data.value.length; i<n; ++i) {
				var doc = data.value[i];
				res.write(' - ' + (doc.name || doc.Name) + '\n');
			}
			res.end();
			process.exit();
		})
		.fail(function(reason) {
			res.writeHead(200, {'Content-Type': 'text/plain'});
			res.end(reason);
			process.exit();
		});
	}
	else {
		res.writeHead(404, {'Content-Type': 'text/plain'});
		res.end('404 - Not found');
		process.exit();
	}

}).listen(config.port, '127.0.0.1');

authorize();

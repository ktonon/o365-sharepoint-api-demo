Office 365 SharePoint API Demo
==============================

This is a short node script that demonstrates authentication with an Active Directory app configured with access to SharePoint. This script only works if read access to SharePoint files is configured on the app.

Instructions
------------

* Clone this repo
* Copy `lib/config.example.json` to `lib/config.json` and fill in the values (see next section)
* Run `npm install`
* Run `node lib/index.json`
* A browser will open. Authenticate and give the script access. After authenticating the script will:
  * Get an access token
  * Use the auth token to get the sharepoint api endpoint from the discovery service
  * Get a new access token given a refresh token
  * Use the new access token to get a list of files from sharepoint. The previously found endpoint is used

Running the script as

```
node lib/index.json
```

will use the preview discovery api. Running the script as

```
node lib/index.json 1.0
```

will use the 1.0 version of the discovery api.

config.json
-----------

* `clientId` - Active Directory app client id. Get this in the [Azure Management Portal][] on the Active Directory app's Configure page.
* `clientSecret` - Active Directory app secret key. Get this in the [Azure Management Portal][] on the Active Directory app's Configure page. You can only see the key immediately after creating it. If you forgot your key, create a new one.
* `port` - Port to use when running a local server. The local server responds to the redirect after authentication takes place. Your Active Directory app must have a reply url configured for this port. For example, if your port is 8000, then a single sign-on reply url must be `http://localhost:8000`. Set this value on the Active Directory app's Configure page.
* `sharepointHost` - sharepoint host name. To figure out this value, login to [Office 365][], and then navigate to OneDrive. Should look something like `some-subdomain.sharepoint.com`

Issues
------

Currently only the endpoint returned from the preview discovery service actually works. If the 1.0 discovery service is used, the sharepoint endpoint complains with this error:

```json
{
	"error": {
		"code": "-2147024891, System.UnauthorizedAccessException",
		"message": "Access denied. You do not have permission to perform this action or access this resource."
	}
}
```

[Azure Management Portal]:https://manage.windowsazure.com
[Office 365]:https://outlook.office365.com

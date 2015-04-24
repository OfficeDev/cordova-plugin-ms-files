Apache Cordova plugin for Files/Discovery Services API
=============================
Provides JavaScript API to work with Microsoft Files/Discovery Services API.
####Supported Platforms####

- Android (cordova-android@>=4.0.0 is supported)
- iOS
- Windows (Windows 8.0, Windows 8.1 and Windows Phone 8.1)

## Sample usage ##
To access the Files/Discovery API you need to acquire an access token and get the SharePoint client. Then, you can send async queries to interact with files data. Note: application ID, authorization and redirect URIs are assigned when you register your app with Microsoft Azure Active Directory.

```javascript
var resource = 'Microsoft.SharePoint';
var tenantId = '17bf7168-5251-44ed-a3cf-37a5997cc451';
var authority = 'https://login.windows.net/' + tenantId + '/';
var appId = '3cfa20df-bca4-4131-ab92-626fb800ebb5';
var redirectUrl = 'http://test.com';

var AuthenticationContext = Microsoft.ADAL.AuthenticationContext;
var DiscoveryServices = Microsoft.Office.Files.DiscoveryServices;
var SharePointClient = Microsoft.Office.Files.SharePointClient;

var authContext = new AuthenticationContext(authority);
var discoveryContext = new DiscoveryServices.Context(authContext, appId, redirectUrl);
var sharePointClient;

discoveryContext.services(resource).then(function (capabilities) {
    capabilities.forEach(function (v) {
        if (v.capability === 'MyFiles') {
            var msg;
            sharePointClient = SharePointClient(v.endpointUri, authContext,
                v.resourceId, appId, redirectUrl);

            sharePointClient.files.getFileSystemItems().fetch().then(function (result) {
                msg = '';
                result.currentPage.forEach(function (item) {
                    msg += item._odataType + ' "' + item.name + '"\n';
                });
                console.log('All file system items: \n' + msg);
            }, function (error) {
                console.error(error);
            });
        }
    });
}, function (error) {
    console.error(error);
});
```

Complete example is available [here](https://github.com/OfficeDev/cordova-plugin-ms-files/tree/master/sample).

## Installation Instructions ##

Use [Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) to create your app and add the plugin.

1. Make sure an up-to-date version of Node.js is installed, then type the following command to install the [Cordova CLI](https://github.com/apache/cordova-cli):

        npm install -g cordova

2. Create a project and add the platforms you want to support:

        cordova create sharepointClientApp
        cd sharepointClientApp
        cordova platform add windows <- support of Windows 8.0, Windows 8.1 and Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Add the plugin to your project:

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-files

4. Build and run, for example:

        cordova run windows

To learn more, read [Apache Cordova CLI Usage Guide](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Copyrights ##
Copyright (c) Microsoft Open Technologies, Inc. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use these files except in compliance with the License. You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.

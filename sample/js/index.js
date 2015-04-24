
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var AuthenticationContext;

var TENANT_ID = '17bf7168-5251-44ed-a3cf-37a5997cc451';
var AUTHORITY = 'https://login.windows.net/' + TENANT_ID + '/';

var APP_ID = '3cfa20df-bca4-4131-ab92-626fb800ebb5';
var REDIRECT_URI = "http://test.com";

var SCOPES = 'MyFiles.Read MyFiles.Write';

var SERVICE_RESOURCE = 'Microsoft.SharePoint';

var USER_EMAIL = '';

var Deferred;
var authenticated = false;

var myFilesCapability, sharePointClient;

function pre(json) {
    return '<pre>' + JSON.stringify(json, null, 4) + '</pre>';
}

var app = {
    // Application Constructor
    initialize: function () {
        this.bindEvents();
    },
    // Bind Event Listeners
    //
    // Bind any events that are required on startup. Common events are:
    // 'load', 'deviceready', 'offline', and 'online'.
    bindEvents: function () {
        document.addEventListener('deviceready', app.onDeviceReady, false);

        document.getElementById('firstSignIn').addEventListener('click', app.firstSignIn);
        document.getElementById('acquireToken').addEventListener('click', app.acquireToken);
        document.getElementById('clear-tokencache').addEventListener('click', app.clearTokenCache);
        document.getElementById('get-all-fs-items').addEventListener('click', app.getAllFileSystemItems);
        document.getElementById('get-capabilities').addEventListener('click', app.getCapabilities);
        document.getElementById('download-file-by-id').addEventListener('click', app.downloadFileById);
        document.getElementById('download-file-by-name').addEventListener('click', app.downloadFileByName);
        document.getElementById('create-file').addEventListener('click', app.createFile);
        document.getElementById('create-folder').addEventListener('click', app.createFolder);
        document.getElementById('copy-file').addEventListener('click', app.copyFile);
        document.getElementById('delete-file').addEventListener('click', app.deleteFile);
        document.getElementById('move-file').addEventListener('click', app.moveFile);
        document.getElementById('upload-file-contents').addEventListener('click', app.uploadFileContents);

        function toggleMenu() {
            // menu must be always shown on desktop/tablet
            if (document.body.clientWidth > 480) return;
            var cl = document.body.classList;
            if (cl.contains('left-nav')) { cl.remove('left-nav'); }
            else { cl.add('left-nav'); }
        }

        document.getElementById('slide-menu-button').addEventListener('click', toggleMenu);
    },

    // deviceready Event Handler
    //
    // The scope of 'this' is the event. In order to call the 'receivedEvent'
    // function, we must explicitly call 'app.receivedEvent(...);'
    onDeviceReady: function () {
        // app.receivedEvent('deviceready');
        app.logArea = document.getElementById("log-area");
        app.log("Cordova initialized, 'deviceready' event was fired");
        AuthenticationContext = Microsoft.ADAL.AuthenticationContext;
        Deferred = cordova.require('cordova-plugin-ms-adal.utility').Utility.Deferred;

        app.authContext = new AuthenticationContext(AUTHORITY);
        app.discoveryContext = new Microsoft.Office.Files.DiscoveryServices.Context(app.authContext, APP_ID, REDIRECT_URI);
    },

    // Update DOM on a Received Event
    receivedEvent: function (id) {
        var parentElement = document.getElementById(id);
        var listeningElement = parentElement.querySelector('.listening');
        var receivedElement = parentElement.querySelector('.received');

        listeningElement.setAttribute('style', 'display:none;');
        receivedElement.setAttribute('style', 'display:block;');

        console.log('Received Event: ' + id);
    },

    // Helper methods
    log: function (message, isError) {
        isError ? console.error(message) : console.log(message);
        var logItem = document.createElement('li');
        logItem.classList.add("topcoat-list__item");
        isError && logItem.classList.add("error-item");
        var timestamp = '<span class="timestamp">' + new Date().toLocaleTimeString() + ': </span>';

        var errObj;
        if (!!message && !!message.response) {
            errObj = JSON.parse(message.response);

            if (!!errObj && !!errObj.error && !!errObj.error.message && !!errObj.error.message.value) {
                message = errObj.error.message.value;
            }
        }

        logItem.innerHTML = (timestamp + message);
        app.logArea.insertBefore(logItem, app.logArea.firstChild);
    },

    error: function (message) {
        app.log(message, true);
    },


    discoverCapabilities: function () {
        var deferral = new Deferred();

        app.discoveryContext.services(SERVICE_RESOURCE).then((function (capabilities) {
            deferral.resolve(capabilities);
        }).bind(this), function (error) {
            deferral.reject(error);
        });

        return deferral;
    },

    ensureSharePointClientCreated: function () {
        var deferral = new Deferred();

        if (typeof myFilesCapability !== 'undefined') {
            deferral.resolve();
        } else {
            app.discoverCapabilities().then(function (capabilities) {
                capabilities.forEach(function (v) {
                    if (v.capability === 'MyFiles') {
                        myFilesCapability = v;

                        sharePointClient = new Microsoft.Office.Files.SharePointClient(myFilesCapability.endpointUri, app.authContext, myFilesCapability.resourceId, APP_ID, REDIRECT_URI);

                        deferral.resolve();
                        return;
                    }
                });

                deferral.reject('MyFiles capability not found');
            }, function (err) {
                deferral.reject(err);
            });
        }

        return deferral;
    },

    ensureAuthenticatedWithUserEmailAndDoAction: function (action) {
        if (USER_EMAIL !== '' && !authenticated) {
            app.authenticate().then(function () {
                app.ensureSharePointClientCreated().then(function () {
                    action();
                });
            });
        } else {
            app.ensureSharePointClientCreated().then(function () {
                action();
            });
        }
    },

    findFileByName: function (fileName) {
        var deferral = new Deferred();

        sharePointClient.files.getFileSystemItems().fetchAll().then(function (items) {
            for (var j = 0; j < items.length; j++) {
                if (items[j].name === fileName) {
                    deferral.resolve(items[j]);
                    return;
                }
            }

            deferral.reject('"' + fileName + '" file not found. Create it first.');
        }, function (error) {
            app.error(error);
        });

        return deferral;
    },

    authenticate: function() {
        var deferral = new Deferred();

        app.authContext.acquireTokenAsync(SERVICE_RESOURCE, APP_ID, REDIRECT_URI, null, 'login_hint=' + USER_EMAIL).then(function (authResult) {
            app.log('Acquired token successfully: ' + pre(authResult));
            authenticated = true;
            deferral.resolve();
        }, function (err) {
            app.error("Failed to acquire token: " + pre(err ? { error: err.error, description: err.errorDescription } : ""));
            deferral.reject(err);
        });

        return deferral;
    },

    // User action handlers
    firstSignIn: function () {
        app.discoveryContext.firstSignIn(SCOPES, REDIRECT_URI).then(function (res) {
            // Logging endpoints and user info
            var msg = 'user_email: ' + res['user_email'] + '<br />';
            msg += 'account_type code: ' + res['account_type'] + '<br />';
            msg += 'account_type: ' + Microsoft.Office.Files.DiscoveryServices.AccountType[res['account_type']] + '<br />';
            msg += 'authorization_service: ' + res['authorization_service'] + '<br />';
            msg += 'token_service: ' + res['token_service'] + '<br />';
            msg += 'discovery_resource: ' + res['discovery_resource'] + '<br />';
            msg += 'discovery_service: ' + res['discovery_service'] + '<br />';
            app.log(msg);

            USER_EMAIL = res.user_email;
        }, function (err) {
            app.error('firstSignIn failed: ' + err);
        });
    },

    acquireToken: function () {
        app.authenticate();
    },

    clearTokenCache: function () {
        app.authContext.tokenCache.clear().then(function () {
            app.log("Cache cleaned up successfully.");
            authenticated = false;
        }, function (err) {
            app.error("Failed to clear token cache: " + pre(err ? { error: err.error, description: err.errorDescription } : ""));
        });
    },

    getCapabilities: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var msg = '';

            app.discoverCapabilities().then(function (capabilities) {
                capabilities.forEach(function (v) {
                    msg += 'Capability: ' + v.capability + ", Name: " + v.name + ', EndPointUri: ' + v.endpointUri + '<br />';
                });

                app.log('Capabilities: <br />' + msg);
            }, function (error) {
                app.error(error);
            });
        });
    },

    getAllFileSystemItems: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var msg = '';

            sharePointClient.files.getFileSystemItems().fetch().then(function (result) {
                result.currentPage.forEach(function (item) {
                    msg += item._odataType + ' "' + item.name + '"<br />';
                });
                app.log('All file system items: <br />' + msg);
            }, function (error) {
                console.log(error);
            });
        });
    },

    createFile: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileName = 'demo.txt';

            var file = new Microsoft.Office.Files.FileServices.File(null, null, {});
            file.name = fileName;

            sharePointClient.files.addFileSystemItem(file).then(function () {
                app.log('Created the file successfully');
            }, function (error) {
                app.error(error);
            });
        });
    },

    downloadFileByName: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileName = 'demo.txt';
            var msg = '';

            app.findFileByName(fileName).then(function (file) {
                file.download().then(function (content) {
                    msg = fileName + ' contents: <br />';
                    msg += content;
                    app.log(msg);
                }, function (error) {
                    app.error(error);
                });
            }, function (error) {
                app.error(error);
            });
        });
    },

    uploadFileContents: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileName = 'demo.txt';
            var contents = 'This is a demo file contents\n:)';

            app.findFileByName(fileName).then(function (sourceFile) {
                sourceFile.upload(contents).then(function () {
                    app.log('Uploaded the file contents successfully');
                }, function (error) {
                    app.error(error);
                });
            }, function (error) {
                app.error(error);
            });
        });
    },

    copyFile: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileName = 'demo.txt';
            var targetName = 'demo2.txt';
            var overwrite = true;

            app.findFileByName(fileName).then(function (sourceFile) {
                sourceFile.copyTo(targetName, overwrite).then(function () {
                    app.log('Copied the file successfully');
                }, function (error) {
                    app.error(error);
                });
            }, function (error) {
                app.error(JSON.stringify(error));
            });
        });
    },

    moveFile: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileName = 'demo.txt';
            var targetFileName = 'demoRenamed.txt';
            var overwrite = true;

            app.findFileByName(fileName).then(function (sourceFile) {
                sourceFile.moveTo(targetFileName, overwrite).then(function () {
                    app.log('Moved the file successfully');
                }, function(error) {
                    app.error(error);
                });
            }, function (error) {
                app.error(error);
            });
        });
    },

    deleteFile: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileName = 'demo.txt';

            app.findFileByName(fileName).then(function (sourceFile) {
                sourceFile.delete().then(function (res) {
                    app.log('Deleted the file successfully');
                }, function (error) {
                    app.error(error);
                });
            }, function (error) {
                app.error(error);
            });
        });
    },

    downloadFileById: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var fileId;
            var msg = '';

            sharePointClient.files.getFileSystemItems().fetchAll().then(function (items) {
                for (var j = 0; j < items.length; j++) {
                    if (items[j]._odataType === 'MS.FileServices.File') {
                        fileId = items[j].id;
                        break;
                    }
                }

                if (typeof fileId === 'undefined') {
                    app.error('No files found in the file system. Please add one.');
                    return;
                }

                sharePointClient.files.getById(fileId).then(function(result) {
                    result.download().then(function(content) {
                        msg = result.name + ' contents: <br />';
                        msg += content;
                        app.log(msg);
                    }, function(error) {
                        app.error(error);
                    });
                }, function(error) {
                    app.error(error);
                });
            }, function(error) {
                app.error(error);
            });
        });
    },

    createFolder: function () {
        app.ensureAuthenticatedWithUserEmailAndDoAction(function () {
            var name = 'Demo';

            var folder = new Microsoft.Office.Files.FileServices.Folder(null, null, {});
            folder.name = name;

            sharePointClient.files.addFileSystemItem(folder).then(function () {
                app.log('Created the folder successfully');
            }, function (error) {
                app.error(error);
            });
        });
    }
};

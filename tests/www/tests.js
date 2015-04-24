
// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

/* global exports, require, Microsoft, jasmine, describe, it, xit, expect, beforeEach, afterEach */

var TENANT_ID = '17bf7168-5251-44ed-a3cf-37a5997cc451';
var AUTH_URL = 'https://login.windows.net/' + TENANT_ID + '/';

var RESOURCE_URL = 'https://testlaboratory-my.sharepoint.com/';
var APP_ID = '3cfa20df-bca4-4131-ab92-626fb800ebb5';
var REDIRECT_URL = "http://test.com";

var SERVICE_RESOURCE = 'Microsoft.SharePoint';
var TEST_USER_ID = '';

var AuthenticationContext = Microsoft.ADAL.AuthenticationContext;
var Deferred = require('cordova-plugin-ms-adal.utility').Utility.Deferred;

var TEST_FILE_NAME = 'testFile.txt';
var TEST_FOLDER_NAME = 'TestFolder';
var TEST_CONTENTS = 'This is a test contents';

var FILE_NOT_FOUND_STATUS = 404;
var FILE_NOT_FOUND_BY_NAME_MSG = 'File not found by name';

var TEST_FILE_COPY_TARGET = 'testFileCopied.txt';
var TEST_FILE_MOVE_TARGET = 'testFileMoved.txt';

var Folder = Microsoft.Office.Files.FileServices.Folder;
var File = Microsoft.Office.Files.FileServices.File;

exports.defineAutoTests = function () {
    jasmine.DEFAULT_TIMEOUT_INTERVAL = 30000;

    function fail(done, err) {
        expect(err).toBeUndefined();
        if (err !== null) {
            if (err.responseText !== null) {
                expect(err.responseText).toBeUndefined();
                console.error('Error: ' + err.responseText);
            } else {
                console.error('Error: ' + err);
            }
        }

        done();
    }

    function createAuthContext() {
        return new AuthenticationContext(AUTH_URL);
    }

    function createDiscoveryServicesContext(authContext) {
        authContext = authContext || createAuthContext();

        return new Microsoft.Office.Files.DiscoveryServices.Context(authContext, APP_ID, REDIRECT_URL);
    }

    function createSharePointClient(discoContext) {
        discoContext = discoContext || createDiscoveryServicesContext();

        var deferral = new Deferred();
        var myFilesCapability;
        var authContext = createAuthContext();

        discoContext.services(SERVICE_RESOURCE).then(function (capabilities) {
            capabilities.forEach(function (v) {
                if (v.capability === 'MyFiles') {
                    myFilesCapability = v;
                }
            });

            if (typeof myFilesCapability !== 'undefined') {
                deferral.resolve(new Microsoft.Office.Files.SharePointClient(myFilesCapability.endpointUri, authContext, myFilesCapability.resourceId, APP_ID, REDIRECT_URL));
            } else {
                deferral.reject('MyFiles capability is missing');
            }
        });

        return deferral;
    }

    describe('Login: ', function () {
        var authContext, backInterval;
        beforeEach(function () {
            authContext = createAuthContext();

            // increase standart jasmine timeout so that user can login
            backInterval = jasmine.DEFAULT_TIMEOUT_INTERVAL;
            jasmine.DEFAULT_TIMEOUT_INTERVAL = 120000;
        });

        afterEach(function () {
            // revert back default jasmine timeout
            jasmine.DEFAULT_TIMEOUT_INTERVAL = backInterval;
        });

        it("login.spec.1 should login first", function (done) {
            authContext.acquireTokenSilentAsync(RESOURCE_URL, APP_ID, TEST_USER_ID).then(function (authResult) {
                console.log("Token is: " + authResult.accessToken);
                expect(authResult).toBeDefined();
                done();
            }, function () {
                console.warn("You should login in the manual tests first");

                authContext.acquireTokenAsync(RESOURCE_URL, APP_ID, REDIRECT_URL).then(function (authResult) {
                    console.log("Token is: " + authResult.accessToken);
                    expect(authResult).toBeDefined();
                    done();
                }, function (err) {
                    console.error(err);
                    expect(err).toBeUndefined();
                    done();
                });
            });
        });
    });

    describe('SharePoint client: ', function () {
        var authContext, client;

        beforeEach(function (done) {
            authContext = createAuthContext();
            createSharePointClient().then(function (sharePointClient) {
                client = sharePointClient;
                done();
            }, fail.bind(this, done));
        });

        it('client.spec.1 SharePoint client should exists', function () {
            expect(Microsoft.Office.Files.SharePointClient).toBeDefined();
            expect(Microsoft.Office.Files.SharePointClient).toEqual(jasmine.any(Function));
        });

        it('client.spec.2 should be able to create a new client', function (done) {
            expect(client).not.toBe(null);
            expect(client.context).toBeDefined();
            expect(client.context.serviceRootUri).toBeDefined();
            //expect(client.context._getAccessTokenFn).toBeDefined();
            //expect(client.context.serviceRootUri).toEqual(ENDPOINT_URL);
            //expect(client.context._getAccessTokenFn).toEqual(jasmine.any(Function));
            expect(client.files).toBeDefined();

            done();
        });

        it('client.spec.3 should contain \'files\' property', function (done) {
            expect(client.files).toBeDefined();
            expect(client.files).toEqual(jasmine.any(Microsoft.Office.Files.FileServices.FileSystemItems));

            // expect that client.directoryObjects is readonly
            var backup = client.files;
            client.files = "somevalue";
            expect(client.files).not.toEqual("somevalue");
            expect(client.files).toEqual(backup);

            done();
        });
    });

    describe('Discovery services API: ', function () {
        beforeEach(function () {
            var that = this;

            that.discoveryContext = createDiscoveryServicesContext();
            that.tempEntities = [];

            that.runSafely = function runSafely(testFunc, localDone) {
                try {
                    // Wrapping the call into try/catch to avoid test suite crashes and `hanging` test entities
                    testFunc(localDone);
                } catch (err) {
                    fail.call(that, localDone, err);
                }
            };
        });

        afterEach(function (done) {
            var removedEntitiesCount = 0;
            var entitiesToRemoveCount = this.tempEntities.length;

            if (entitiesToRemoveCount === 0) {
                done();
            } else {
                this.tempEntities.forEach(function (entity) {
                    try {
                        entity.delete().then(function () {
                            removedEntitiesCount++;
                            if (removedEntitiesCount === entitiesToRemoveCount) {
                                done();
                            }
                        }, function (err) {
                            expect(err).toBeUndefined();
                            done();
                        });
                    } catch (e) {
                        expect(e).toBeUndefined();
                        done();
                    }
                });
            }
        });

        it("discovery.spec.1 Should be able to get services (capabilities)", function (done) {
            var that = this;

            that.runSafely(function (localDone) {
                expect(that.discoveryContext).toBeDefined();
                expect(that.discoveryContext.services).toEqual(jasmine.any(Function));
                that.discoveryContext.services(SERVICE_RESOURCE).then(function (capabilities) {
                    expect(capabilities).toEqual(jasmine.any(Array));
                    expect(capabilities[0]).toEqual(jasmine.any(Microsoft.Office.Files.DiscoveryServices.ServiceCapability));

                    localDone();
                }, fail.bind(that, localDone));
            }, done);
        });
    });

    describe('SharePoint API: ', function () {
        beforeEach(function (done) {
            var that = this;

            that.discoveryContext = createDiscoveryServicesContext();
            createSharePointClient(that.discoveryContext).then(function(client) {
                that.sharePointClient = client;

                that.tempEntities = [];

                that.runSafely = function runSafely(testFunc, localDone) {
                    try {
                        // Wrapping the call into try/catch to avoid test suite crashes and `hanging` test entities
                        testFunc(localDone);
                    } catch (err) {
                        fail.call(that, localDone, err);
                    }
                };

                that.findFileByName = function findFileByName(name) {
                    var deferral = new Deferred();

                    that.sharePointClient.files.getFileSystemItems().fetchAll().then(function(items) {
                        for (var j = 0; j < items.length; j++) {
                            if (items[j].name === name) {
                                deferral.resolve(items[j]);
                                return;
                            }
                        }

                        deferral.reject(FILE_NOT_FOUND_BY_NAME_MSG);
                    }, function(err) {
                        deferral.reject(err);
                    });

                    return deferral;
                };

                done();
            }, fail.bind(that, done));
        });

        afterEach(function (done) {
            var removedEntitiesCount = 0;
            var entitiesToRemoveCount = this.tempEntities.length;

            if (entitiesToRemoveCount === 0) {
                done();
            } else {
                this.tempEntities.forEach(function (entity) {
                    try {
                        entity.delete().then(function () {
                            removedEntitiesCount++;
                            if (removedEntitiesCount === entitiesToRemoveCount) {
                                done();
                            }
                        }, function (err) {
                            expect(err).toBeUndefined();

                            if (err !== null) {
                                if (err.responseText !== null) {
                                    expect(err.responseText).toBeUndefined();
                                    console.error('Error: ' + err.responseText);
                                } else {
                                    console.error('Error: ' + err);
                                }
                            }

                            done();
                        });
                    } catch (err) {
                        expect(err).toBeUndefined();

                        if (err !== null) {
                            if (err.responseText !== null) {
                                expect(err.responseText).toBeUndefined();
                                console.error('Error: ' + err.responseText);
                            } else {
                                console.error('Error: ' + err);
                            }
                        }

                        done();
                    }
                });
            }
        });

        describe('FileSystemItems: ', function() {
            it("FileSystemItems.spec.1 Should be able to get all FS items", function (done) {
                var that = this;

                that.runSafely(function (localDone) {
                    that.sharePointClient.files.getFileSystemItems().fetchAll().then(function (items) {
                        expect(items).toEqual(jasmine.any(Array));
                        expect(items[0]._odataType === 'MS.FileServices.Folder' || items[0]._odataType === 'MS.FileServices.File').toBeTruthy();

                        localDone();
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("FileSystemItems.spec.2 Should be able to get an item by id", function (done) {
                var that = this;

                that.runSafely(function (localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function (result) {
                        that.tempEntities.push(result);
                    
                        that.sharePointClient.files.getById(result.id).then(function (original) {
                            expect(original).toEqual(jasmine.any(File));

                            localDone();
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });
        });

        describe('File: ', function() {
            it("File.spec.1 Should be able to create a file", function (done) {
                var that = this;

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);
                        expect(result).toEqual(jasmine.any(File));
                        expect(result.name).toEqual(TEST_FILE_NAME);

                        localDone();
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("File.spec.2 Should be able upload a file contents", function (done) {
                var that = this;

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);

                        result.upload(TEST_CONTENTS).then(function() {
                            that.sharePointClient.files.getById(result.id).then(function(modified) {
                                expect(modified).toEqual(jasmine.any(File));

                                modified.download().then(function(content) {
                                    expect(content).toEqual(TEST_CONTENTS);

                                    localDone();
                                }, fail.bind(that, localDone));
                            }, fail.bind(that, localDone));
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("File.spec.3 Should be able download a file contents", function (done) {
                var that = this;

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);

                        result.download().then(function(content) {
                            expect(content).toBeDefined();

                            localDone();
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            // MS.FileServices.File does not support PATCH method
            xit("File.spec.4 Should be able update a file properties", function (done) {
                var that = this;
                var updatedFileName = "updatedTestFileName.txt";

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);

                        result.name = updatedFileName;

                        result.update().then(function(updatedFile) {
                            expect(updatedFile).toEqual(jasmine.any(File));
                            expect(updatedFile.name).toEqual(updatedFileName);

                            localDone();
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("File.spec.5 Should be able to delete a file", function (done) {
                var that = this;

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);

                        result.delete().then(function() {
                            that.sharePointClient.files.getById(result.id).then(function (deleted) {
                                expect(deleted).not.toBeDefined();

                                localDone();
                            }, function(err) {
                                expect(err.status).toEqual(FILE_NOT_FOUND_STATUS);
                                that.tempEntities.pop();

                                localDone();
                            });
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("File.spec.6 Should be able to copy a file", function (done) {
                var that = this;

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);

                        result.copyTo(TEST_FILE_COPY_TARGET, true).then(function() {
                            that.sharePointClient.files.getById(result.id).then(function(original) {
                                expect(original).toEqual(jasmine.any(File));
                                expect(original.name).toEqual(TEST_FILE_NAME);

                                that.findFileByName(TEST_FILE_COPY_TARGET).then(function(copied) {
                                    that.tempEntities.push(copied);
                                    expect(copied).toEqual(jasmine.any(File));
                                    expect(copied.name).toEqual(TEST_FILE_COPY_TARGET);

                                    localDone();
                                }, fail.bind(that, localDone));
                            }, fail.bind(that, localDone));
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("File.spec.7 Should be able to move a file", function (done) {
                var that = this;

                that.runSafely(function(localDone) {
                    var file = new File(null, null, {});
                    file.name = TEST_FILE_NAME;

                    that.sharePointClient.files.addFileSystemItem(file).then(function(result) {
                        that.tempEntities.push(result);

                        result.upload(TEST_CONTENTS).then(function() {
                            result.moveTo(TEST_FILE_MOVE_TARGET, true).then(function() {
                                that.findFileByName(TEST_FILE_NAME).then(function(original) {
                                    expect(original).not.toBeDefined();
                                    localDone();
                                }, function(err) {
                                    expect(err).toEqual(FILE_NOT_FOUND_BY_NAME_MSG);
                                    that.tempEntities.pop();

                                    that.findFileByName(TEST_FILE_MOVE_TARGET).then(function(moved) {
                                        that.tempEntities.push(moved);
                                        expect(moved).toEqual(jasmine.any(File));
                                        expect(moved.name).toEqual(TEST_FILE_MOVE_TARGET);

                                        moved.download().then(function(content) {
                                            expect(content).toEqual(TEST_CONTENTS);

                                            localDone();
                                        }, fail.bind(that, localDone));
                                    }, fail.bind(that, localDone));
                                }, fail.bind(that, localDone));
                            }, fail.bind(that, localDone));
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });
        });

        describe('Folder: ', function() {
            it("Folder.spec.1 Should be able to create a folder", function (done) {
                var that = this;

                that.runSafely(function (localDone) {
                    var folder = new Folder(null, null, {});
                    folder.name = TEST_FOLDER_NAME;

                    that.sharePointClient.files.addFileSystemItem(folder).then(function (result) {
                        that.tempEntities.push(result);
                        expect(result).toEqual(jasmine.any(Folder));
                        expect(result.name).toEqual(TEST_FOLDER_NAME);

                        localDone();
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("Folder.spec.2 Should be able to delete a folder", function (done) {
                var that = this;

                that.runSafely(function (localDone) {
                    var folder = new Folder(null, null, {});
                    folder.name = TEST_FOLDER_NAME;

                    that.sharePointClient.files.addFileSystemItem(folder).then(function (result) {
                        that.tempEntities.push(result);

                        result.delete().then(function () {
                            that.sharePointClient.files.getById(result.id).then(function (deleted) {
                                expect(deleted).not.toBeDefined();

                                localDone();
                            }, function (err) {
                                expect(err.status).toEqual(FILE_NOT_FOUND_STATUS);
                                that.tempEntities.pop();

                                localDone();
                            });
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("Folder.spec.3 Should be able to get a folder children", function (done) {
                var that = this;

                that.runSafely(function (localDone) {
                    var folder = new Folder(null, null, {});
                    folder.name = TEST_FOLDER_NAME;

                    that.sharePointClient.files.addFileSystemItem(folder).then(function (createdFolder) {
                        that.tempEntities.push(createdFolder);

                        var file = new File(null, null, {});
                        file.name = TEST_FOLDER_NAME + '/' + TEST_FILE_NAME;

                        that.sharePointClient.files.addFileSystemItem(file).then(function(addedFile) {
                            createdFolder.children.getFileSystemItems().fetchAll().then(function(children) {
                                expect(children).toEqual(jasmine.any(Array));
                                expect(children[0]).toEqual(jasmine.any(File));
                                expect(children[0].id).toEqual(addedFile.id);

                                localDone();
                            }, fail.bind(that, localDone));
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });

            it("Folder.spec.4 Should be able to get a folder children count", function (done) {
                var that = this;

                that.runSafely(function (localDone) {
                    var folder = new Folder(null, null, {});
                    folder.name = TEST_FOLDER_NAME;

                    that.sharePointClient.files.addFileSystemItem(folder).then(function (createdFolder) {
                        that.tempEntities.push(createdFolder);

                        var file = new File(null, null, {});
                        file.name = TEST_FOLDER_NAME + '/' + TEST_FILE_NAME;

                        that.sharePointClient.files.addFileSystemItem(file).then(function () {
                            that.sharePointClient.files.getById(createdFolder.id).then(function (createdFolderUpd) {
                                expect(createdFolderUpd.childrenCount).toEqual(1);

                                localDone();
                            }, fail.bind(that, localDone));
                        }, fail.bind(that, localDone));
                    }, fail.bind(that, localDone));
                }, done);
            });
        });
    });
};

exports.defineManualTests = function (contentEl, createActionButton) {
    var authContext;

    createActionButton('Log in', function () {
        authContext = new AuthenticationContext(AUTH_URL);
        authContext.acquireTokenAsync(RESOURCE_URL, APP_ID, REDIRECT_URL).then(function (authRes) {
            // Save acquired userId for further usage
            TEST_USER_ID = authRes.userInfo && authRes.userInfo.userId;

            console.log("Token is: " + authRes.accessToken);
            console.log("TEST_USER_ID is: " + TEST_USER_ID);
        }, function (err) {
            console.error(err);
        });
    });

    createActionButton('Log out', function () {
        authContext = authContext || new AuthenticationContext(AUTH_URL);
        return authContext.tokenCache.clear().then(function () {
            console.log("Logged out");
        }, function (err) {
            console.error(err);
        });
    });
};

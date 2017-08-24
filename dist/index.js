#!/usr/bin/env node
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var program = require("Commander");
var node_sp_auth_1 = require("node-sp-auth");
var request = require("request-promise");
var preferences = require('preferences');
program
    .version('0.1.0')
    .option('-c, --connect <siteurl>', 'Connect to SharePoint Online', null)
    .option('-w, --web', 'Show extentions at the web level')
    .option('-s, --sitecollection', 'Show extentions at the site collection level')
    .option('-l, --list <list title>', 'Show extentions at the list level')
    .parse(process.argv);
var prefs = new preferences('vman.sp.extentions.cli', {
    siteUrl: '',
    authHeaders: null
});
if (program.connect) {
    prefs.siteUrl = program.connect;
    node_sp_auth_1.getAuth(prefs.siteUrl, {
        ondemand: true,
        electron: require('electron'),
        force: false,
        persist: true
    }).then(function (authResponse) {
        prefs.authHeaders = authResponse.headers;
        prefs.authHeaders.Accept = 'application/json;odata=nometadata';
    });
}
if (program.web) {
    displayExtentions('web');
}
if (program.sitecollection) {
    displayExtentions('site');
}
if (program.list) {
    displayExtentions("web/lists/GetByTitle('" + program.list + "')");
}
function displayExtentions(path) {
    if (!prefs.siteUrl) {
        console.error('Please use --connect <siteurl');
    }
    request.get({
        url: prefs.siteUrl + "/_api/" + path + "/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'",
        headers: prefs.authHeaders
    }).then(function (response) {
        var userCustomActions = JSON.parse(response).value;
        for (var _i = 0, userCustomActions_1 = userCustomActions; _i < userCustomActions_1.length; _i++) {
            var uca = userCustomActions_1[_i];
            console.log("Title: " + uca.Title + ",\n                   ClientSideComponentId: " + uca.ClientSideComponentId + ",\n                   ClientSideComponentProperties: " + uca.ClientSideComponentProperties + ",\n                   Location: " + uca.Location);
        }
    }, function (error) {
        console.error(error.message);
    });
}

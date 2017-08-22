#!/usr/bin/env node
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var program = require("Commander");
var node_sp_auth_1 = require("node-sp-auth");
var request = require("request-promise");
program
    .version('0.1.0')
    .option('-p, --permissions', 'Show current user permissions')
    .parse(process.argv);
var authHeaders;
var siteUrl;
/* Need preferences js */
if (program.connect) {
    siteUrl = program.connect;
    node_sp_auth_1.getAuth(siteUrl, {
        ondemand: true,
        electron: require('electron'),
        force: false,
        persist: true
    }).then(function (authResponse) {
        authHeaders = authResponse.headers;
        authHeaders.Accept = 'application/json;odata=nometadata';
    });
}
if (program.permissions) {
    request.get({
        url: program.siteurl + "/_api/SP.Utilities.Utility.GetUserPermissionLevels",
        headers: authHeaders
    }).then(function (response) {
        console.log(response);
    });
}

#!/usr/bin/env node
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var program = require("Commander");
var node_sp_auth_1 = require("node-sp-auth");
var sp_pnp_js_1 = require("sp-pnp-js");
var node_pnp_js_1 = require("node-pnp-js");
program
    .version('0.1.0')
    .option('-c, --connect', 'Connect to SharePoint Online')
    .option('-s, --siteurl [siteurl]', 'SharePoint Online Site Url [null]', null)
    .parse(process.argv);
if (program.connect) {
    console.log(program.siteurl);
    if (!program.siteurl) {
        console.error('Please enter siteurl with --siteurl or -s');
        process.exit();
    }
    var siteUrl_1 = program.siteurl;
    node_sp_auth_1.getAuth(siteUrl_1, {
        ondemand: true,
        electron: require('electron'),
        force: false,
        persist: true
    }).then(function () {
        sp_pnp_js_1.setup({
            fetchClientFactory: function () {
                return new node_pnp_js_1.default({
                    ondemand: true
                });
            },
            baseUrl: siteUrl_1
        });
    }).then(function () {
        // we need to use the Web constructor to ensure we have the absolute url
        var web = new sp_pnp_js_1.Web(siteUrl_1);
        web.select('Title').get().then(function (w) {
            console.log("Web's title: " + w.Title);
        });
    });
}
// if (program.pineapple) console.log('  - pineapple');
// if (program.bbqSauce) console.log('  - bbq');
// console.log('  - %s cheese', program.cheese); 
#!/usr/bin/env node
import * as program from 'Commander';
import { getAuth, IAuthResponse } from 'node-sp-auth';
import * as request from 'request-promise';

program
  .version('0.1.0')
  //.option('-c, --connect <siteurl>', 'Connect to SharePoint Online', null)
  .option('-p, --permissions', 'Show current user permissions')
  .parse(process.argv);

let authHeaders: any;
let siteUrl: string;

/* Need preferences js */

if (program.connect) {

  siteUrl = program.connect;
  getAuth(siteUrl, {
    ondemand: true,
    electron: require('electron'),
    force: false,
    persist: true
  }).then((authResponse: IAuthResponse) => {

    authHeaders = authResponse.headers;
    authHeaders.Accept = 'application/json;odata=nometadata';
  });
}

if (program.permissions) {

  request.get({
    url: `${program.siteurl}/_api/SP.Utilities.Utility.GetUserPermissionLevels`,
    headers: authHeaders
  }).then((response) => {
    console.log(response);
  });
}
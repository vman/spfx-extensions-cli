#!/usr/bin/env node
import * as program from 'Commander';
import { getAuth, IAuthResponse } from 'node-sp-auth';
import * as request from 'request-promise';
const Preferences: any = require('preferences');
import { IExtention, IFieldCustomizer } from './interfaces';
const colors = require('colors/safe');

program
  .version('0.1.0')
  .option('-c, --connect <siteurl>', 'Connect to SharePoint Online', null)
  .option('-w, --web', 'Show extentions at the web level')
  .option('-s, --sitecollection', 'Show extentions at the site collection level')
  .option('-l, --list <list title>', 'Show extentions at the list level')
  .parse(process.argv);

const prefs = new Preferences('vman.sp.extentions.cli', {
  siteUrl: '',
  authHeaders: null
});

if (program.connect) {
  prefs.siteUrl = program.connect;
  getAuth(prefs.siteUrl, {
    ondemand: true,
    electron: require('electron'),
    force: false,
    persist: true
  }).then((authResponse: IAuthResponse) => {

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
  displayListExtentions();
}

async function displayExtentions(scope: string) {
  try {
    ensureAuth();

    const restUrl: string = `${prefs.siteUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location, 'ClientSideExtension')
    &$select=ClientSideComponentId,Title,Location,ClientSideComponentProperties`;
    const extention: IExtention[] = await fetchExtentions(restUrl);

    //add cli-table back in
    console.log(colors.magenta(`'${scope}' level spfx extentions at '${prefs.siteUrl}'`));
    console.log(colors.yellow('Title, ClientSideComponentId, Location, ClientSideComponentProperties'));
    for (const ext of extention) {
      console.log(colors.green([ext.Title, ext.ClientSideComponentId, ext.Location, ext.ClientSideComponentProperties].join(', ')));
    }

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

function ensureAuth() {
  if (!prefs.siteUrl) {
    throw new Error('Please use --connect <siteurl> to auth with SPO. Type --help for help.');
  }
}

async function fetchExtentions(restUrl: string) {
  const response: any = await request.get({
    url: restUrl,
    headers: prefs.authHeaders
  });
  return JSON.parse(response).value;
}

async function displayListExtentions() {
  try {
    ensureAuth();
    const restUrl: string = `${prefs.siteUrl}/_api/web/1lists/GetByTitle('${program.list}')/fields?
  $select=Title,ClientSideComponentId,ClientSideComponentProperties`;

    const fieldCustomizers: IFieldCustomizer[] = (await fetchExtentions(restUrl) as any[])
      .filter((field) => field.ClientSideComponentId !== '00000000-0000-0000-0000-000000000000');

    console.log(colors.magenta(`field customizer spfx extentions on '${program.list}' at '${prefs.siteUrl}'`));
    console.log(colors.yellow('Title, ClientSideComponentId, ClientSideComponentProperties'));
    for (const fc of fieldCustomizers) {
      console.log(colors.green([fc.Title, fc.ClientSideComponentId, fc.ClientSideComponentProperties].join(', ')));
    }
  } catch (error) {
    console.log(colors.red(error.message));
  }
}
#!/usr/bin/env node
import * as program from 'Commander';
import { getAuth, IAuthResponse } from 'node-sp-auth';
import * as request from 'request-promise';
import { IExtension } from './interfaces';
import { ExtensionScope, RegistrationType } from './enums';

const Preferences = require('preferences');
const colors = require('colors/safe');
const pjson = require('../package.json');
const Table = require('easy-table');

const prefs = new Preferences('vman.spfx.extensions.cli', {
  siteUrl: '',
  authHeaders: null
});

program
  .version(pjson.version)
  .option('-c, --connect <siteurl>', 'Connect to SharePoint Online at <siteurl>', null)
  .option('-w, --web', 'Show extensions at the web level')
  .option('-s, --sitecollection', 'Show extensions at the site collection level')
  .option('-l, --list <listtitle>', 'Show extensions at the list level for <listtitle>');

program
  .command('add <title> <extensionType> <scope> <clientSideComponentId>')
  .action(addExtension)
  .option('-rid, --registrationid', 'Only required if extention type is ListViewCommandSet')
  .option('-rtype, --registrationType', 'Only required if extention type is ListViewCommandSet (List | ContentType)', RegistrationType.None)
  .option('-cprops, --clientprops <json>', 'properties to add to the extension in json format')
  .on('--help', () => {
    console.log('');
    console.log('Required arguments:');
    console.log('<title> of the extension');
    console.log('<extensionType> of the extension (ApplicationCustomizer | ListViewCommandSet | ListViewCommandSet.CommandBar | ListViewCommandSet.ContextMenu)');
    console.log('<scope> Scope at which to add the extension (sitecollection | web )');
    console.log('<clientSideComponentId> from the manifest.json file of the extension');
    console.log('');
  });

program
  .command('remove <scope> <id>')
  .action(removeExtension)
  .on('--help', () => {
    console.log('');
    console.log('<scope> Scope from which to remove the extension (sitecollection | web )');
    console.log('<id> of the extension');
    console.log('');
  });

program.parse(process.argv);

if (program.connect) {
  prefs.siteUrl = program.connect;
  getAuth(prefs.siteUrl, {
    ondemand: true,
    electron: require('electron'),
    force: false,
    persist: true
  }).then((authResponse: IAuthResponse) => {

    prefs.authHeaders = authResponse.headers;
    prefs.authHeaders = { ...prefs.authHeaders, Accept: 'application/json;odata=nometadata' };
  });
}

if (program.web) {
  displayExtensions(ExtensionScope.Web);
}

if (program.sitecollection) {
  displayExtensions(ExtensionScope.SiteCollection);
}

if (program.list) {
  displayListExtensions();
}

async function removeExtension(scope: ExtensionScope, id: string) {
  try {

    ensureAuth();
    const path = `${scope === ExtensionScope.Web ? ExtensionScope.Web : ExtensionScope.SiteCollection}/UserCustomActions('${id}')`;
    const userCustomActionUrl: string = `${prefs.siteUrl}/_api/${path}`;

    console.log(await postExtension(userCustomActionUrl, undefined, 'DELETE'));

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

async function addExtension(title: string, extensionType: string, scope: string, clientSideComponentId: string) {

  try {

    ensureAuth();
    const path = `${scope === ExtensionScope.Web ? ExtensionScope.Web : ExtensionScope.SiteCollection}/UserCustomActions`;
    const userCustomActionUrl: string = `${prefs.siteUrl}/_api/${path}`;
    const requestBody: string = JSON.stringify({
      Title: title,
      Location: `ClientSideExtension.${extensionType}`,
      ClientSideComponentId: clientSideComponentId,
      ClientSideComponentProperties: program.clientprops ? program.clientprops : null,
      RegistrationId: program.registrationid ? program.registrationid : null,
      RegistrationType: RegistrationType[program.registrationType]
    });

    console.log(await postExtension(userCustomActionUrl, requestBody));

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

async function displayExtensions(scope: ExtensionScope) {
  try {
    ensureAuth();

    const userCustomActionUrl: string = `${prefs.siteUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location, 'ClientSideExtension')&$select=Id,ClientSideComponentId,Title,Location,ClientSideComponentProperties`;

    const fieldsPath: string = (scope === ExtensionScope.Web) ? 'fields' : 'rootWeb/availablefields';
    const fieldCustomizerUrl: string = `${prefs.siteUrl}/_api/${scope}/${fieldsPath}?$select=Id,ClientSideComponentId,Title,ClientSideComponentProperties`;

    const [exts, fields] = await Promise.all([getExtensions(userCustomActionUrl), getExtensions(fieldCustomizerUrl)]);

    const siteExtensions: IExtension[] = exts as IExtension[];
    const fieldCustomizers: IExtension[] = getFieldCustomizers(fields as IExtension[]);

    const extensions = siteExtensions.concat(fieldCustomizers);

    console.log('');
    console.log(colors.magenta(`'${scope}' level spfx extensions at '${prefs.siteUrl}'`));
    console.log('');
    printToConsole(extensions);

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

function printToConsole(extensions: IExtension[]) {
  const t = new Table();
  extensions.forEach((extention: IExtension) => {
    t.cell(colors.yellow('Id'), colors.green(extention.Id));
    t.cell(colors.yellow('Title'), colors.green(extention.Title));
    t.cell(colors.yellow('ClientSideComponentId'), colors.green(extention.ClientSideComponentId));
    t.cell(colors.yellow('Location'), colors.green(extention.Location));
    t.cell(colors.yellow('ClientSideComponentProperties'), colors.green(extention.ClientSideComponentProperties));
    t.newRow();
  });
  console.log(t.toString());
}

function getFieldCustomizers(fields: IExtension[]) {
  return fields
    .filter((field) => field.ClientSideComponentId !== '00000000-0000-0000-0000-000000000000')
    .map((fieldCustomizer) => {
      fieldCustomizer.Location = 'FieldCustomizer';
      return fieldCustomizer;
    });
}

function ensureAuth() {
  if (!prefs.siteUrl) {
    throw new Error('Please use spfx-ext --connect <siteurl> to auth with SPO. Type spfx-ext --help for help.');
  }
}

async function getExtensions(restUrl: string) {
  const response: any = await request.get({
    url: restUrl,
    headers: prefs.authHeaders
  });
  return JSON.parse(response).value;
}

async function postExtension(restUrl: string, requestBody: string = '', requestMethod: string = 'POST') {
  const reqDigestResponse: any = await request.post({
    url: `${prefs.siteUrl}/_api/contextinfo`,
    headers: prefs.authHeaders
  });
  const requestDigest = JSON.parse(reqDigestResponse).FormDigestValue;
  let postHeaders = {
    ...prefs.authHeaders,
    'X-RequestDigest': requestDigest,
    'content-type': 'application/json;odata=nometadata'
  };

  if (requestMethod === 'DELETE') {
    postHeaders = { ...postHeaders, 'X-HTTP-Method': 'DELETE' };
  }

  const response: any = await request.post({
    url: restUrl,
    body: requestBody,
    method: requestMethod,
    headers: postHeaders
  });
  return response ? JSON.parse(response) : '';
}

async function displayListExtensions() {
  try {
    ensureAuth();
    const restUrl: string = `${prefs.siteUrl}/_api/web/lists/GetByTitle('${program.list}')/fields?$select=Title,ClientSideComponentId,ClientSideComponentProperties`;

    const extensions: IExtension[] = getFieldCustomizers(await getExtensions(restUrl) as any[]);

    console.log(colors.magenta(`FieldCustomizer spfx extensions on '${program.list}' at '${prefs.siteUrl}'`));

    printToConsole(extensions);

  } catch (error) {
    console.log(colors.red(error.message));
  }
}
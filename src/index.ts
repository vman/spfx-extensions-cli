#!/usr/bin/env node
import * as program from 'commander';
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
  .option('-s, --site', 'Show extensions at the site collection level')
  .option('-l, --list <listtitle>', 'Show extensions at the list level for <listtitle>');

program
  .command('add <title> <extensionType> <scope> <clientSideComponentId>')
  .action(addExtension)
  .option('-p, --clientProps <json>', 'properties to add to the extension in json format', '')
  .option('-lt, --listtitle <title>', 'Only required if scope is list', null)
  .option('-i, --registrationId <id>', 'Only required if extension type is ListViewCommandSet')
  .option('-t, --registrationType <type>', 'Only required if extension type is ListViewCommandSet (List | ContentType)')
  .on('--help', () => {
    console.log('');
    console.log('Required arguments:');
    console.log('<title> of the extension');
    console.log('<extensionType> of the extension (ApplicationCustomizer | ListViewCommandSet | ListViewCommandSet.CommandBar | ListViewCommandSet.ContextMenu)');
    console.log('<scope> Scope at which to add the extension (site | web | list)');
    console.log('<clientSideComponentId> from the manifest.json file of the extension');
    console.log('');
  });

program
  .command('remove <scope> <id>')
  .action(removeExtension)
  .option('-lt, --listtitle <title>', 'Only required if scope is list', null)
  .on('--help', () => {
    console.log('');
    console.log('<scope> Scope from which to remove the extension (site | web | list )');
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

if (program.site) {
  displayExtensions(ExtensionScope.Site);
}

if (program.list) {
  displayExtensions(ExtensionScope.List, program.list);
}

async function displayExtensions(scope: ExtensionScope, listtitle?: string) {
  try {
    ensureAuth();

    const resourcePath: string = (scope === ExtensionScope.List) ? `web/lists/GetByTitle('${listtitle}')` : scope;
    const userCustomActionUrl: string = `${prefs.siteUrl}/_api/${resourcePath}/UserCustomActions?$filter=startswith(Location, 'ClientSideExtension')&$select=Id,ClientSideComponentId,Title,Location,ClientSideComponentProperties`;

    const fieldsPath: string = (scope === ExtensionScope.Web || scope === ExtensionScope.List) ? 'fields' : 'rootWeb/availablefields';
    const fieldCustomizerUrl: string = `${prefs.siteUrl}/_api/${resourcePath}/${fieldsPath}?$select=Id,ClientSideComponentId,Title,ClientSideComponentProperties`;

    const [exts, fields] = await Promise.all([getExtensions(userCustomActionUrl), getExtensions(fieldCustomizerUrl)]);

    const siteExtensions: IExtension[] = exts as IExtension[];
    const fieldCustomizers: IExtension[] = getFieldCustomizers(fields as IExtension[]);

    const extensions = siteExtensions.concat(fieldCustomizers);

    console.log('');
    console.log(colors.magenta(`'${scope}' level spfx extensions${listtitle ? ` on '${listtitle}'` : ''} at '${prefs.siteUrl}'`));
    console.log('');
    printToConsole(extensions);

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

async function addExtension(title: string, extensionType: string, scope: string, clientSideComponentId: string, options: any) {

  try {
    ensureAuth();

    const resourcePath = `${scope === ExtensionScope.List ? `web/lists/GetByTitle('${options.listtitle}')` : scope}`;
    const userCustomActionUrl: string = `${prefs.siteUrl}/_api/${resourcePath}/UserCustomActions`;

    const requestBody: any = {
      Title: title,
      Location: `ClientSideExtension.${extensionType}`,
      ClientSideComponentId: clientSideComponentId,
      ClientSideComponentProperties: options.clientProps
    };

    if (options.registrationId) {
      requestBody.RegistrationId = options.registrationId;
    }
    if (options.RegistrationType) {
      requestBody.RegistrationType = RegistrationType[options.registrationType];
    }

    console.log(await postExtension(userCustomActionUrl, JSON.stringify(requestBody)));

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

async function removeExtension(scope: ExtensionScope, id: string, options: any) {
  try {

    ensureAuth();
    const resourcePath = `${scope === ExtensionScope.List ? `web/lists/GetByTitle('${options.listtitle}')` : scope}`;
    const userCustomActionUrl: string = `${prefs.siteUrl}/_api/${resourcePath}/UserCustomActions('${id}')`;

    console.log(await postExtension(userCustomActionUrl, undefined, 'DELETE'));

  } catch (error) {
    console.log(colors.red(error.message));
  }
}

function printToConsole(extensions: IExtension[]) {
  const t = new Table();
  extensions.forEach((extension: IExtension) => {
    t.cell(colors.yellow('Id'), colors.green(extension.Id));
    t.cell(colors.yellow('Title'), colors.green(extension.Title));
    t.cell(colors.yellow('ClientSideComponentId'), colors.green(extension.ClientSideComponentId));
    t.cell(colors.yellow('Location'), colors.green(extension.Location));
    t.cell(colors.yellow('ClientSideComponentProperties'), colors.green(extension.ClientSideComponentProperties));
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
  console.log(response);
  return response ? JSON.parse(response) : '';
}

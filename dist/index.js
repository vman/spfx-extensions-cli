#!/usr/bin/env node
"use strict";
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var program = require("commander");
var node_sp_auth_1 = require("node-sp-auth");
var request = require("request-promise");
var enums_1 = require("./enums");
var Preferences = require('preferences');
var colors = require('colors/safe');
var pjson = require('../package.json');
var Table = require('easy-table');
var prefs = new Preferences('vman.spfx.extensions.cli', {
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
    .on('--help', function () {
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
    .on('--help', function () {
    console.log('');
    console.log('<scope> Scope from which to remove the extension (site | web | list )');
    console.log('<id> of the extension');
    console.log('');
});
program.parse(process.argv);
if (program.connect) {
    prefs.siteUrl = program.connect;
    node_sp_auth_1.getAuth(prefs.siteUrl, {
        ondemand: true,
        electron: require('electron'),
        force: false,
        persist: true
    }).then(function (authResponse) {
        prefs.authHeaders = authResponse.headers;
        prefs.authHeaders = __assign({}, prefs.authHeaders, { Accept: 'application/json;odata=nometadata' });
    });
}
if (program.web) {
    displayExtensions(enums_1.ExtensionScope.Web);
}
if (program.site) {
    displayExtensions(enums_1.ExtensionScope.Site);
}
if (program.list) {
    displayExtensions(enums_1.ExtensionScope.List, program.list);
}
function displayExtensions(scope, listtitle) {
    return __awaiter(this, void 0, void 0, function () {
        var resourcePath, userCustomActionUrl, fieldsPath, fieldCustomizerUrl, _a, exts, fields, siteExtensions, fieldCustomizers, extensions, error_1;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 2, , 3]);
                    ensureAuth();
                    resourcePath = (scope === enums_1.ExtensionScope.List) ? "web/lists/GetByTitle('" + listtitle + "')" : scope;
                    userCustomActionUrl = prefs.siteUrl + "/_api/" + resourcePath + "/UserCustomActions?$filter=startswith(Location, 'ClientSideExtension')&$select=Id,ClientSideComponentId,Title,Location,ClientSideComponentProperties";
                    fieldsPath = (scope === enums_1.ExtensionScope.Web || scope === enums_1.ExtensionScope.List) ? 'fields' : 'rootWeb/availablefields';
                    fieldCustomizerUrl = prefs.siteUrl + "/_api/" + resourcePath + "/" + fieldsPath + "?$select=Id,ClientSideComponentId,Title,ClientSideComponentProperties";
                    return [4 /*yield*/, Promise.all([getExtensions(userCustomActionUrl), getExtensions(fieldCustomizerUrl)])];
                case 1:
                    _a = _b.sent(), exts = _a[0], fields = _a[1];
                    siteExtensions = exts;
                    fieldCustomizers = getFieldCustomizers(fields);
                    extensions = siteExtensions.concat(fieldCustomizers);
                    console.log('');
                    console.log(colors.magenta("'" + scope + "' level spfx extensions" + (listtitle ? " on '" + listtitle + "'" : '') + " at '" + prefs.siteUrl + "'"));
                    console.log('');
                    printToConsole(extensions);
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _b.sent();
                    console.log(colors.red(error_1.message));
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
function addExtension(title, extensionType, scope, clientSideComponentId, options) {
    return __awaiter(this, void 0, void 0, function () {
        var resourcePath, userCustomActionUrl, requestBody, _a, _b, error_2;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    _c.trys.push([0, 2, , 3]);
                    ensureAuth();
                    resourcePath = "" + (scope === enums_1.ExtensionScope.List ? "web/lists/GetByTitle('" + options.listtitle + "')" : scope);
                    userCustomActionUrl = prefs.siteUrl + "/_api/" + resourcePath + "/UserCustomActions";
                    requestBody = {
                        Title: title,
                        Location: "ClientSideExtension." + extensionType,
                        ClientSideComponentId: clientSideComponentId,
                        ClientSideComponentProperties: options.clientProps
                    };
                    if (options.registrationId) {
                        requestBody.RegistrationId = options.registrationId;
                    }
                    if (options.RegistrationType) {
                        requestBody.RegistrationType = enums_1.RegistrationType[options.registrationType];
                    }
                    _b = (_a = console).log;
                    return [4 /*yield*/, postExtension(userCustomActionUrl, JSON.stringify(requestBody))];
                case 1:
                    _b.apply(_a, [_c.sent()]);
                    return [3 /*break*/, 3];
                case 2:
                    error_2 = _c.sent();
                    console.log(colors.red(error_2.message));
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
function removeExtension(scope, id, options) {
    return __awaiter(this, void 0, void 0, function () {
        var resourcePath, userCustomActionUrl, _a, _b, error_3;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    _c.trys.push([0, 2, , 3]);
                    ensureAuth();
                    resourcePath = "" + (scope === enums_1.ExtensionScope.List ? "web/lists/GetByTitle('" + options.listtitle + "')" : scope);
                    userCustomActionUrl = prefs.siteUrl + "/_api/" + resourcePath + "/UserCustomActions('" + id + "')";
                    _b = (_a = console).log;
                    return [4 /*yield*/, postExtension(userCustomActionUrl, undefined, 'DELETE')];
                case 1:
                    _b.apply(_a, [_c.sent()]);
                    return [3 /*break*/, 3];
                case 2:
                    error_3 = _c.sent();
                    console.log(colors.red(error_3.message));
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
function printToConsole(extensions) {
    var t = new Table();
    extensions.forEach(function (extension) {
        t.cell(colors.yellow('Id'), colors.green(extension.Id));
        t.cell(colors.yellow('Title'), colors.green(extension.Title));
        t.cell(colors.yellow('ClientSideComponentId'), colors.green(extension.ClientSideComponentId));
        t.cell(colors.yellow('Location'), colors.green(extension.Location));
        t.cell(colors.yellow('ClientSideComponentProperties'), colors.green(extension.ClientSideComponentProperties));
        t.newRow();
    });
    console.log(t.toString());
}
function getFieldCustomizers(fields) {
    return fields
        .filter(function (field) { return field.ClientSideComponentId !== '00000000-0000-0000-0000-000000000000'; })
        .map(function (fieldCustomizer) {
        fieldCustomizer.Location = 'FieldCustomizer';
        return fieldCustomizer;
    });
}
function ensureAuth() {
    if (!prefs.siteUrl) {
        throw new Error('Please use spfx-ext --connect <siteurl> to auth with SPO. Type spfx-ext --help for help.');
    }
}
function getExtensions(restUrl) {
    return __awaiter(this, void 0, void 0, function () {
        var response;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, request.get({
                        url: restUrl,
                        headers: prefs.authHeaders
                    })];
                case 1:
                    response = _a.sent();
                    return [2 /*return*/, JSON.parse(response).value];
            }
        });
    });
}
function postExtension(restUrl, requestBody, requestMethod) {
    if (requestBody === void 0) { requestBody = ''; }
    if (requestMethod === void 0) { requestMethod = 'POST'; }
    return __awaiter(this, void 0, void 0, function () {
        var reqDigestResponse, requestDigest, postHeaders, response;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, request.post({
                        url: prefs.siteUrl + "/_api/contextinfo",
                        headers: prefs.authHeaders
                    })];
                case 1:
                    reqDigestResponse = _a.sent();
                    requestDigest = JSON.parse(reqDigestResponse).FormDigestValue;
                    postHeaders = __assign({}, prefs.authHeaders, { 'X-RequestDigest': requestDigest, 'content-type': 'application/json;odata=nometadata' });
                    if (requestMethod === 'DELETE') {
                        postHeaders = __assign({}, postHeaders, { 'X-HTTP-Method': 'DELETE' });
                    }
                    return [4 /*yield*/, request.post({
                            url: restUrl,
                            body: requestBody,
                            method: requestMethod,
                            headers: postHeaders
                        })];
                case 2:
                    response = _a.sent();
                    console.log(response);
                    return [2 /*return*/, response ? JSON.parse(response) : ''];
            }
        });
    });
}

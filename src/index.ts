#!/usr/bin/env node
import * as program from 'Commander';
import { getAuth } from 'node-sp-auth';
import * as pnp from 'sp-pnp-js';
import NodeFetchClient from 'node-pnp-js';

program
  .version('0.1.0')
  .option('-c, --connect', 'Connect to SharePoint Online')
  .option('-s, --siteurl [siteurl]', 'SharePoint Online Site Url [null]', null)
  .parse(process.argv);

if (program.connect) {

  if (!program.siteurl) {
    console.error('Please enter siteurl with --siteurl or -s');
    process.exit();
  }

  const siteUrl = program.siteurl;
  getAuth(siteUrl, {
    ondemand: true,
    electron: require('electron'),
    force: false,
    persist: true
  }).then(() => {
    pnp.setup({
      fetchClientFactory: () => {
        return new NodeFetchClient({
          ondemand: true
        });
      },
      baseUrl: siteUrl
    });
  }).then(() => {
    // we need to use the Web constructor to ensure we have the absolute url
    const web = new pnp.Web(siteUrl);
    //pnp.sp.profiles.myProperties.get("");
    web.select('Title').get().then((w) => {
      console.log(`Web's title: ${w.Title}`);
    });
  });
}
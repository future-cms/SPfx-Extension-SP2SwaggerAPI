import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { sp, ClientSidePage, CanvasSection, CanvasColumn, ClientSideText, ClientSideWebpart} from "@pnp/sp";
import {parse, stringify} from 'flatted/esm';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SendPageContentToWebBackEndApiApplicationCustomizerStrings';
import { JSONParser } from '@pnp/odata';

const LOG_SOURCE: string = 'SendPageContentToWebBackEndApiApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISendPageContentToWebBackEndApiApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  _pushPageChangesToApi: ()=>Promise<object>;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SendPageContentToWebBackEndApiApplicationCustomizer
  extends BaseApplicationCustomizer<ISendPageContentToWebBackEndApiApplicationCustomizerProperties> {
    private getPageContent(page:ClientSidePage): Promise<void>{
    
      return Promise.resolve();
    }
    private _pushPageChangesToApi(page:ClientSidePage): Promise<object>{
      const jsonToApi = {rows:[]};
      let rowIterator = 1;
      for (const section of page.sections as Array<CanvasSection>) {
        let columnIterator = 1;
        jsonToApi.rows.push({id:rowIterator,columns:[]});
        for (const column of section.columns as Array<CanvasColumn>) {
          let controlIterator = 1;
          jsonToApi.rows[rowIterator].columns.push({id:columnIterator,controls:[]});
          for (const control of column.controls as Array<ClientSideText>) {
            jsonToApi.rows[rowIterator][columnIterator].controls.push(
              {
                id:controlIterator,
                text:control.text,
                controlData: control.jsonData
              }
              );    
              controlIterator++; 
          }
          columnIterator++;
        }
        rowIterator++;
      }
      return Promise.resolve(jsonToApi);
    }

  @override
  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }
    sp.setup({
      spfxContext: this.context
    });
   // Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);
    // override browser's fetch to check if page published
        // tslint:disable-next-line:no-function-expression
        (function(ns, fetch){
          if(typeof fetch !== 'function') return;
          
          ns.fetch = function() {
            // Only trigger for spesific site and 'sitepages' and on 'publish'
        // tslint:disable-next-line:use-named-parameter 
        if(undefined===arguments || undefined===arguments[0] || undefined===arguments[0].url){
          return fetch.apply(this, arguments);
        }
        // tslint:disable-next-line:use-named-parameter 
            if(arguments[0].url.indexOf('{mysharepoint}.sharepoint.com') > -1 && arguments[0].url.indexOf('sitepages') > -1 && arguments[0].url.indexOf('publish') > -1){
               ClientSidePage.fromFile(sp.web.getFileByServerRelativeUrl(window.location.pathname))
               .then(page=>{
                 let metaFields = [
                   'name',
                   'metaTitle',
                   'metadescription',
                   'metaImage',
                   'metaContentType',
                   'metaTwitter',
                   'metaFacebookAppId',
                   'metaNoCrawl',
                   'created',
                   'modified',
                   'metaCategory',
                   'metaTags'
                 ];
                page.listItemAllFields.select(metaFields.join(',')).get().then(
                  meta=>{
                    let sections = page.sections;
                    for (const section of sections as Array<CanvasSection>) {
                      delete section.page;
                      for (const column of section.columns as Array<CanvasColumn>) {
                        delete column.section;
                        for (const control of column.controls as Array<ClientSideText>&Array<ClientSideWebpart>) {
                          delete control.column;
                        }
                      }
                    }
                    let jsonToApi:string = JSON.stringify(page);
                    let bodyText = {
                      name: "Widget Adapter",
                      releaseDate: Date(),
                      id: window.location.pathname,
                      manufacturer:{
                        phone: "555-000-0000",
                        name: document.title,
                        homePage: window.location.origin + window.location.pathname,
                        content: page,
                        meta: meta
                      }
                    };
                    fetch('https://{MY_SWAGGER_API_SERVER}/ConnectorAPI/1.0.0/inventory/', {
                      method: 'post',
                      headers: {
                        "Content-Type": "application/json; charset=utf-8",
                        "API_KEY_AUTH": "SOME_RANDOM_STRING_SHARED_WITH_SPFX_PLUGIN_OR_MS_FLOW",
                        "Access-Control-Allow-Origin": "*",
                        "Referer" :  window.location.origin + window.location.pathname
                        // "Content-Type": "application/x-www-form-urlencoded",
                    },
                      body: JSON.stringify(bodyText)
                    })
                    .then(response=>response.json())
                    .then(data => "Api Response" + console.log(data)).catch(err=>console.log(err));
                  }
                );


                
                
/*
                const jsonToApi = {rows:[]};
                let rowIterator = 0;
                let sections = page.sections;
                for (const section of sections as Array<CanvasSection>) {
                  let columnIterator = 0;
                  jsonToApi.rows.push({id:rowIterator,columns:[]});
                  for (const column of section.columns as Array<CanvasColumn>) {
                    let controlIterator = 0;
                    jsonToApi.rows[rowIterator].columns.push({id:columnIterator,controls:[]});
                    for (const control of column.controls as Array<ClientSideText>&Array<ClientSideWebpart>) {
                      jsonToApi.rows[rowIterator].columns[columnIterator].controls.push(
                        {
                          id:controlIterator,
                          // tslint:disable-next-line:member-access
                          text:control.text?control.text:control.htmlProperties,
                          controlData: control.jsonData
                        }
                        );    
                        controlIterator++; 
                    }
                    columnIterator++;
                  }
                  rowIterator++;
                }
                */


               });
              // tslint:disable-next-line:use-named-parameter 
              /*
              this.getPageContent(arguments[0].url).then(content=>{
                  this.pushPageChangesToApi(content).then(res=>{
                    Dialog.alert(`Hello from ${strings.Title}:\n\n${res}`);
                  });
              });
              */
            }        
            return fetch.apply(this, arguments);
          };
          
        }(window, window.fetch));
    return Promise.resolve();
  }

}

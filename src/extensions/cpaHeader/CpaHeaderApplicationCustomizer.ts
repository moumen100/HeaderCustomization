import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CpaHeaderApplicationCustomizerStrings';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
const LOG_SOURCE: string = 'CpaHeaderApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICpaHeaderApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  webUrl:string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}
export interface ISPLISTITEM {
  value:any[];
}
/** A Custom Action which can be run during execution of a Client Side Application */
export default class CpaHeaderApplicationCustomizer
  extends BaseApplicationCustomizer<ICpaHeaderApplicationCustomizerProperties> {
    private topPlaceholder: PlaceholderContent | undefined;
    private HostUrl : string = "http://orgaplums/";
    private ListName: string = "Footer";
    private LinksList=[];
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // let cssUrl: string = this.properties.cssurl
    //     const cssUrl: string = "http://10.1.0.10:9090/StyleInject/bnaStyle.css"

    // if (cssUrl) {
    //     // inject the style sheet
    //     const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
    //     let customStyle: HTMLLinkElement = document.createElement("link");
    //     customStyle.href = cssUrl;
    //     customStyle.rel = "stylesheet";
    //     customStyle.type = "text/css";
    //     head.insertAdjacentElement("beforeEnd", customStyle);
    // }
    
    this.context.placeholderProvider.changedEvent.add(this,this.renderPlaceholders);
        console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
    

    return Promise.resolve();
  }
  private async _getListData(Title:string): Promise<ISPLists> {
    return this.context.spHttpClient.get(`${this.HostUrl}/_api/Web/Lists?&$select=Title, Id&$filter=Title eq '${Title}'&$Top=1`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private async getListItems(listId:string) : Promise<ISPLISTITEM> {

    return this.context.spHttpClient.get( `${this.HostUrl}/_api/Web/Lists/getbytitle('${this.ListName}')/Items?$select=ID,Title,url&$Top=5`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
   private async renderPlaceholders():Promise<void>{

    console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
     //get topPlaceholder
     if(!this.topPlaceholder)
     {
       // tring to get topPlaceholder
       this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top,{onDispose:this.onPlaceholderDispose})
      }
     //check if topPlaceholder is available to you
     if(this.topPlaceholder)
     {

        if(this.topPlaceholder.domElement)
        {
          let listname = await this._getListData(this.ListName) 
           this.LinksList = await (await this.getListItems(listname.value[0].Id)).value
        //   await this._getListData(this.ListName).then(lists =>{
        //     let listId = lists.value[0].Id
        //     this.getListItems(listId).then(listItems=>{
        //       console.log(listItems.value)
        //       this.LinksList = listItems.value
        //     })
        // })
          let style:String = `
          <style>
.uls {
  list-style-type: none;
  margin: 0;
  padding: 0;
  overflow: hidden;
  background-color: #3c3c3b;
}

.lis {
  float: left;
}

.lis:last-child {
  border-right: none;
}

.lis a {
  display: block;
  color: white;
  text-align: center;
  padding: 14px 16px;
  text-decoration: none;
}

.lis a:hover:not(.active) {
  background-color: #459F36;
}

.active {
  background-color: #459F36;
}
</style>`
let Menu:String = `
<nav class="navbar">
<ul class="uls">`
if (this.LinksList.length>0) {
  this.LinksList.forEach(Element=>{
    Menu+=
    `
    <li class="lis"><a href="${Element.url}">${Element.Title} </a></li>
    `
  }) 
   
}
else {
  Menu+=
    `
  <li class="lis"><a href="#">No link data available</a></li>
  `
}



Menu+=
`
</ul>
</nav>`
this.topPlaceholder.domElement.innerHTML = `${style}  ${Menu}`
console.log(this.LinksList)

console.log(this.topPlaceholder.domElement)
        }
     }
     else{
       //could not get top placeholder
       console.log("Could not get top placeholder");
       return;
     }
}
private onPlaceholderDispose(){
  console.log("Placeholder disposed");
}
}

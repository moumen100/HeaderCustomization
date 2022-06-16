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
    this._getListData(this.ListName).then(lists =>{
        let listId = lists.value[0].Id
        this.getListItems(listId).then(listItems=>{
          console.log(listItems.value)
        })
    })
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

    return this.context.spHttpClient.get( `${this.HostUrl}/_api/Web/Lists/getbytitle('${this.ListName}')/Items?$select=ID,Title&Top='5'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });

  }
  private renderPlaceholders():void{

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
          this.topPlaceholder.domElement.innerHTML=`
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
</style>
<nav class="navbar">
<ul class="uls">
  <li class="lis"><a href="http://sharepoint2019:100/sites/Site_BNA">Accueil </a></li>
  <li class="lis"><a href="http://sharepoint2019:100/sites/Site_BNA/SitePages/Notre-Banque(2).aspx">Notre Banque </a></li>
  <li class="lis"><a href="http://sharepoint2019:100/sites/DRH/SitePages/Home.aspx">Ressources Humaines</a></li>
  <li class="lis"><a href="http://sharepoint2019:100/sites/DirCom">Direction Commerciale</a></li>
  <li class="lis"><a href="http://sharepoint2019:100/sites/Site_BNA/SitePages/Organisation.aspx">Organisation</a></li>

</ul>
</nav>`
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

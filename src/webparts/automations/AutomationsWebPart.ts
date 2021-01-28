import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AutomationsWebPart.module.scss';
import * as strings from 'AutomationsWebPartStrings';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "../../ExternalRef/CSS/Style.css";
export interface IAutomationsWebPartProps {
  description: string;
}
const SiteUrl = "https://simpleiddemoaccount.sharepoint.com";
export default class AutomationsWebPart extends BaseClientSideWebPart<IAutomationsWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      sp.setup({
        sp: {
          baseUrl: "https://simpleiddemoaccount.sharepoint.com/sites/HelpDeskSite",
        },
      });
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="content">
    <h3 class="heading">HALSTEAD AUTOMATION HUB</h3>
    <div id="tile-append">
    </div>
    </div>`;
    this.getTiles();
  }
 
  getTiles()
  { 
     
    sp.web.getFolderByServerRelativeUrl("PowerAppsPanel").files.select("ListItemAllFields/AppDescription,ListItemAllFields/AppUrl,ListItemAllFields/Title,ListItemAllFields/Active,ListItemAllFields/Name,ListItemAllFields/FileRef").expand("ListItemAllFields","Name").filter("ListItemAllFields/Active eq "+true+"").get().then((items1)=>{
      console.log(items1);
      var allItems=items1
      var html="<ul>";
      for(let i=0;i<allItems.length;i++)
      {
        // html+="<li><a href="+allItems[i]["ListItemAllFields"].AppUrl+">"+allItems[i]["ListItemAllFields"].Title+"</a></li>"
        html+=`<li><a href="${allItems[i]["ListItemAllFields"].AppUrl}"><img src="${SiteUrl}${allItems[i]["ListItemAllFields"].FileRef}" alt=""></a></li>`
      }
      // document.getElementById("tile-append").insertAdjacentHTML("afterend", 
      // html); 
      document.getElementById("tile-append").innerHTML = html;
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

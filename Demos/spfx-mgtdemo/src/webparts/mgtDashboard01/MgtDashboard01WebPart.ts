import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MgtDashboard01WebPart.module.scss';
import * as strings from 'MgtDashboard01WebPartStrings';
import { Providers, SharePointProvider, MgtPeoplePicker } from '@microsoft/mgt';

export interface IMgtDashboard01WebPartProps {
  description: string;
}


export default class MgtDashboard01WebPart extends BaseClientSideWebPart<IMgtDashboard01WebPartProps> {

  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
  }
  

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.mgtDashboard01}">
        <div class="${ styles.container}">
          <div class="${styles.row}">
          <h2 class="${styles.title}">SPS Chennai - Microsoft Graph Toolkit</h2>
            <div class="${styles.column}">
            <mgt-get resource="/groups" id="mgtGroup" scope="group.read.all" version="v1.0">
            <template id="td" data-type="default">
            <select id="selGroups" class="${styles["ms-Select"]}"   >
            <option data-for="group of value" value={{group.id}}>{{group.displayName}}</option>
            </select>              
            </template>
          </mgt-get>
          <button  id="btnGetMembers" class="${styles["ms-Button"]}" type="button">
            <span class="ms-Button-label">Get Members</span> 
        </button>
        <div id="showgrpmembers"></div>
              

            </div>
            <div class="${styles.column}">
            <mgt-people-picker id="mgtpicker" ></mgt-people-picker>
            <button  id="btnAddName" class="${styles["ms-Button"]}" type="button">
            <span class="ms-Button-label"> Add Member</span> 
          </button>
            </div>
          </div>
        </div>
      </div>`;

    this.attachEventBinders();
  }


  private attachEventBinders(): void {
   this.domElement.querySelector('#btnGetMembers').addEventListener('click', () => {
     this.getmembers();
   });
   this.domElement.querySelector('#btnAddName').addEventListener('click', () => {
     this.groupAddName();
  });
  }

  private getmembers(): void {    
   var e: HTMLSelectElement = this.domElement.querySelector('#selGroups');
   var strUser = e.options[e.selectedIndex].value;
 this.generateMembers(strUser);
  }

  private generateMembers(groupid): void {
   var html = "<mgt-people id=\"grpPeople\" group-id=" + groupid + " show-max='10' ></mgt-people>"
   document.getElementById("showgrpmembers").innerHTML = html;
  }

// ADD USER TO THE GROUP
  private groupAddName(): void {
   var mt: MgtPeoplePicker = document.querySelector('mgt-people-picker');
   const p = Providers.globalProvider;
   var e: HTMLSelectElement = this.domElement.querySelector('#selGroups');
   var strUser = e.options[e.selectedIndex].value;
   if (p) {
     let graphClient = p.graph.client;
     const directoryObject = {
       "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + mt.selectedPeople[0].id
     };
     let userDetails = graphClient.api('/groups/' + strUser + '/members/$ref').post(directoryObject).then((resp) => {
       console.log(resp);
       this.generateMembers(strUser);
     });
   }
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

import { Version } from '@microsoft/sp-core-library';
import * as $ from "jquery"
import "datatables.net";
// import "datatables-epresponsive";
// import "datatables.net-dt"; 
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxsharepointWebPart.module.scss';
import * as strings from 'SpfxsharepointWebPartStrings';
import pnp from 'sp-pnp-js';

export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Employee_x0020_ID: string;
  Employee_x0020_Salary: string;
  Employee_x0020_Name: string;
  EmpDOB: string;
  EmpJobTitle:string;
  Employee_Mobile_Number:string;
  Emp_Department:string;
  Emp_sex:string;
}
export interface ISpfxsharepointWebPartProps {
  description: string;
}

export default class SpfxsharepointWebPart extends BaseClientSideWebPart<ISpfxsharepointWebPartProps> {

  private _getListData(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle("Employee data").items.get().then((response) => {
     
       return response;
     });
       
    }
  private _renderListAsync(): void {


    this._getListData()
      .then((response) => {
        this._renderList(response);
      });

  }
  private _renderList(items: ISPList[]): void {
    let htm: string = '<div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Student Details</div>';
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;"> <thead>';
    html += `<th>Employee ID</th><th>Employee Salary</th><th>Employee Name</th><th>Employee DOB</th><th>Employee Job Title</th><th>Employee Mob No</th><th>Employee Department</th><th>Employee sex</th></thead>`;
    items.forEach((item: ISPList) => {
      html += `<tbody>  
         <tr>  
        <td>${item.Employee_x0020_ID}</td>  
        <td>${item.Employee_x0020_Salary}</td>  
        <td>${item.Employee_x0020_Name}</td>  
        <td>${item.EmpDOB}</td>
        <td>${item.EmpJobTitle}</td>
        <td>${item.Employee_Mobile_Number}</td>
        <td>${item.Emp_Department}</td>
        <td>${item.Emp_sex}</td>  
        </tr> </tbody>
        `;
    });
    html += `</table>`;
    const heading: Element = this.domElement.querySelector('#heading');
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    heading.innerHTML = htm;
    listContainer.innerHTML = html;
    $('.TFtable').DataTable();
  }
 public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxsharepoint}">
    <div class="${ styles.container}">
      <div class="${ styles.row}">
        <div class="${ styles.column}">
          <span class="${ styles.title}">Welcome to SharePoint!</span>
  <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
    <p class="${ styles.description}">${escape(this.properties.description)}</p>
     
          <div id="heading" /> </div>
          <br>
          <div id="spListContainer" />  
            </div>  
          </div>
          </div>
          </div>
          </div>`;
    this._renderListAsync();
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

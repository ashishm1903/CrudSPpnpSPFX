import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudSPpnpJsWebPart.module.scss';
import * as strings from 'CrudSPpnpJsWebPartStrings';
import * as pnp from 'sp-pnp-js';
export interface ICrudSPpnpJsWebPartProps {
  description: string;
}

export default class CrudSPpnpJsWebPart extends BaseClientSideWebPart<ICrudSPpnpJsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div>
    <div >
    <h2>Student Registration</h2>
    <hr/>
    <table border='5' bgcolor='aqua'>
      <tr>
      <td>Please Enter Student ID </td>
      <td><input type='text' id='txtID' />
      <td><input type='submit' id='btnRead' value='Read Details' />
      </td>
      </tr>
      <tr>
      <td>Title</td>
      <td><input type='text' id='txtTitle' />
      </tr>
      <tr>
      <td>Name</td>
      <td><input type='text' id='txtName' />
      </tr>
      <tr>
      <td>Gender</td>
      <td>
      <select id="ddlGender">
        <option value="Male">Male</option>
        <option value="Female">Female</option>
      </select>  
      </td>
      </tr>
      <tr>
      <td>Mobile Number</td>
      <td><input type='text' id='txtMobileNo' />
      </tr>
      <tr>
      <td>Email</td>
      <td><input type='text' id='txtEmail' />
      </tr>
      <tr>
      <td>Address</td>
      <td><textarea rows='5' cols='40' id='txtAddress'> </textarea> </td>
      </tr>
      <tr>
      <td colspan='2' align='center'>
      <input type='submit'  value='Insert Item' id='btnSubmit' />
      <input type='submit'  value='Update' id='btnUpdate' />
      <input type='submit'  value='Delete' id='btnDelete' />
      </td>
    </table>
    </div>
    <div id="divStatus"/>
    <h2>Student List</h2>
    <hr/>
    <div id="spListData" />
    </div>`;
    this._bindEvents();
    this.readAllItems();
  }
  public readAllItems(): void {
    let html: string = '<table border=1 width=100% style="bordercollapse: collapse;">';
    html += `<th>ID</th><th>Title</th><th>Name</th><th>Mobile No</th><th>Gender</th><th>Email</th><th>Address</th>`;
  pnp.sp.web.lists.getByTitle("StudentData").items.get().then((items: any[]) => {
    items.forEach(function (item) {  
      html += `
      <tr>
      <td>${item["ID"]}</td>
      <td>${item["Title"]}</td>
      <td>${item["Name"]}</td>
      <td>${item["Mobile"]}</td>
      <td>${item["Gender"]}</td>
      <td>${item["Email"]}</td>
      <td>${item["Address"]}</td>
      </tr>
      `;
    });  
    html += `</table>`;
    const allitems: Element = this.domElement.querySelector('#spListData');
    allitems.innerHTML = html;
});
  }
  private _bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnRead').addEventListener('click', () => { this.readListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this.deleteListItem(); });
  }
  private deleteListItem(): void {
    const id = document.getElementById("txtID")["value"];
    pnp.sp.web.lists.getByTitle("StudentData").items.getById(id).delete();
    alert("Record Deleted Succesfully..!");
    this.clearData();
  }
  private updateListItem(): void {
    var title = document.getElementById("txtTitle")["value"];
    var name = document.getElementById("txtName")["value"];
    var mobile = document.getElementById("txtMobileNo")["value"];
    var gender = document.getElementById("ddlGender")["value"];
    var email = document.getElementById("txtEmail")["value"];
    var address = document.getElementById("txtAddress")["value"];
    let id: number = document.getElementById("txtID")["value"];

    pnp.sp.web.lists.getByTitle("StudentData").items.getById(id).update({
      Title: title,
      Name: name,
      Gender: gender,
      Mobile: mobile,
      Email: email ,
      Address: address 
    }).then(r => {
      alert("Record Updated Succesfully..!");
    });
    this.clearData();
  }
  private readListItem(): void {
    const id = document.getElementById("txtID")["value"];
    if(!id)
    {
      alert("Please Enter Student ID")
    }
    else{
    pnp.sp.web.lists.getByTitle("StudentData").items.getById(id).get().then((item: any) => {
      document.getElementById("txtTitle")["value"] = item["Title"];
      document.getElementById("txtName")["value"] = item["Name"];
      document.getElementById("txtMobileNo")["value"] = item["Mobile"];
      document.getElementById("ddlGender")["value"] = item["Gender"];      
      document.getElementById("txtEmail")["value"] = item["Email"];
      document.getElementById("txtAddress")["value"] = item["Address"];
    });
  }
  }
  private addListItem() : void {
    var title = document.getElementById("txtTitle")["value"];
    var name = document.getElementById("txtName")["value"];
    var mobile = document.getElementById("txtMobileNo")["value"];
    var gender = document.getElementById("ddlGender")["value"];
    var email = document.getElementById("txtEmail")["value"];
    var address = document.getElementById("txtAddress")["value"];
    const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('StudentData')/items";
    pnp.sp.web.lists.getByTitle("StudentData").items.add({
      Title: title,
      Name: name,
      Mobile: mobile,
      Gender: gender,
      Email: email,
      Address: address
    }).then(r => {
      alert("Record Inserted Succesfully..!");
    });   
    this.clearData();
  }
  private clearData(){
    document.getElementById("txtID")["value"] = "";
    document.getElementById("txtTitle")["value"] = "";
      document.getElementById("txtName")["value"] = "";
      document.getElementById("txtMobileNo")["value"] = "";
      document.getElementById("ddlGender")["value"] ="Male";      
      document.getElementById("txtEmail")["value"] ="";
      document.getElementById("txtAddress")["value"] = "";
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

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as $ from 'jquery';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";

var technologyChoices;
export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart <IHelloWorldWebPartProps> {

  public async render(): Promise<void> {
    var currentobj = this;
    technologyChoices = await getChoiceFields(this.context.pageContext.web.absoluteUrl);
    let web = Web(this.context.pageContext.web.absoluteUrl);
    const items: any[] = await web.lists.getByTitle("ProjectDetails").items.getAll();
    console.log(items);
    this.domElement.innerHTML = `
        ${await this.getHTML()}
        `;
    
    var create = document.getElementById('Create');
    create.addEventListener('click', async function () {
      // var e = document.getElementById("technology");
      //var strUser = e.options[e.selectedIndex].text;  
      await web.lists.getByTitle("ProjectDetails").items.add({

        Title: $("#projectname").val(),
        CompletionStatus: $("#completed:checked").val() == "completed" ? true : false,
        Technology: $("#technology :selected").text()

      }).then(i => {
        console.log(i);
      });
      alert("Created Successfully");
      await currentobj.render();
      $("#projectname").val("");
      $("#completed").prop('checked', false);
      $("#technology").val("");

    });

    var update = document.getElementById('Update');
    update.addEventListener('click', async function () {
      // var e = document.getElementById("technology");
      //var strUser = e.options[e.selectedIndex].text;
      let id:number=Number($('input[name="itemID"]:checked').val());
      await web.lists.getByTitle("ProjectDetails").items.getById(id).update({
        Title: $("#projectname").val(),
        CompletionStatus: $("#completed:checked").val() == "completed" ? true : false,
        Technology: $("#technology :selected").text()

      }).then(i => {
        console.log(i);
      });
      alert("Updated Successfully");
      await currentobj.render();
      $("#projectname").val("");
      $("#completed").prop('checked', false);
      $("#technology").val("");
    });


    var deletedata = document.getElementById('Delete');
    deletedata.addEventListener('click', async function () {
      // var e = document.getElementById("technology");
      //var strUser = e.options[e.selectedIndex].text;  
      let id:number=Number($('input[name="itemID"]:checked').val());
      await web.lists.getByTitle("ProjectDetails").items.getById(id).delete()
        .then(i => {
          console.log(i);
        });
      alert("Deleted Successfully");
      await currentobj.render();
      $("#projectname").val("");
      $("#completed").prop('checked', false);
      $("#technology").val("");
    });

    var UpdateDeleteItem = document.getElementById('UpdateDeleteItem');
    UpdateDeleteItem.addEventListener('click', async function () {
      let id:number=Number($('input[name="itemID"]:checked').val());
      const item: any[] = await web.lists.getByTitle("ProjectDetails").items.getById(id).get();
      console.log(item);

      $("#projectname").val(item["Title"]);
      $("#completed").prop('checked', item["CompletionStatus"]);
      $("#technology").val(item["Technology"]);
    });

    $('input[name*="itemID"]').on('change', async function () {
      let id:number=Number($('input[name="itemID"]:checked').val());
      const item: any[] = await web.lists.getByTitle("ProjectDetails").items.getById(id).get();
      console.log(item);

      $("#projectname").val(item["Title"]);
      $("#completed").prop('checked', item["CompletionStatus"]);
      $("#technology").val(item["Technology"]);
    });

  }

  public async getHTML() {
    let web = Web(this.context.pageContext.web.absoluteUrl);
    const items: any[] = await web.lists.getByTitle("ProjectDetails").items.getAll();
    console.log(items);
    return this.domElement.innerHTML = `<div>
        <h1>CRUD Operations with No Javascript in SPFx</h1>
        <table border="1" class="${styles.table}">
      <thead>
        <tr>
          <th></th>
          <th>Project Name</th>
          <th>Completion Status</th>
          <th>Technology</th>
        </tr>
      </thead>
      <tbody>
        ${items && items.map((item, i) => {
      return [
        "<tr id=UpdateDeleteItem class=UpdateDeleteItem>" +
        "<td><input type=radio id=" + item.ID + " name=itemID value=" + item.ID + "></td>" +
        "<td>" + item.Title + "</td>" +
        "<td>" + checkStatus(item.CompletionStatus) + "</td>" +
        "<td>" + item.Technology + "</td>" +
        "</tr>"
      ];
    })}
      </tbody>

    </table>
    <form>

        <button class="${styles.button}" type="button" id="Create">Create</button>
        <button class="${styles.button}" type="button" id="Update">Update</button>
        <button class="${styles.button}" type="button" id="Delete">Delete</button>
        <br><br>

        <label class="${styles.label}" for="projectname">Project Name:</label>

        <input class="${styles.borderset}" type="text" id="projectname" name="projectname"><br><br>

        <label class="${styles.label}" for="completionstatus">Completion Status:</label>
        <input class="${styles.borderset}" type="radio" id="completed" name="completed" value="completed">
        <label for="completed">Completed</label><br><br>
        <label class="${styles.label}" for="technology">Technology:</label>
<select class="${styles.borderset}" name="technology" id="technology" form="technologyform">
${technologyChoices.map((item) => {
      return [
        `<option value='${item.key}'>${item.value}</option>`
      ];
    })}
  
  </select><br><br></form>
  </div>`;

  }

  //@ts-ignore
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

export const getChoiceFields = async (webURL) => {
  let resultarr = [];
  await fetch(webURL + "/_api/web/lists/GetByTitle('ProjectDetails')/fields?$filter=EntityPropertyName eq 'Technology'", {
    method: 'GET',
    mode: 'cors',
    credentials: 'same-origin',
    headers: new Headers({
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-cache',
      'pragma': 'no-cache',
    }),
  }).then(async (response) => await response.json())
    .then(async (data) => {
      for (var i = 0; i < data.value[0].Choices.length; i++) {
        //for(var j=0;j<=i;j++)
        //{
        await resultarr.push({
          key: data.value[0].Choices[i],
          value: data.value[0].Choices[i]
          //key:data.value[i].Choices[j],
          //text:data.value[i].Choices[j]
        });
        //}
      }
    });
  return await resultarr;
};
export const checkStatus = (value): string => {

  if (value) {
    return 'Completed';
  }
  else {
    return 'In Progress';
  }
};


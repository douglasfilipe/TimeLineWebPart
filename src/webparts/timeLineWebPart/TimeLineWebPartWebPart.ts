import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'TimeLineWebPartWebPartStrings';
import * as $ from 'jquery';
require('../../stylelibrary/css/timeline.css');
export interface ITimeLineWebPartWebPartProps {
  description: string;
  listName: string;
  endDate: string;
  listInt: string;
}

export default class TimeLineWebPartWebPart extends BaseClientSideWebPart<ITimeLineWebPartWebPartProps> {

 public getdeliver() {

    $.ajax({url: `${this.context.pageContext.web.absoluteUrl}` +
   `/_api/web/lists/getByTitle('` + this.properties.listName + `')/items?$select=Title,` + this.properties.endDate +`,Description,ID&$orderby=`+ this.properties.endDate +`&$top=1000`,
      method: 'GET',
      async: false,
      headers: {
        Accept: 'application/json; odata=verbose'
      },
      success: (data) => {
        let html = '';

        if (data.d.results.length > 0) {
          var lines = 0;
          html += `<div class="text-center"><h2>` + this.properties.description + `</h2></div>`;
          $.each(data.d.results, (i, result) => {
          const linenumber = lines ++;
          const alternate = (linenumber % 2 === 0 ? "left":"right");
          const currentPageUrl = this.context.pageContext.site.serverRequestPath;
          const editform = (encodeURI(this.context.pageContext.web.absoluteUrl + `/Lists/` + this.properties.listInt +`/EditForm.aspx?ID=` + result.ID + `&Source=` + window.location));
            html += `<div>` +
                    `<div class="timeline">` +
                    `<div class="container ` + alternate + `">` +
                    `<div class="content"><a href=` + editform + `><h2>${result.Title}</h2></a><p>${result.Description}</p></div>` +
                    `<div class="date">${new Date(result.CommitmentFinish).toLocaleDateString('pt-br')}</div>` +
                    `</div>`;

          });
          $('#divDeliver').html(html);
        }
        else
        {
          $('#divDeliver').html(html);
        }
      },
      error: (errorCode, errorMessage) => {
        console.log('Erro ao recuperar os itens. \nError: ' + errorCode + '\nStackTrace: ' + errorMessage);
      }
    });
  }
  public render(): void {
    // load template layout
    this.domElement.innerHTML = require('./timeline.html');
    this.getdeliver();
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
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneTextField('listInt', {
                  label: strings.ListIntFieldLabel
                }),
                PropertyPaneTextField('endDate', {
                  label: strings.endDateFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

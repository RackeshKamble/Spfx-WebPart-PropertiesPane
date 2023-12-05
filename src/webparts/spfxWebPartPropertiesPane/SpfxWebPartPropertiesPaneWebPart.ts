import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  //Define Controls here
  PropertyPaneDropdown,
  PropertyPaneDropdownOptionType
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxWebPartPropertiesPaneWebPart.module.scss';
import * as strings from 'SpfxWebPartPropertiesPaneWebPartStrings';

export interface ISpfxWebPartPropertiesPaneWebPartProps {
  description: string;
  firstName : string;
  lastName : string;
  gender : string;
}

export default class SpfxWebPartPropertiesPaneWebPart extends BaseClientSideWebPart<ISpfxWebPartPropertiesPaneWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxWebPartPropertiesPane }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">${escape(this.properties.firstName)} ${escape(this.properties.lastName)} </p>
              <p class="${ styles.description }">${escape(this.properties.gender)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Field events --> Non-Reactive - Updates values in webpart after event like Apply button click
  protected get disableReactivePropertyChanges() : boolean{
    return true;
  }

  //Validate web part properties
  private validateFirstName(value: string): string {
    if (value === null ||  value.trim().length === 0) {
        return 'Please provide first name';
        }

    if (value.length > 30) {
    return 'First Name should not be greate than 40 characters';
    }

        return '';
    }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            //description: strings.PropertyPaneDescription
            description: "This is a custom Header"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            // New Properties Group Added
            {
              groupName: "Custom Props Group",
              groupFields: [
                PropertyPaneTextField('firstName', {
                  label: 'First Name',

                  //call validate firstname method here
                  onGetErrorMessage : this.validateFirstName.bind(this)

                }),
                PropertyPaneTextField('lastName', {
                  label: 'Last Name'
                }),
                //Drop Down Added
                PropertyPaneDropdown('Gender', {
                  label: 'gender',
                  options:
                  [
                    {key : 'Male' , text :'Male'},
                    {key : 'Female' , text :'Female'},
                    {key : 'Others' , text :'Others'}
                  ]
                })
              ]
            }
          ]
        },
        //shows next option with next page for properties
        {
          header: {
            //description: strings.PropertyPaneDescription
            description: "This is a second page Header"
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

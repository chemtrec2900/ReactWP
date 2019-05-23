import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import * as strings from 'ReactWebpartDemoWebPartStrings';
import ReactWebpartDemo from './components/ReactWebpartDemo';
import { IReactWebpartDemoProps } from './components/IReactWebpartDemoProps';

export interface IReactWebpartDemoWebPartProps {
  description: string;
  collectionData: any[];
}

export default class ReactWebpartDemoWebPart extends BaseClientSideWebPart<IReactWebpartDemoWebPartProps> {
  

  public render(): void {
    const element: React.ReactElement<IReactWebpartDemoProps > = React.createElement(
      ReactWebpartDemo,
      {
        description: this.properties.description,
        spHttpClient: this.context.spHttpClient,
        currentSiteUrl: this.context.pageContext.web.absoluteUrl
      }
    );
      console.log(this.properties.collectionData);
    ReactDom.render(element, this.domElement);
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
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: "Title",
                      title: "Firstname",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "Lastname",
                      title: "Lastname",
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: "Age",
                      title: "Age",
                      type: CustomCollectionFieldType.number,
                      required: true
                    },
                    {
                      id: "City",
                      title: "Favorite city",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "antwerp",
                          text: "Antwerp"
                        },
                        {
                          key: "helsinki",
                          text: "Helsinki"
                        },
                        {
                          key: "montreal",
                          text: "Montreal"
                        }
                      ],
                      required: true
                    },
                    {
                      id: "Sign",
                      title: "Signed",
                      type: CustomCollectionFieldType.boolean
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

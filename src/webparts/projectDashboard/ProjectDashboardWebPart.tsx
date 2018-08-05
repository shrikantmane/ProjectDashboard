import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import 'core-js/es6/number'; 
import 'core-js/es6/array';
import 'core-js/es6/symbol';

require('./ProjectDashbaord.overrides.scss');
import { BrowserRouter, Route, Switch, DefaultRoute  } from 'react-router-dom';
import * as strings from 'ProjectDashboardWebPartStrings';
import ProjectDashboard from './components/ProjectDashboard';
import { IProjectDashboardProps } from './components/IProjectDashboardProps';
import { sp } from '@pnp/sp';
import { Environment, EnvironmentType} from '@microsoft/sp-core-library';

export interface IProjectDashboardWebPartProps {
  list: string;
}

let spCurrentPageUrl : string;

export default class ProjectDashboardWebPart extends BaseClientSideWebPart<IProjectDashboardWebPartProps> {

  public onInit(): Promise<void> {

    if(Environment.type == EnvironmentType.ClassicSharePoint){   //Classic SharePoint page

    }else if(Environment.type === EnvironmentType.Local){        //Workbenck page
      spCurrentPageUrl =  window.location.pathname;
      return Promise.resolve();
    }else if(Environment.type === EnvironmentType.SharePoint){   //Modern SharePoint page 
        spCurrentPageUrl= "/sites/hbctest/SitePages/Dashboard.aspx";
        //spCurrentPageUrl= "/sites/hbctest/_layouts/15/workbench.aspx";
      return Promise.resolve();
    }else if(Environment.type === EnvironmentType.Test){         //Running on Unit test enveironment 
      return Promise.resolve();
    }

    return super.onInit().then(_ => {
      // establish SPFx context
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IProjectDashboardProps > = React.createElement(      
      ProjectDashboard,
      {
        list: this.properties.list,
        context: this.context
      }
    );    
    ReactDom.render( 
      <BrowserRouter basename={spCurrentPageUrl}>
        {element}
      </BrowserRouter>, this.domElement);    
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmployeeDashboardWebPartStrings';
import {
  EmployeeDashboard,
  IEmployeeDashboardProps,
  ITask,
  MAX_COLUMNS,
  MIN_COLUMNS,
  SAMPLE_TASKS,
  clampColumns
} from '../../components';

export interface IEmployeeDashboardWebPartProps {
  isDarkTheme: boolean;
  quickLinksColumns: number;
  showSampleTasks: boolean;
}

export default class EmployeeDashboardWebPart extends BaseClientSideWebPart<IEmployeeDashboardWebPartProps> {

  public render(): void {
    const { user } = this.context.pageContext;
    const photoUrl: string = user.email
      ? `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(user.email)}`
      : '';

    const element: React.ReactElement<IEmployeeDashboardProps> = React.createElement(EmployeeDashboard, {
      user: {
        displayName: user.displayName,
        email: user.email,
        photoUrl: photoUrl || undefined
      },
      isDarkTheme: !!this.properties.isDarkTheme,
      quickLinksColumns: clampColumns(this.properties.quickLinksColumns),
      // Replace with tasks from Microsoft Graph (To Do / Planner) when a task source is connected.
      tasks: this.properties.showSampleTasks === false ? [] : SAMPLE_TASKS,
      onTaskToggle: (task: ITask): void => {
        // Persist the change here once tasks come from a real source.
        console.debug('[EmployeeDashboard] task toggled', task.id, task.completed);
      }
    });

    ReactDom.render(element, this.domElement);
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneToggle('isDarkTheme', {
                  label: strings.DarkModeLabel,
                  onText: strings.On,
                  offText: strings.Off
                })
              ]
            },
            {
              groupName: strings.QuickLinksGroupName,
              groupFields: [
                PropertyPaneSlider('quickLinksColumns', {
                  label: strings.ColumnsLabel,
                  min: MIN_COLUMNS,
                  max: MAX_COLUMNS,
                  step: 1,
                  showValue: true
                })
              ]
            },
            {
              groupName: strings.TasksGroupName,
              groupFields: [
                PropertyPaneToggle('showSampleTasks', {
                  label: strings.SampleTasksLabel,
                  onText: strings.On,
                  offText: strings.Off
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

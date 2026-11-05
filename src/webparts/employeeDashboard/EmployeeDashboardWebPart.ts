import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
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
import { TODO_WEB_URL, TodoTaskService } from '../../services/TodoTaskService';

export type TaskSource = 'todo' | 'sample';

export interface IEmployeeDashboardWebPartProps {
  isDarkTheme: boolean;
  quickLinksColumns: number;
  /** Missing on instances added before v1.1, which defaults to To Do. */
  taskSource?: TaskSource;
}

export default class EmployeeDashboardWebPart extends BaseClientSideWebPart<IEmployeeDashboardWebPartProps> {
  private _todo: TodoTaskService | undefined;
  private _tasks: ITask[] = [];
  private _tasksLoading: boolean = false;
  private _tasksError: string | undefined;
  /** Ignores responses from an older load if a newer one has started. */
  private _loadId: number = 0;

  private get _taskSource(): TaskSource {
    return this.properties.taskSource === 'sample' ? 'sample' : 'todo';
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Load tasks after the first render so the page isn't held up by Graph.
    this._loadTasks().catch(() => undefined);
  }

  public render(): void {
    const { user } = this.context.pageContext;
    const photoUrl: string = user.email
      ? `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(user.email)}`
      : '';
    const fromTodo = this._taskSource === 'todo';

    const element: React.ReactElement<IEmployeeDashboardProps> = React.createElement(EmployeeDashboard, {
      user: {
        displayName: user.displayName,
        email: user.email,
        photoUrl: photoUrl || undefined
      },
      isDarkTheme: !!this.properties.isDarkTheme,
      quickLinksColumns: clampColumns(this.properties.quickLinksColumns),
      tasks: fromTodo ? this._tasks : SAMPLE_TASKS,
      tasksLoading: fromTodo && this._tasksLoading,
      tasksError: fromTodo ? this._tasksError : undefined,
      onRetryTasks: fromTodo ? () => { this._loadTasks().catch(() => undefined); } : undefined,
      tasksSourceName: fromTodo ? strings.TodoSourceName : strings.SampleSourceName,
      tasksSourceUrl: fromTodo ? TODO_WEB_URL : undefined,
      onTaskToggle: fromTodo ? (task: ITask) => this._saveTask(task) : undefined
    });

    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'taskSource' && newValue === 'todo' && oldValue !== newValue) {
      this._loadTasks().catch(() => undefined);
    }
  }

  private async _loadTasks(): Promise<void> {
    if (this._taskSource !== 'todo') return;
    const loadId = ++this._loadId;
    this._tasksLoading = true;
    this._tasksError = undefined;
    this.render();

    try {
      if (!this._todo) {
        this._todo = new TodoTaskService(await this.context.msGraphClientFactory.getClient('3'));
      }
      const tasks = await this._todo.getTasks();
      if (loadId !== this._loadId) return;
      this._tasks = tasks;
    } catch (err) {
      if (loadId !== this._loadId) return;
      console.error('[EmployeeDashboard] Could not load To Do tasks', err);
      this._tasks = [];
      this._tasksError = strings.TodoLoadError;
    }
    this._tasksLoading = false;
    this.render();
  }

  private async _saveTask(task: ITask): Promise<void> {
    if (!this._todo) throw new Error('To Do is not connected');
    await this._todo.setCompleted(task);
    // Keep our copy in step, so a re-render (e.g. a property change) doesn't undo the tick.
    this._tasks = this._tasks.map((t) => (t.id === task.id ? task : t));
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
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
                PropertyPaneDropdown('taskSource', {
                  label: strings.TaskSourceLabel,
                  selectedKey: this._taskSource,
                  options: [
                    { key: 'todo', text: strings.TodoSourceName },
                    { key: 'sample', text: strings.SampleSourceName }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

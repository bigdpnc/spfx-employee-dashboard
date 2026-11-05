import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context';

import * as strings from 'EmployeeDashboardWebPartStrings';
import {
  DEFAULT_QUICK_LINKS,
  EmployeeDashboard,
  IDashboardNotice,
  IEmployeeDashboardProps,
  IQuickLink,
  ITask,
  MAX_COLUMNS,
  MIN_COLUMNS,
  SAMPLE_TASKS,
  clampColumns
} from '../../components';
import { TODO_WEB_URL, TodoTaskService, describeTodoError } from '../../services/TodoTaskService';
import { DEFAULT_QUICK_LINKS_LIST, ISeedLink, QuickLinksService } from '../../services/QuickLinksService';

export type TaskSource = 'todo' | 'sample';

export interface IEmployeeDashboardWebPartProps {
  isDarkTheme: boolean;
  quickLinksColumns: number;
  /** Missing on instances added before v1.1, which defaults to To Do. */
  taskSource?: TaskSource;
  /** Title of the SharePoint list on this site that holds the Quick Links. */
  quickLinksListTitle?: string;
}

/** Example links become the list's starter rows; placeholder (#...) URLs start hidden. */
const SEED_LINKS: ISeedLink[] = DEFAULT_QUICK_LINKS.map((l) => ({
  title: l.title,
  description: l.description,
  url: l.url.charAt(0) === '#' ? undefined : l.url,
  icon: l.icon,
  color: l.accent
}));

export default class EmployeeDashboardWebPart extends BaseClientSideWebPart<IEmployeeDashboardWebPartProps> {
  private _todo: TodoTaskService | undefined;
  private _tasks: ITask[] = [];
  private _tasksLoading: boolean = false;
  private _tasksError: string | undefined;
  /** Ignores responses from an older load if a newer one has started. */
  private _loadId: number = 0;

  private _links: QuickLinksService | undefined;
  private _quickLinks: IQuickLink[] = [];
  private _linksLoading: boolean = true;
  private _linksState: 'ok' | 'missing' | 'error' = 'ok';
  private _linksError: string | undefined;
  private _linksListUrl: string | undefined;
  private _creatingList: boolean = false;
  private _createError: string | undefined;
  private _linksLoadId: number = 0;

  private get _listTitle(): string {
    return (this.properties.quickLinksListTitle || '').trim() || DEFAULT_QUICK_LINKS_LIST;
  }

  private _can(permission: SPPermission): boolean {
    const perms = this.context.pageContext.web.permissions;
    return !!perms && perms.hasPermission(permission);
  }

  private get _taskSource(): TaskSource {
    return this.properties.taskSource === 'sample' ? 'sample' : 'todo';
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._links = new QuickLinksService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    // Start loading without holding up the page; the first render shows the loading state.
    this._loadTasks().catch(() => undefined);
    this._loadLinks().catch(() => undefined);
  }

  /** SharePoint must do the first render itself; calling render() during onInit breaks the page. */
  private _rerender(): void {
    if (this.renderedOnce) this.render();
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
      quickLinks: this._quickLinks,
      quickLinksLoading: this._linksLoading,
      quickLinksNotice: this._linksNotice(),
      quickLinksManageUrl:
        this._linksState === 'ok' && this._linksListUrl && this._can(SPPermission.editListItems) ? this._linksListUrl : undefined,
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
    if (propertyPath === 'quickLinksListTitle' && oldValue !== newValue) {
      this._createError = undefined;
      this._loadLinks().catch(() => undefined);
    }
  }

  private async _loadTasks(): Promise<void> {
    if (this._taskSource !== 'todo') return;
    const loadId = ++this._loadId;
    this._tasksLoading = true;
    this._tasksError = undefined;
    this._rerender();

    try {
      if (!this._todo) {
        this._todo = new TodoTaskService(await this.context.msGraphClientFactory.getClient('3'));
      }
      const tasks = await this._todo.getTasks();
      if (loadId !== this._loadId) return;
      this._tasks = tasks;
    } catch (err) {
      if (loadId !== this._loadId) return;
      const { kind, detail } = describeTodoError(err);
      console.error('[EmployeeDashboard] Could not load To Do tasks', detail, err);
      this._tasks = [];
      const message =
        kind === 'permission' ? strings.TodoPermissionError : kind === 'noMailbox' ? strings.TodoNoMailboxError : strings.TodoLoadError;
      this._tasksError = detail ? `${message} (${detail})` : message;
    }
    this._tasksLoading = false;
    this._rerender();
  }

  private _linksNotice(): IDashboardNotice | undefined {
    if (this._linksLoading) return undefined;
    if (this._createError) {
      return {
        tone: 'error',
        message: this._createError,
        actionLabel: strings.TryAgain,
        onAction: () => { this._createList().catch(() => undefined); },
        busy: this._creatingList,
        busyLabel: strings.CreatingList
      };
    }
    if (this._linksState === 'missing') {
      return this._can(SPPermission.manageLists)
        ? {
            message: strings.LinksListMissingOwner.replace('{0}', this._listTitle),
            actionLabel: strings.CreateList,
            onAction: () => { this._createList().catch(() => undefined); },
            busy: this._creatingList,
            busyLabel: strings.CreatingList
          }
        : { message: strings.LinksListMissingVisitor };
    }
    if (this._linksState === 'error') {
      return {
        tone: 'error',
        message: this._linksError || strings.LinksLoadError,
        actionLabel: strings.TryAgain,
        onAction: () => { this._loadLinks().catch(() => undefined); }
      };
    }
    return undefined;
  }

  private async _loadLinks(): Promise<void> {
    if (!this._links) return;
    const loadId = ++this._linksLoadId;
    this._linksLoading = true;
    this._rerender();
    try {
      const result = await this._links.getLinks(this._listTitle);
      if (loadId !== this._linksLoadId) return;
      if (result.status === 'missing') {
        this._linksState = 'missing';
        this._quickLinks = [];
        this._linksListUrl = undefined;
      } else {
        this._linksState = 'ok';
        this._quickLinks = result.links;
        this._linksListUrl = result.listUrl;
      }
      this._linksError = undefined;
    } catch (err) {
      if (loadId !== this._linksLoadId) return;
      console.error('[EmployeeDashboard] Could not load Quick Links', err);
      this._linksState = 'error';
      this._quickLinks = [];
      this._linksError = `${strings.LinksLoadError} (${(err as Error).message})`;
    }
    this._linksLoading = false;
    this._rerender();
  }

  private async _createList(): Promise<void> {
    if (!this._links || this._creatingList) return;
    this._creatingList = true;
    this._createError = undefined;
    this._rerender();
    try {
      await this._links.createList(this._listTitle, SEED_LINKS);
    } catch (err) {
      console.error('[EmployeeDashboard] Could not create the Quick Links list', err);
      this._createError = `${strings.CreateListError} (${(err as Error).message})`;
    }
    this._creatingList = false;
    if (this._createError) {
      this._rerender();
    } else {
      await this._loadLinks();
    }
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
    return Version.parse('1.2');
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
                PropertyPaneTextField('quickLinksListTitle', {
                  label: strings.LinksListLabel,
                  description: strings.LinksListDescription,
                  placeholder: DEFAULT_QUICK_LINKS_LIST,
                  deferredValidationTime: 800
                }),
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

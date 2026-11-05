import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ITask, TaskPriority } from '../components';

/**
 * Reads and updates the signed-in user's Microsoft To Do tasks through Microsoft Graph.
 *
 * Needs the delegated Graph permission `Tasks.ReadWrite`, requested in
 * config/package-solution.json and approved once by an admin under
 * SharePoint admin center → Advanced → API access.
 */

export const TODO_WEB_URL = 'https://to-do.office.com/tasks/';

/** Completed tasks kept per list. To Do lists can hold years of history. */
const COMPLETED_PER_LIST = 15;
/** Open tasks fetched per list (one Graph page). */
const PENDING_PER_LIST = 100;

interface IGraphDateTime {
  dateTime: string;
  timeZone: string;
}

interface IGraphCollection<T> {
  value: T[];
}

interface ITodoList {
  id: string;
  displayName: string;
  wellknownListName?: 'none' | 'defaultList' | 'flaggedEmails' | 'unknownFutureValue';
}

interface ITodoTask {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred';
  importance?: 'low' | 'normal' | 'high';
  dueDateTime?: IGraphDateTime;
  completedDateTime?: IGraphDateTime;
}

/** To Do stores due dates as a date at midnight; keep only the yyyy-mm-dd part. */
function datePart(value?: IGraphDateTime): string | undefined {
  return value && value.dateTime ? value.dateTime.substring(0, 10) : undefined;
}

function toPriority(importance?: string): TaskPriority | undefined {
  // To Do only has "important" (high); normal tasks get no badge.
  return importance === 'high' ? 'high' : undefined;
}

/** Shape of errors thrown by MSGraphClientV3 (GraphError). */
interface IGraphErrorLike {
  statusCode?: number;
  code?: string;
  message?: string;
}

function isBadRequest(err: unknown): boolean {
  return !!err && (err as IGraphErrorLike).statusCode === 400;
}

export type TodoErrorKind = 'permission' | 'noMailbox' | 'other';

/** Sorts a Graph failure into something we can explain to the user, plus a short technical detail. */
export function describeTodoError(err: unknown): { kind: TodoErrorKind; detail: string } {
  const e = (err || {}) as IGraphErrorLike;
  const status = e.statusCode;
  const code = e.code || '';
  const detail = [status ? `HTTP ${status}` : '', code, e.message || ''].filter(Boolean).join(' · ');
  if (status === 401 || status === 403 || /consent|AADSTS65001|Authorization_RequestDenied/i.test(`${code} ${e.message}`)) {
    return { kind: 'permission', detail };
  }
  if (/MailboxNotEnabledForRESTAPI|MailboxNotFound|ResourceNotFound/i.test(code) || status === 404) {
    return { kind: 'noMailbox', detail };
  }
  return { kind: 'other', detail };
}

export class TodoTaskService {
  /** Task id → list id, needed to PATCH a task. */
  private readonly _listByTask: Map<string, string> = new Map<string, string>();

  public constructor(private readonly _client: MSGraphClientV3) {}

  public async getTasks(): Promise<ITask[]> {
    const lists = await this._getLists();
    const perList = await Promise.all(lists.map((list) => this._getListTasks(list)));
    return perList.reduce((all, tasks) => all.concat(tasks), [] as ITask[]);
  }

  /** Marks the task completed or not started in To Do. */
  public async setCompleted(task: ITask): Promise<void> {
    const listId = this._listByTask.get(task.id);
    if (!listId) throw new Error(`Unknown To Do task ${task.id}`);
    await this._client
      .api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(task.id)}`)
      .patch({ status: task.completed ? 'completed' : 'notStarted' });
  }

  private async _getLists(): Promise<ITodoList[]> {
    // Note: no $select. Graph rejects it on /me/todo/lists with 400 "Invalid request".
    const res: IGraphCollection<ITodoList> = await this._client.api('/me/todo/lists').get();
    // "Flagged email" is a virtual list of Outlook flags. Leave it out so the dashboard shows real tasks.
    return res.value.filter((l) => l.wellknownListName !== 'flaggedEmails');
  }

  private async _getListTasks(list: ITodoList): Promise<ITask[]> {
    const base = `/me/todo/lists/${encodeURIComponent(list.id)}/tasks`;

    const [pending, completed] = await Promise.all([this._getPending(base), this._getRecentlyCompleted(base)]);

    return pending.concat(completed).map((t) => {
      this._listByTask.set(t.id, list.id);
      return {
        id: t.id,
        title: t.title,
        source: list.displayName,
        dueDate: datePart(t.dueDateTime),
        priority: toPriority(t.importance),
        completed: t.status === 'completed',
        completedOn: datePart(t.completedDateTime),
      };
    });
  }

  private async _getPending(base: string): Promise<ITodoTask[]> {
    try {
      const res: IGraphCollection<ITodoTask> = await this._client
        .api(base)
        .filter("status ne 'completed'")
        .top(PENDING_PER_LIST)
        .get();
      return res.value;
    } catch (err) {
      if (!isBadRequest(err)) throw err;
      // Query option rejected: fetch a page unfiltered and filter here instead.
      const res: IGraphCollection<ITodoTask> = await this._client.api(base).top(PENDING_PER_LIST).get();
      return res.value.filter((t) => t.status !== 'completed');
    }
  }

  private async _getRecentlyCompleted(base: string): Promise<ITodoTask[]> {
    const attempts: Array<() => Promise<IGraphCollection<ITodoTask>>> = [
      // Newest first, so the Completed tab shows recent work rather than the oldest history.
      () => this._client.api(base).filter("status eq 'completed'").orderby('lastModifiedDateTime desc').top(COMPLETED_PER_LIST).get(),
      () => this._client.api(base).filter("status eq 'completed'").top(COMPLETED_PER_LIST).get(),
    ];
    for (const attempt of attempts) {
      try {
        return (await attempt()).value;
      } catch (err) {
        if (!isBadRequest(err)) throw err;
      }
    }
    // Completed history is a nice-to-have: show open tasks rather than failing.
    return [];
  }
}

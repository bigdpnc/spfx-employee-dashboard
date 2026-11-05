import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DashboardIconName, IQuickLink } from '../components';

/**
 * Reads Quick Links from a SharePoint list on the current site, and can create that list
 * (columns + starter rows) for site owners. Uses the viewer's own SharePoint permissions,
 * so no Graph or API approval is needed.
 */

export const DEFAULT_QUICK_LINKS_LIST = 'Dashboard Quick Links';

/** Internal column names. Prefixed so they never clash with existing site columns. */
export const FIELDS = {
  description: 'EDDescription',
  url: 'EDUrl',
  icon: 'EDIcon',
  color: 'EDColor',
  order: 'EDOrder',
  newTab: 'EDNewTab',
  show: 'EDShow',
} as const;

export const ICON_CHOICES: DashboardIconName[] = [
  'mail', 'people', 'receipt', 'calendar', 'chat', 'laptop', 'book', 'heart', 'folder', 'chart',
];

/** A list row. Empty columns come back as null from SharePoint; the mapping only checks truthiness. */
interface IListItem {
  Id: number;
  Title: string;
  EDDescription?: string;
  EDUrl?: { Url?: string; Description?: string };
  EDIcon?: string;
  EDColor?: string;
  EDOrder?: number;
  EDNewTab?: boolean;
  EDShow?: boolean;
}

export type QuickLinksResult =
  | { status: 'ok'; links: IQuickLink[]; listUrl: string }
  | { status: 'missing' };

/** Starter rows written when the list is created. Rows without a URL start hidden. */
export interface ISeedLink {
  title: string;
  description?: string;
  url?: string;
  icon: DashboardIconName;
  color?: string;
}

/** http(s) links and site-relative paths only, so a list editor can't add `javascript:` links. */
export function isSafeUrl(url: string): boolean {
  return /^https?:\/\//i.test(url) || /^\/(?!\/)/.test(url);
}

/** Hex colours only (#rgb, #rrggbb, #rrggbbaa); anything else falls back to the theme accent. */
export function isSafeColor(color: string): boolean {
  return /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(color.trim());
}

function toIcon(value?: string): DashboardIconName {
  const v = (value || '').trim().toLowerCase() as DashboardIconName;
  return ICON_CHOICES.indexOf(v) > -1 ? v : 'folder';
}

/** Turns list rows into dashboard links: hidden, URL-less and unsafe rows are dropped, then sorted. */
export function mapItems(items: IListItem[]): IQuickLink[] {
  return items
    .filter((i) => i.EDShow !== false)
    .map((i) => ({ item: i, url: ((i.EDUrl && i.EDUrl.Url) || '').trim() }))
    .filter((x) => x.url && isSafeUrl(x.url))
    .sort((a, b) => {
      const ao = typeof a.item.EDOrder === 'number' ? a.item.EDOrder : Number.MAX_SAFE_INTEGER;
      const bo = typeof b.item.EDOrder === 'number' ? b.item.EDOrder : Number.MAX_SAFE_INTEGER;
      return ao !== bo ? ao - bo : (a.item.Title || '').localeCompare(b.item.Title || '');
    })
    .map(({ item, url }) => ({
      id: String(item.Id),
      title: item.Title || url,
      description: item.EDDescription || undefined,
      url,
      icon: toIcon(item.EDIcon),
      accent: item.EDColor && isSafeColor(item.EDColor) ? item.EDColor.trim() : undefined,
      openInNewTab: item.EDNewTab !== false,
    }));
}

const VERBOSE = {
  Accept: 'application/json;odata=verbose',
  'Content-Type': 'application/json;odata=verbose',
  'odata-version': '',
};

export class QuickLinksService {
  public constructor(
    private readonly _http: SPHttpClient,
    private readonly _webUrl: string,
  ) {}

  private _listPath(title: string): string {
    // OData string literal: single quotes are doubled.
    return `${this._webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(title.replace(/'/g, "''"))}')`;
  }

  public async getLinks(listTitle: string): Promise<QuickLinksResult> {
    const listRes = await this._http.get(
      `${this._listPath(listTitle)}?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
      SPHttpClient.configurations.v1,
    );
    if (listRes.status === 404) return { status: 'missing' };
    await ensureOk(listRes, 'read the Quick Links list');
    const list: { RootFolder: { ServerRelativeUrl: string } } = await listRes.json();

    const select = ['Id', 'Title'].concat(Object.keys(FIELDS).map((k) => FIELDS[k as keyof typeof FIELDS])).join(',');
    const itemsRes = await this._http.get(
      `${this._listPath(listTitle)}/items?$select=${select}&$top=500`,
      SPHttpClient.configurations.v1,
    );
    await ensureOk(itemsRes, 'read the Quick Links');
    const items: { value: IListItem[] } = await itemsRes.json();

    return {
      status: 'ok',
      links: mapItems(items.value),
      listUrl: `${new URL(this._webUrl).origin}${list.RootFolder.ServerRelativeUrl}`,
    };
  }

  /**
   * Creates the list, its columns, adds them to the default view, locks it down so only
   * site owners can edit it, and writes the starter rows.
   */
  public async createList(listTitle: string, seed: ISeedLink[]): Promise<void> {
    const created = await this._post(`${this._webUrl}/_api/web/lists`, {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: listTitle,
      Description: 'Links shown in the Quick Links section of the Employee Dashboard web part.',
      ContentTypesEnabled: false,
    }, 'create the Quick Links list');
    const listId: string = created.d.Id;
    const entityType: string = created.d.ListItemEntityTypeFullName;
    const listApi = `${this._webUrl}/_api/web/lists(guid'${listId}')`;

    const iconChoices = ICON_CHOICES.map((c) => `<CHOICE>${c}</CHOICE>`).join('');
    const schemas = [
      `<Field Type="Note" Name="${FIELDS.description}" StaticName="${FIELDS.description}" DisplayName="Description" NumLines="2" RichText="FALSE" />`,
      `<Field Type="URL" Name="${FIELDS.url}" StaticName="${FIELDS.url}" DisplayName="URL" Format="Hyperlink" />`,
      `<Field Type="Choice" Name="${FIELDS.icon}" StaticName="${FIELDS.icon}" DisplayName="Icon" Format="Dropdown"><Default>folder</Default><CHOICES>${iconChoices}</CHOICES></Field>`,
      `<Field Type="Text" Name="${FIELDS.color}" StaticName="${FIELDS.color}" DisplayName="Colour" MaxLength="9" Description="Hex colour for the icon tile, e.g. #2563eb. Leave blank for the theme colour." />`,
      `<Field Type="Number" Name="${FIELDS.order}" StaticName="${FIELDS.order}" DisplayName="Order" Decimals="0" Description="Lower numbers show first." />`,
      `<Field Type="Boolean" Name="${FIELDS.newTab}" StaticName="${FIELDS.newTab}" DisplayName="Open in new tab"><Default>1</Default></Field>`,
      `<Field Type="Boolean" Name="${FIELDS.show}" StaticName="${FIELDS.show}" DisplayName="Show"><Default>1</Default></Field>`,
    ];
    for (const schemaXml of schemas) {
      await this._post(`${listApi}/fields/createfieldasxml`, {
        parameters: {
          __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
          SchemaXml: schemaXml,
          Options: 8, // AddFieldInternalNameHint: keep the internal names above
        },
      }, 'add a column to the Quick Links list');
    }
    for (const field of [FIELDS.url, FIELDS.icon, FIELDS.order, FIELDS.show]) {
      await this._post(`${listApi}/DefaultView/ViewFields/addviewfield('${field}')`, undefined, 'update the list view');
    }

    await this._lockDown(listApi);

    let order = 10;
    for (const link of seed) {
      const hasUrl = !!link.url && isSafeUrl(link.url);
      await this._post(`${listApi}/items`, {
        __metadata: { type: entityType },
        Title: link.title,
        [FIELDS.description]: link.description || undefined,
        [FIELDS.url]: hasUrl ? { __metadata: { type: 'SP.FieldUrlValue' }, Url: link.url, Description: link.title } : undefined,
        [FIELDS.icon]: link.icon,
        [FIELDS.color]: link.color || undefined,
        [FIELDS.order]: order,
        [FIELDS.newTab]: true,
        // Placeholders without a real URL start hidden until someone fills them in.
        [FIELDS.show]: hasUrl,
      }, 'add a starter link');
      order += 10;
    }
  }

  /**
   * Gives the list its own permissions so only site owners can change links:
   * stop inheriting (copying the site's current assignments), then switch the
   * site's Members group from Edit to Read. Owners keep Full Control, Visitors keep Read.
   */
  private async _lockDown(listApi: string): Promise<void> {
    await this._post(`${listApi}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`, undefined, 'set permissions on the list');

    const membersRes = await this._http.get(`${this._webUrl}/_api/web/associatedmembergroup?$select=Id`, SPHttpClient.configurations.v1);
    if (membersRes.status === 404 || membersRes.status === 204) return; // site has no default Members group
    await ensureOk(membersRes, 'find the site Members group');
    const members: { Id?: number } = await membersRes.json();
    if (!members || typeof members.Id !== 'number') return;

    const readerRes = await this._http.get(`${this._webUrl}/_api/web/roledefinitions/getbytype(2)?$select=Id`, SPHttpClient.configurations.v1);
    await ensureOk(readerRes, 'find the Read permission level');
    const reader: { Id: number } = await readerRes.json();

    // Remove Members' copied assignment (usually Edit), then grant Read.
    const removeRes = await this._http.post(
      `${listApi}/roleassignments/getbyprincipalid(${members.Id})/deleteObject()`,
      SPHttpClient.configurations.v1,
      { headers: VERBOSE }
    );
    if (!removeRes.ok && removeRes.status !== 404) await ensureOk(removeRes, 'remove edit access for site members');
    await this._post(
      `${listApi}/roleassignments/addroleassignment(principalid=${members.Id},roledefid=${reader.Id})`,
      undefined,
      'give site members read access'
    );
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _post(url: string, body: unknown, action: string): Promise<any> {
    const res = await this._http.post(url, SPHttpClient.configurations.v1, {
      headers: VERBOSE,
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    await ensureOk(res, action);
    return res.status === 204 ? undefined : res.json();
  }
}

async function ensureOk(res: SPHttpClientResponse, action: string): Promise<void> {
  if (res.ok) return;
  let detail = '';
  try {
    const body = await res.json();
    detail = (body && body.error && (body.error.message?.value || body.error.message)) || '';
  } catch {
    /* body wasn't JSON */
  }
  const err = new Error(`Couldn't ${action} (HTTP ${res.status}${detail ? ` · ${detail}` : ''})`);
  (err as Error & { statusCode?: number }).statusCode = res.status;
  throw err;
}

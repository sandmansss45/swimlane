import { IPublicClientApplication } from '@azure/msal-browser';
import { IDataService } from './IDataService';
import { IProcessStep, parseDependsOn } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';
import { GraphClient } from '../auth/graphClient';
import { SHAREPOINT_SITE_HOSTNAME, SHAREPOINT_SITE_PATH } from '../auth/authConfig';

// TODO - CONFIRM AGAINST THE REAL "QLE Finance" SITE, same as the SPFx
// build's SharePointDataService - swap in the real display names once
// known (or run logFieldSchema/resolveListId manually from the browser
// console to check).
const PROCESS_LIST_TITLE = 'TODO-CONFIRM-PROCESS-LIST-TITLE';
const EMPLOYEES_LIST_TITLE = 'TODO-CONFIRM-EMPLOYEES-LIST-TITLE';
const RISK_LIST_TITLE = 'risk regnew';

type FieldMap = { [displayName: string]: string };
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type GraphItem = { id: string; fields: Record<string, any> };

export class GraphDataService implements IDataService {
  private _graph: GraphClient;
  private _siteId: string | undefined;
  private _listIdCache = new Map<string, string>();
  private _fieldMapCache = new Map<string, FieldMap>();

  constructor(msal: IPublicClientApplication) {
    this._graph = new GraphClient(msal);
  }

  private async _resolveSiteId(): Promise<string> {
    if (this._siteId) return this._siteId;
    const site = await this._graph.get<{ id: string }>(`/sites/${SHAREPOINT_SITE_HOSTNAME}:${SHAREPOINT_SITE_PATH}`);
    this._siteId = site.id;
    return site.id;
  }

  private async _resolveListId(listTitle: string): Promise<string> {
    const cached = this._listIdCache.get(listTitle);
    if (cached) return cached;
    const siteId = await this._resolveSiteId();
    const lists = await this._graph.getAllPages<{ id: string; displayName: string }>(
      `/sites/${siteId}/lists?$select=id,displayName`
    );
    const match = lists.find(l => l.displayName === listTitle);
    if (!match) {
      const available = lists.map(l => l.displayName).join(', ') || '(no lists found - check Sites.ReadWrite.All was actually granted)';
      throw new Error(`No list titled "${listTitle}" found on the site. Real lists on this site: ${available}`);
    }
    this._listIdCache.set(listTitle, match.id);
    return match.id;
  }

  /**
   * Resolves Display Name -> internal column name at runtime instead of
   * hardcoding guessed internal names (SharePoint mangles spaces etc.,
   * same reasoning as the SPFx build's SharePointDataService). Cached per
   * list for the session.
   */
  private async _resolveFieldMap(listTitle: string): Promise<FieldMap> {
    const cached = this._fieldMapCache.get(listTitle);
    if (cached) return cached;
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(listTitle);
    const columns = await this._graph.getAllPages<{ name: string; displayName: string }>(
      `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`
    );
    const map: FieldMap = {};
    columns.forEach(c => { map[c.displayName] = c.name; });
    this._fieldMapCache.set(listTitle, map);
    return map;
  }

  private static _get(item: GraphItem, fieldMap: FieldMap, displayName: string): string {
    const internalName = fieldMap[displayName];
    if (!internalName) {
      // eslint-disable-next-line no-console
      console.warn(`[SwimlaneStudio] No field found for display name "${displayName}" - check the list's real columns.`);
      return '';
    }
    const value = item.fields[internalName];
    return value === null || value === undefined ? '' : String(value);
  }

  private async _getItems(listTitle: string): Promise<GraphItem[]> {
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(listTitle);
    const items = await this._graph.getAllPages<GraphItem>(
      `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`
    );
    // orderBy is not reliable across Graph pages, and DependsOn tokens are
    // resolved by row number against original import order (header
    // counted as row 1) - sort client-side by numeric item ID so that
    // order is guaranteed regardless of what order Graph returned pages in.
    return items.slice().sort((a, b) => Number(a.id) - Number(b.id));
  }

  public async getProcessSteps(): Promise<IProcessStep[]> {
    const fieldMap = await this._resolveFieldMap(PROCESS_LIST_TITLE);
    const items = await this._getItems(PROCESS_LIST_TITLE);
    // eslint-disable-next-line no-console
    console.log(`[SwimlaneStudio] "${PROCESS_LIST_TITLE}" live row count: ${items.length} - confirm this matches the list in SharePoint.`);

    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): IProcessStep => ({
      id: item.id,
      apqcTitle: get(item, 'APQC Title'),
      processDescription: get(item, 'Process Description'),
      processStepId: get(item, 'Process Step ID'),
      processStepName: get(item, 'Process Step Name'),
      actionType: get(item, 'Action Type'),
      action: get(item, 'Action'),
      actionDescription: get(item, 'Action Description'),
      responsibleJobTitle: get(item, 'ResponsibleJobTitle'),
      shapeOverride: get(item, 'ShapeOverride'),
      dependsOn: parseDependsOn(get(item, 'DependsOn'))
    }));
  }

  public async getEmployees(): Promise<IEmployee[]> {
    const fieldMap = await this._resolveFieldMap(EMPLOYEES_LIST_TITLE);
    const items = await this._getItems(EMPLOYEES_LIST_TITLE);
    // eslint-disable-next-line no-console
    console.log(`[SwimlaneStudio] "${EMPLOYEES_LIST_TITLE}" live row count: ${items.length} - confirm this matches the list in SharePoint.`);

    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): IEmployee => ({
      id: item.id,
      name: get(item, 'Title'),
      jobTitle: get(item, 'Position'),
      region: get(item, 'Region') || undefined
    }));
  }

  public async getRiskStatements(): Promise<IRiskStatement[]> {
    const fieldMap = await this._resolveFieldMap(RISK_LIST_TITLE);
    const items = await this._getItems(RISK_LIST_TITLE);
    // eslint-disable-next-line no-console
    console.log(`[SwimlaneStudio] "${RISK_LIST_TITLE}" live row count: ${items.length} - confirm this matches the list in SharePoint.`);

    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): IRiskStatement => ({
      id: item.id,
      title: get(item, 'Title'),
      riskStatement: get(item, 'Risk Statement')
    }));
  }

  public async addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep> {
    const fieldMap = await this._resolveFieldMap(PROCESS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROCESS_LIST_TITLE);

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('APQC Title', step.apqcTitle);
    set('Process Description', step.processDescription);
    set('Process Step ID', step.processStepId);
    set('Process Step Name', step.processStepName);
    set('Action Type', step.actionType);
    set('Action', step.action);
    set('Action Description', step.actionDescription);
    set('ResponsibleJobTitle', step.responsibleJobTitle);
    set('ShapeOverride', step.shapeOverride || '');
    set('DependsOn', step.dependsOn.join(', '));

    const created = await this._graph.post<GraphItem>(`/sites/${siteId}/lists/${listId}/items`, { fields });
    return { ...step, id: created.id };
  }

  public async updateProcessStep(step: IProcessStep): Promise<void> {
    const fieldMap = await this._resolveFieldMap(PROCESS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROCESS_LIST_TITLE);

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('APQC Title', step.apqcTitle);
    set('Process Description', step.processDescription);
    set('Process Step ID', step.processStepId);
    set('Process Step Name', step.processStepName);
    set('Action Type', step.actionType);
    set('Action', step.action);
    set('Action Description', step.actionDescription);
    set('ResponsibleJobTitle', step.responsibleJobTitle);
    set('ShapeOverride', step.shapeOverride || '');
    set('DependsOn', step.dependsOn.join(', '));

    await this._graph.patch(`/sites/${siteId}/lists/${listId}/items/${step.id}/fields`, fields);
  }

  public async deleteProcessStep(id: string): Promise<void> {
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROCESS_LIST_TITLE);
    await this._graph.delete(`/sites/${siteId}/lists/${listId}/items/${id}`);
  }
}

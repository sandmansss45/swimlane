import { IPublicClientApplication } from '@azure/msal-browser';
import { IDataService } from './IDataService';
import { IProcessStep, parseDependsOn, parseEdgeLabels, serializeEdgeLabels } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement, parseLinkedRisks, serializeLinkedRisks } from '../models/IRiskStatement';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';
import { IProgressIdLabel } from '../models/IProgressIdLabel';
import { IProgressIdLock } from '../models/IProgressIdLock';
import { ISwimlaneComment } from '../models/ISwimlaneComment';
import { GraphClient } from '../auth/graphClient';
import { SHAREPOINT_SITE_HOSTNAME, SHAREPOINT_SITE_PATH } from '../auth/authConfig';

// CONFIRMED 2026-08-13 against the real "Swimlane Studio" site - a list
// with exactly the expected columns (APQC Title, Process Description,
// Process Step ID, Process Step Name, Action Type, Action, Action
// Description, ResponsibleJobTitle, ShapeOverride, DependsOn), already
// populated with real data. Renamed from "9.6 tester" to "Master File"
// 2026-08-19 - list resolution is by exact display name (see
// _resolveListId below), so this constant has to be kept in lockstep
// with whatever the real list is actually called on the site right now.
const PROCESS_LIST_TITLE = 'Master File';
// CONFIRMED 2026-08-17 - real list on the site, columns Display name,
// Department, Job title, Reports to, Start date, Hobbies. Only Job title
// and Department are used - see the comments in models/IEmployee.ts for
// why Display name isn't mapped at all.
const EMPLOYEES_LIST_TITLE = 'QLE Existing Organisation';
// CONFIRMED 2026-08-17 against a live screenshot of the real list -
// columns Risk ID, Category, Risk Statement, Root Cause, Likelihood (P),
// Materiality ($Mn), Inherent Risk Rating, Risk Response. This is a
// standing enterprise register the app reads from, not one it owns - see
// the schema comment in models/IRiskStatement.ts.
const RISK_LIST_TITLE = 'risk register data';
// CONFIRMED 2026-08-17 - created on the real "Swimlane Studio" site with
// the built-in Title column (group name) plus a single line of text
// column "Group ID". Stores names for Process Groups the static
// APQC_PROCESS_GROUP_NAMES table (apqcHierarchy.ts) doesn't already cover.
const PROCESS_GROUP_LABELS_LIST_TITLE = 'Process Group labels';
// CONFIRMED 2026-08-18 - created on the real "Swimlane Studio" site, same
// shape as "Process Group labels": the built-in Title column (the
// Progress ID's name) plus a single line of text column "Progress ID".
// Stores names for a Progress ID created as an empty shell (see "+ Add
// new progress ID") before it has any real steps of its own to derive a
// name from.
const PROGRESS_ID_LABELS_LIST_TITLE = 'Progress ID labels';
// CONFIRMED 2026-08-18 - created on the real "Swimlane Studio" site, with
// the built-in Title column (unused - left blank) plus single line of
// text columns "Progress ID", "Region", "Locked By", "Locked At",
// "Reason", "Unlocked By", "Unlocked At", "Unlock Reason". Append-only
// audit trail for swimlane sign-off/locking - see models/IProgressIdLock
// for why this is never edited in place except to fill in the three
// Unlocked* columns once, on unlock.
const PROGRESS_ID_LOCKS_LIST_TITLE = 'Progress ID locks';
// CONFIRMED 2026-08-19 - created on the real site, same shape as
// "Progress ID locks" above - the built-in Title column (unused - left
// blank) plus single line of text columns "Progress ID", "Region",
// "Author", "Comment", "Posted At". Append-only feedback log - see
// models/ISwimlaneComment - never edited or deleted once posted, so
// there's no update/delete method here at all, unlike the locks list.
const SWIMLANE_COMMENTS_LIST_TITLE = 'Swimlane comments';

type FieldMap = { [displayName: string]: string };
// createdBy/lastModifiedBy/createdDateTime/lastModifiedDateTime are
// SharePoint's own native, system-managed item metadata - always present
// on a real list item, distinct from the custom `fields` object. Never
// written by this app (see the schema comment on IProcessStep.createdBy)
// - only ever read.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type GraphItem = {
  id: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  fields: Record<string, any>;
  createdBy?: { user?: { displayName?: string; email?: string } };
  lastModifiedBy?: { user?: { displayName?: string; email?: string } };
  createdDateTime?: string;
  lastModifiedDateTime?: string;
};

export class GraphDataService implements IDataService {
  private _graph: GraphClient;
  private _siteId: string | undefined;
  private _listIdCache = new Map<string, string>();
  private _fieldMapCache = new Map<string, FieldMap>();

  // Only used for the same-session OPTIMISTIC stamp on newly-created steps
  // (see addProcessStep) - never written to SharePoint as a real column,
  // since createdBy/lastModifiedBy above are already accurate, system-
  // managed native fields. Purely so a step someone just added shows
  // "created by [them]" immediately, without waiting for the next reload
  // to re-fetch the true native value SharePoint independently records
  // anyway.
  private _currentUserName: string;

  constructor(msal: IPublicClientApplication, currentUserName: string) {
    this._graph = new GraphClient(msal);
    this._currentUserName = currentUserName;
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
    const columns = await this._graph.getAllPages<{ name: string; displayName: string; readOnly?: boolean }>(
      `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,readOnly`
    );
    const map: FieldMap = {};
    columns.forEach(c => {
      // Every SharePoint list has a hidden, computed "LinkTitle" column
      // (sometimes "LinkTitleNoMenu") that renders the real Title field as
      // a clickable link in default views - it reports the SAME display
      // name ("Title") as the real, editable Title column. Building this
      // map naively let whichever one the API happened to return last win,
      // which could silently point a 'Title' write at the read-only
      // LinkTitle column instead - Graph then rejects the write with
      // "Field 'LinkTitle' is read-only" (403). Skipping every read-only
      // column here means a display name always resolves to the one
      // column that's actually safe to write to, for any list, not just
      // this specific Title case.
      if (c.readOnly) return;
      map[c.displayName] = c.name;
    });
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

  // Graph's identitySet shape for createdBy/lastModifiedBy - falls back to
  // email when displayName isn't populated (rare, but seen for some
  // service-principal-driven writes), and to '' (never shown) rather than
  // throwing when neither is present.
  private static _identityName(identity: { user?: { displayName?: string; email?: string } } | undefined): string {
    return identity?.user?.displayName || identity?.user?.email || '';
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
      // CONFIRMED 2026-08-18 - "Region" column created on the real "9.6
      // tester" list (single line of text, values like "UK"/"US"/"SA") -
      // see FlowRegionTabs and the schema comment on IProcessStep.region
      // for why this is a separate concept from the employee Department
      // field.
      region: get(item, 'Region') || undefined,
      shapeOverride: get(item, 'ShapeOverride'),
      // TODO-CONFIRM: guessed display names, not verified against the real
      // process list yet - same caveat as ShapeOverride's own history.
      riskLevelOverride: get(item, 'RiskLevelOverride'),
      manualOrder: (() => {
        const raw = get(item, 'ManualOrder');
        const parsed = raw ? parseFloat(raw) : NaN;
        return isNaN(parsed) ? undefined : parsed;
      })(),
      dependsOn: parseDependsOn(get(item, 'DependsOn')),
      // TODO-CONFIRM: "Linked Risks" doesn't exist on the real Master File
      // list yet - needs creating as a single line of text column, same as
      // DependsOn - see the schema comment on IProcessStep.linkedRisks for
      // the "riskId:severity" format it expects.
      linkedRisks: parseLinkedRisks(get(item, 'Linked Risks')),
      // CONFIRMED 2026-08-19 - "Edge Labels" column created on the real
      // Master File list (single line of text), same format as DependsOn/
      // Linked Risks - see parseEdgeLabels in IProcessStep.ts for the
      // "token:label" format it expects.
      edgeLabels: (() => {
        const parsed = parseEdgeLabels(get(item, 'Edge Labels'));
        return Object.keys(parsed).length > 0 ? parsed : undefined;
      })(),
      // TODO-CONFIRM: "SOP Link" and "Delegation of Authority Link" don't
      // exist on the real Master File list yet - guessed display names,
      // needs creating as two single line of text columns (not Hyperlink -
      // that column type returns a {Url, Description} object from Graph
      // instead of a plain string, which _get/set here aren't set up to
      // handle, unlike every other text-ish field in this list). See the
      // schema comment on IProcessStep.sopLink for what each is for.
      sopLink: get(item, 'SOP Link') || undefined,
      delegationOfAuthorityLink: get(item, 'Delegation of Authority Link') || undefined,
      // Native SharePoint item metadata, not a custom column - see the
      // GraphItem type comment and IProcessStep.createdBy for why.
      createdBy: GraphDataService._identityName(item.createdBy),
      createdAt: item.createdDateTime,
      modifiedBy: GraphDataService._identityName(item.lastModifiedBy),
      modifiedAt: item.lastModifiedDateTime
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
      jobTitle: get(item, 'Job title'),
      department: get(item, 'Department') || undefined
    }));
  }

  public async getRiskStatements(): Promise<IRiskStatement[]> {
    const fieldMap = await this._resolveFieldMap(RISK_LIST_TITLE);
    const items = await this._getItems(RISK_LIST_TITLE);
    // eslint-disable-next-line no-console
    console.log(`[SwimlaneStudio] "${RISK_LIST_TITLE}" live row count: ${items.length} - confirm this matches the list in SharePoint.`);

    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);
    const getNumber = (item: GraphItem, displayName: string): number | undefined => {
      const raw = get(item, displayName);
      if (!raw) return undefined;
      const parsed = parseFloat(raw);
      return isNaN(parsed) ? undefined : parsed;
    };

    return items.map((item): IRiskStatement => ({
      id: item.id,
      riskId: get(item, 'Risk ID'),
      category: get(item, 'Category'),
      riskStatement: get(item, 'Risk Statement'),
      rootCause: get(item, 'Root Cause'),
      likelihood: getNumber(item, 'Likelihood (P)'),
      materiality: getNumber(item, 'Materiality ($Mn)'),
      inherentRiskRating: getNumber(item, 'Inherent Risk Rating'),
      riskResponse: get(item, 'Risk Response')
    }));
  }

  public async getProcessGroupLabels(): Promise<IProcessGroupLabel[]> {
    const fieldMap = await this._resolveFieldMap(PROCESS_GROUP_LABELS_LIST_TITLE);
    const items = await this._getItems(PROCESS_GROUP_LABELS_LIST_TITLE);
    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): IProcessGroupLabel => ({
      id: item.id,
      groupId: get(item, 'Group ID'),
      name: get(item, 'Title')
    }));
  }

  public async addProcessGroupLabel(groupId: string, name: string): Promise<IProcessGroupLabel> {
    const fieldMap = await this._resolveFieldMap(PROCESS_GROUP_LABELS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROCESS_GROUP_LABELS_LIST_TITLE);

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('Group ID', groupId);
    set('Title', name);

    const created = await this._graph.post<GraphItem>(`/sites/${siteId}/lists/${listId}/items`, { fields });
    return { id: created.id, groupId, name };
  }

  public async updateProcessGroupLabel(id: string, name: string): Promise<void> {
    const fieldMap = await this._resolveFieldMap(PROCESS_GROUP_LABELS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROCESS_GROUP_LABELS_LIST_TITLE);
    const titleField = fieldMap['Title'];
    if (!titleField) return;
    await this._graph.patch(`/sites/${siteId}/lists/${listId}/items/${id}/fields`, { [titleField]: name });
  }

  public async getProgressIdLabels(): Promise<IProgressIdLabel[]> {
    const fieldMap = await this._resolveFieldMap(PROGRESS_ID_LABELS_LIST_TITLE);
    const items = await this._getItems(PROGRESS_ID_LABELS_LIST_TITLE);
    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): IProgressIdLabel => ({
      id: item.id,
      progressId: get(item, 'Progress ID'),
      name: get(item, 'Title')
    }));
  }

  public async addProgressIdLabel(progressId: string, name: string): Promise<IProgressIdLabel> {
    const fieldMap = await this._resolveFieldMap(PROGRESS_ID_LABELS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROGRESS_ID_LABELS_LIST_TITLE);

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('Progress ID', progressId);
    set('Title', name);

    const created = await this._graph.post<GraphItem>(`/sites/${siteId}/lists/${listId}/items`, { fields });
    return { id: created.id, progressId, name };
  }

  public async updateProgressIdLabel(id: string, name: string): Promise<void> {
    const fieldMap = await this._resolveFieldMap(PROGRESS_ID_LABELS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROGRESS_ID_LABELS_LIST_TITLE);
    const titleField = fieldMap['Title'];
    if (!titleField) return;
    await this._graph.patch(`/sites/${siteId}/lists/${listId}/items/${id}/fields`, { [titleField]: name });
  }

  public async getProgressIdLocks(): Promise<IProgressIdLock[]> {
    const fieldMap = await this._resolveFieldMap(PROGRESS_ID_LOCKS_LIST_TITLE);
    const items = await this._getItems(PROGRESS_ID_LOCKS_LIST_TITLE);
    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): IProgressIdLock => ({
      id: item.id,
      progressId: get(item, 'Progress ID'),
      region: get(item, 'Region'),
      lockedBy: get(item, 'Locked By'),
      lockedAt: get(item, 'Locked At'),
      reason: get(item, 'Reason'),
      unlockedBy: get(item, 'Unlocked By'),
      unlockedAt: get(item, 'Unlocked At'),
      unlockReason: get(item, 'Unlock Reason')
    }));
  }

  public async lockProgressId(progressId: string, region: string, lockedBy: string, reason: string): Promise<IProgressIdLock> {
    const fieldMap = await this._resolveFieldMap(PROGRESS_ID_LOCKS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROGRESS_ID_LOCKS_LIST_TITLE);
    const lockedAt = new Date().toISOString();

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('Progress ID', progressId);
    set('Region', region);
    set('Locked By', lockedBy);
    set('Locked At', lockedAt);
    set('Reason', reason);

    const created = await this._graph.post<GraphItem>(`/sites/${siteId}/lists/${listId}/items`, { fields });
    return { id: created.id, progressId, region, lockedBy, lockedAt, reason, unlockedBy: '', unlockedAt: '', unlockReason: '' };
  }

  public async unlockProgressId(id: string, unlockedBy: string, reason: string): Promise<void> {
    const fieldMap = await this._resolveFieldMap(PROGRESS_ID_LOCKS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROGRESS_ID_LOCKS_LIST_TITLE);

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('Unlocked By', unlockedBy);
    set('Unlocked At', new Date().toISOString());
    set('Unlock Reason', reason);

    await this._graph.patch(`/sites/${siteId}/lists/${listId}/items/${id}/fields`, fields);
  }

  public async getSwimlaneComments(): Promise<ISwimlaneComment[]> {
    const fieldMap = await this._resolveFieldMap(SWIMLANE_COMMENTS_LIST_TITLE);
    const items = await this._getItems(SWIMLANE_COMMENTS_LIST_TITLE);
    const get = (item: GraphItem, displayName: string): string => GraphDataService._get(item, fieldMap, displayName);

    return items.map((item): ISwimlaneComment => ({
      id: item.id,
      progressId: get(item, 'Progress ID'),
      region: get(item, 'Region'),
      author: get(item, 'Author'),
      comment: get(item, 'Comment'),
      postedAt: get(item, 'Posted At')
    }));
  }

  public async addSwimlaneComment(progressId: string, region: string, author: string, comment: string): Promise<ISwimlaneComment> {
    const fieldMap = await this._resolveFieldMap(SWIMLANE_COMMENTS_LIST_TITLE);
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(SWIMLANE_COMMENTS_LIST_TITLE);
    const postedAt = new Date().toISOString();

    const fields: Record<string, string> = {};
    const set = (displayName: string, value: string): void => {
      const internalName = fieldMap[displayName];
      if (internalName) fields[internalName] = value;
    };
    set('Progress ID', progressId);
    set('Region', region);
    set('Author', author);
    set('Comment', comment);
    set('Posted At', postedAt);

    const created = await this._graph.post<GraphItem>(`/sites/${siteId}/lists/${listId}/items`, { fields });
    return { id: created.id, progressId, region, author, comment, postedAt };
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
    set('Region', step.region || '');
    set('ShapeOverride', step.shapeOverride || '');
    set('RiskLevelOverride', step.riskLevelOverride || '');
    set('ManualOrder', step.manualOrder !== undefined ? String(step.manualOrder) : '');
    set('DependsOn', step.dependsOn.join(', '));
    set('Linked Risks', serializeLinkedRisks(step.linkedRisks || []));
    set('Edge Labels', serializeEdgeLabels(step.edgeLabels));
    set('SOP Link', step.sopLink || '');
    set('Delegation of Authority Link', step.delegationOfAuthorityLink || '');

    const created = await this._graph.post<GraphItem>(`/sites/${siteId}/lists/${listId}/items`, { fields });
    const now = new Date().toISOString();
    return {
      ...step, id: created.id,
      createdBy: this._currentUserName, createdAt: now,
      modifiedBy: this._currentUserName, modifiedAt: now
    };
  }

  public async addProcessSteps(steps: Array<Omit<IProcessStep, 'id'>>): Promise<IProcessStep[]> {
    // Sequential, not Promise.all - row order must match creation order
    // (later rows' DependsOn tokens are meaningless if rows land out of
    // order), and it keeps well clear of Graph's per-request throttling.
    const created: IProcessStep[] = [];
    for (const step of steps) {
      created.push(await this.addProcessStep(step));
    }
    return created;
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
    set('Region', step.region || '');
    set('ShapeOverride', step.shapeOverride || '');
    set('RiskLevelOverride', step.riskLevelOverride || '');
    set('ManualOrder', step.manualOrder !== undefined ? String(step.manualOrder) : '');
    set('DependsOn', step.dependsOn.join(', '));
    set('Linked Risks', serializeLinkedRisks(step.linkedRisks || []));
    set('Edge Labels', serializeEdgeLabels(step.edgeLabels));
    set('SOP Link', step.sopLink || '');
    set('Delegation of Authority Link', step.delegationOfAuthorityLink || '');

    await this._graph.patch(`/sites/${siteId}/lists/${listId}/items/${step.id}/fields`, fields);
  }

  public async deleteProcessStep(id: string): Promise<void> {
    const siteId = await this._resolveSiteId();
    const listId = await this._resolveListId(PROCESS_LIST_TITLE);
    await this._graph.delete(`/sites/${siteId}/lists/${listId}/items/${id}`);
  }
}

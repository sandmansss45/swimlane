import * as React from 'react';
import {
  Spinner, MessageBar, MessageBarType, MessageBarButton, DefaultButton, PrimaryButton, IconButton, Pivot, PivotItem,
  Dialog, DialogType, DialogFooter, TextField
} from '@fluentui/react';
import styles from './SwimlaneStudio.module.scss';
import type { ISwimlaneStudioProps } from './ISwimlaneStudioProps';
import { IProcessStep, getProgressId } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';
import { IProgressIdLabel } from '../models/IProgressIdLabel';
import { IProgressIdLock, findActiveLock } from '../models/IProgressIdLock';
import { ISwimlaneComment } from '../models/ISwimlaneComment';
import { resolveDependencyEdges, stepIdsToDependsOnTokens, buildDependsOnOptions } from '../utils/dependencyResolution';
import { computeInsertOrderBefore } from '../utils/columns';
import {
  getCategoryId, getProcessGroupId, getCategoryName, getProcessGroupName, getProgressIdName,
  APQC_CATEGORY_NAMES, APQC_PROCESS_GROUP_NAMES, APQC_PROGRESS_ID_NAMES
} from '../utils/apqcHierarchy';
import HierarchyPicker from './HierarchyPicker';
import GlobalSearch from './GlobalSearch';
import ProcessStepTabs from './ProcessStepTabs';
import FlowRegionTabs from './FlowRegionTabs';
import EmployeesList from './EmployeesList';
import RiskRegisterList from './RiskRegisterList';
import AuditView from './AuditView';
import ImprovementsView from './ImprovementsView';
import SwimlaneComments from './SwimlaneComments';
import SwimlaneCanvas from './SwimlaneCanvas';
import ImportCsvModal from './ImportCsvModal';
import NewProcessModal from './NewProcessModal';
import AddHierarchyShellModal, { HierarchyShellLevel } from './AddHierarchyShellModal';
import AddStepSectionModal from './AddStepSectionModal';
import ProcessStepForm, { IProcessStepFormValue } from './ProcessStepForm';
import qleLogo from '../../assets/qle-logo.svg';

type MainTab = 'flows' | 'employees' | 'risks' | 'audit' | 'improvements';

// Single-level undo (the last destructive action only, not a full stack) -
// covers the three actions that lose data outright: deleting a step,
// bulk-deleting a whole flow, and editing (which overwrites the previous
// field values). Adding/moving a step isn't covered - a mistaken add is
// trivial to just delete, and a mistaken drag is trivial to just drag
// back. NOTE: undoing a delete re-creates the step via addProcessStep,
// which appends it as a new row rather than reinserting it at its exact
// original position - if another step's DependsOn token referenced it by
// original row number (see dependencyResolution.ts), that link won't
// perfectly restore. Pre-existing limitation of the row-number dependency
// system, not something undo makes worse - deleting already shifts every
// later row's number regardless of whether the delete is later undone.
type UndoAction =
  | { type: 'delete'; step: IProcessStep }
  | { type: 'bulkDelete'; steps: IProcessStep[] }
  | { type: 'edit'; previous: IProcessStep };

// Promise.allSettled gives back whatever the rejection actually was, not
// necessarily an Error instance - GraphDataService rejects with a real
// Error, but this stays safe even if something else throws a plain string.
function describeError(reason: unknown): string {
  return reason instanceof Error ? reason.message : String(reason);
}

const emptyStepDraft = (): IProcessStepFormValue => ({
  action: '',
  actionDescription: '',
  actionType: 'Execute (Within Limits)',
  shapeOverride: '',
  responsibleJobTitle: '',
  dependsOnStepIds: [],
  linkedRisks: []
});

const SwimlaneStudio: React.FC<ISwimlaneStudioProps> = (props) => {
  const { dataService, onSignOut, signOutLabel, currentUserName } = props;

  const [steps, setSteps] = React.useState<IProcessStep[]>([]);
  const [employees, setEmployees] = React.useState<IEmployee[]>([]);
  const [riskStatements, setRiskStatements] = React.useState<IRiskStatement[]>([]);
  const [processGroupLabels, setProcessGroupLabels] = React.useState<IProcessGroupLabel[]>([]);
  const [progressIdLabels, setProgressIdLabels] = React.useState<IProgressIdLabel[]>([]);
  const [progressIdLocks, setProgressIdLocks] = React.useState<IProgressIdLock[]>([]);
  const [swimlaneComments, setSwimlaneComments] = React.useState<ISwimlaneComment[]>([]);
  // One-shot signal telling SwimlaneCanvas to open a just-created step's
  // edit panel automatically - see handleCreateStepFromShape and the
  // matching prop comment on ISwimlaneCanvasProps.
  const [autoOpenStepId, setAutoOpenStepId] = React.useState<string | undefined>(undefined);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // APQC drill-down: Category (e.g. "9") -> Process Group (e.g. "9.6") ->
  // Progress ID (e.g. "9.6.1", the existing swimlane-per-flow level) - see
  // utils/apqcHierarchy.ts. Each level's picker is only reachable once its
  // parent is chosen, and clearing a level clears everything below it.
  const [selectedCategoryId, setSelectedCategoryId] = React.useState<string | undefined>(undefined);
  const [selectedProcessGroupId, setSelectedProcessGroupId] = React.useState<string | undefined>(undefined);
  const [selectedProgressId, setSelectedProgressId] = React.useState<string | undefined>(undefined);
  const [drilledDownStepId, setDrilledDownStepId] = React.useState<string | undefined>(undefined);
  // Which region's variant of the current Progress ID's flow is showing -
  // undefined = "All" regions combined. See FlowRegionTabs for the full
  // reasoning.
  const [selectedFlowRegion, setSelectedFlowRegion] = React.useState<string | undefined>(undefined);

  const [newStepDraft, setNewStepDraft] = React.useState<IProcessStepFormValue>(emptyStepDraft());
  const [saving, setSaving] = React.useState(false);
  const [importOpen, setImportOpen] = React.useState(false);
  const [newProcessOpen, setNewProcessOpen] = React.useState(false);
  // The lightweight "just name and reserve an ID" flow (see
  // AddHierarchyShellModal) - separate from newProcessOpen, which is the
  // full "create a complete step" flow. idPrefix is the parent ID plus a
  // trailing dot (e.g. "13." when adding a Process Group under Category
  // 13), computed fresh each time it's opened rather than reusing
  // newProcessPrefix so the two flows stay independent.
  const [addShellState, setAddShellState] = React.useState<{ level: HierarchyShellLevel; idPrefix: string } | undefined>(undefined);
  // The lightweight "add a 4th-level section" flow (see
  // AddStepSectionModal) - a real minimal step, not a label-only shell
  // like addShellState above, since a section's name only ever lives on
  // real step rows. true only once a Progress ID is actually selected
  // (there's nowhere to add a section before that).
  const [addSectionOpen, setAddSectionOpen] = React.useState(false);
  const [bulkDeleteOpen, setBulkDeleteOpen] = React.useState(false);
  const [renameTarget, setRenameTarget] = React.useState<{ level: 'processGroup' | 'progressId'; id: string; currentLabel: string } | undefined>(undefined);
  const [renameValue, setRenameValue] = React.useState('');
  // Which lock dialog is open, if any, and the reason text being typed
  // into it - 'lock' and 'unlock' share one dialog/one reason field since
  // they're never open at the same time.
  const [lockDialogMode, setLockDialogMode] = React.useState<'lock' | 'unlock' | undefined>(undefined);
  const [lockReasonValue, setLockReasonValue] = React.useState('');
  // "Leave a comment" dialog - never gated by activeLock (see
  // ISwimlaneComment), so it's a plain independent boolean rather than
  // sharing lockDialogMode's 'lock' | 'unlock' pattern.
  const [commentDialogOpen, setCommentDialogOpen] = React.useState(false);
  const [commentValue, setCommentValue] = React.useState('');
  const [lastAction, setLastAction] = React.useState<UndoAction | undefined>(undefined);
  const [activeTab, setActiveTab] = React.useState<MainTab>('flows');

  // The undo banner doesn't linger forever - clears itself a while after
  // the action it's offering to undo, same as most "Undo" toasts. Resets
  // on every new action (the effect re-runs since lastAction is a new
  // object reference each time), so a fresh action always gets the full
  // window rather than inheriting whatever was left of a previous one.
  React.useEffect(() => {
    if (!lastAction) return;
    const timer = window.setTimeout(() => setLastAction(undefined), 10000);
    return () => window.clearTimeout(timer);
  }, [lastAction]);

  // Promise.allSettled, not Promise.all - the three sources are genuinely
  // independent (separate SharePoint lists, each with its own chance of
  // being misconfigured or not existing yet), so one failing (e.g. the
  // Employees list title not confirmed yet) shouldn't block the other two
  // from showing. A real case this fixed: Process steps resolving fine
  // while Employees still 404s would previously blank out the whole app
  // instead of just the Employees-dependent pieces.
  const loadAll = React.useCallback(() => {
    setLoading(true);
    setError(undefined);
    Promise.allSettled([
      dataService.getProcessSteps(), dataService.getEmployees(), dataService.getRiskStatements(),
      dataService.getProcessGroupLabels(), dataService.getProgressIdLabels(), dataService.getProgressIdLocks(),
      dataService.getSwimlaneComments()
    ])
      .then(([stepsResult, employeesResult, risksResult, groupLabelsResult, progressIdLabelsResult, locksResult, commentsResult]) => {
        const errors: string[] = [];

        if (stepsResult.status === 'fulfilled') {
          setSteps(stepsResult.value);
        } else {
          errors.push(`Process steps: ${describeError(stepsResult.reason)}`);
        }

        if (employeesResult.status === 'fulfilled') {
          setEmployees(employeesResult.value);
        } else {
          errors.push(`Employees: ${describeError(employeesResult.reason)}`);
        }

        if (risksResult.status === 'fulfilled') {
          setRiskStatements(risksResult.value);
        } else {
          errors.push(`Risk statements: ${describeError(risksResult.reason)}`);
        }

        // Not surfaced as an error - this list is a nice-to-have that most
        // sites won't have created yet, and every group it would name
        // already has a working fallback (the static table, or just
        // "Process Group X.Y"), so a missing/misconfigured list here
        // shouldn't read as something broken the way the other three do.
        if (groupLabelsResult.status === 'fulfilled') {
          setProcessGroupLabels(groupLabelsResult.value);
        }

        if (progressIdLabelsResult.status === 'fulfilled') {
          setProgressIdLabels(progressIdLabelsResult.value);
        }

        // Same "nice to have, not a blocking error" treatment as the label
        // lists above, for the same reason (this list won't exist on most
        // sites yet). Deliberate trade-off: a failure here means locks are
        // treated as "nothing is locked" rather than making the whole app
        // unusable - reasonable given locking is a v1, anyone-signed-in
        // social control (see IProgressIdLock), not a hard permission
        // system, but worth knowing this is fail-open, not fail-closed.
        if (locksResult.status === 'fulfilled') {
          setProgressIdLocks(locksResult.value);
        }

        // Same fail-open treatment as locks/labels above - a missing
        // "Swimlane comments" list means comments are treated as "none
        // posted yet" rather than blocking the app.
        if (commentsResult.status === 'fulfilled') {
          setSwimlaneComments(commentsResult.value);
        }

        setError(errors.length > 0 ? errors.join(' | ') : undefined);
        setLoading(false);
      });
  }, [dataService]);

  React.useEffect(() => { loadAll(); }, [loadAll]);

  // User-added names for Process Groups the static apqcHierarchy.ts table
  // doesn't already cover (see IProcessGroupLabel) - merged in wherever a
  // Process Group name is looked up or listed, same as the static table.
  const customGroupNames = React.useMemo(
    () => Object.fromEntries(processGroupLabels.map(l => [l.groupId, l.name])) as Record<string, string>,
    [processGroupLabels]
  );

  // Same idea as customGroupNames, one level down - a name for a Progress
  // ID that has no real steps yet to derive one from (see getLabel on the
  // Progress ID HierarchyPicker below, which otherwise falls back to a
  // real step's processDescription or the static APQC table).
  const customProgressIdNames = React.useMemo(
    () => Object.fromEntries(progressIdLabels.map(l => [l.progressId, l.name])) as Record<string, string>,
    [progressIdLabels]
  );

  const stepsInCategory = React.useMemo(
    () => selectedCategoryId ? steps.filter(s => getCategoryId(s.processStepId) === selectedCategoryId) : [],
    [steps, selectedCategoryId]
  );

  const stepsInProcessGroup = React.useMemo(
    () => selectedProcessGroupId ? stepsInCategory.filter(s => getProcessGroupId(s.processStepId) === selectedProcessGroupId) : [],
    [stepsInCategory, selectedProcessGroupId]
  );

  const stepsInProgressId = React.useMemo(
    () => selectedProgressId ? steps.filter(s => getProgressId(s.processStepId) === selectedProgressId) : [],
    [steps, selectedProgressId]
  );

  // The readable name for wherever the header breadcrumb is currently
  // pointing - the breadcrumb itself stays the compact numeric trail
  // ("9 / 9.2 / 9.2.3") since that's genuinely useful for quick reference,
  // but showing ONLY numbers up there left no way to tell at a glance
  // which actual swimlane is open without drilling back down through the
  // pickers - a real user flagged this directly. Same name-resolution
  // logic each picker level already uses - a custom label (an explicit,
  // deliberate rename) always wins over a step's own processDescription,
  // not the other way around, so a wrong/junk value that ended up in real
  // data (e.g. someone typing "yes" into the wrong field) can always be
  // corrected via rename instead of being permanently stuck showing.
  const currentLevelName = React.useMemo(() => {
    if (drilledDownStepId) {
      return stepsInProgressId.find(s => s.processStepId === drilledDownStepId)?.processStepName;
    }
    if (selectedProgressId) {
      return customProgressIdNames[selectedProgressId] || stepsInProgressId[0]?.processDescription || getProgressIdName(selectedProgressId);
    }
    if (selectedProcessGroupId) {
      return customGroupNames[selectedProcessGroupId] || getProcessGroupName(selectedProcessGroupId);
    }
    if (selectedCategoryId) {
      return getCategoryName(selectedCategoryId);
    }
    return undefined;
  }, [drilledDownStepId, selectedProgressId, selectedProcessGroupId, selectedCategoryId, stepsInProgressId, customGroupNames, customProgressIdNames]);

  // The lock currently in effect for the swimlane actually on screen right
  // now (this Progress ID + this region), if any - undefined means it's
  // editable. Scoped to progressId+region together, not just progressId,
  // matching the confirmed design rule that regions are genuinely separate
  // swimlanes (see IProgressIdLock/FlowRegionTabs) - locking the UK
  // version of a flow never touches the US/SA versions on the same
  // Progress ID.
  const activeLock = React.useMemo(
    () => selectedProgressId ? findActiveLock(progressIdLocks, selectedProgressId, selectedFlowRegion) : undefined,
    [progressIdLocks, selectedProgressId, selectedFlowRegion]
  );

  // Same progressId + region scoping as activeLock above - only the
  // comments that actually belong to the swimlane currently on screen,
  // not every comment ever left anywhere (that's what the separate
  // Improvements tab is for).
  const commentsForSwimlane = React.useMemo(
    () => selectedProgressId
      ? swimlaneComments.filter(c => c.progressId === selectedProgressId && c.region === (selectedFlowRegion || ''))
      : [],
    [swimlaneComments, selectedProgressId, selectedFlowRegion]
  );

  // A Progress ID can hold several genuinely separate swimlanes side by
  // side, one per region (see FlowRegionTabs) - narrowed here, upstream of
  // everything else derived from stepsInProgressId, so picking a region
  // acts as the primary partition and Process Step ID tabs/Depends-on
  // options/the canvas itself only ever see that region's own steps.
  const stepsInRegion = React.useMemo(
    () => selectedFlowRegion ? stepsInProgressId.filter(s => (s.region || '') === selectedFlowRegion) : stepsInProgressId,
    [stepsInProgressId, selectedFlowRegion]
  );

  // Scoped to the current swimlane (this region's own steps within the
  // Progress ID), not the full cross-progress-ID dataset - a step
  // realistically only ever depends on something in its own flow, and
  // listing all ~40 steps from every unrelated flow made the real option
  // buried in noise. No excludeStepId - a brand-new step has no "self" to
  // leave out, unlike the edit panel's version of this same list.
  const addStepDependsOnOptions = React.useMemo(() => buildDependsOnOptions(stepsInRegion), [stepsInRegion]);

  const visibleSteps = React.useMemo(
    () => drilledDownStepId ? stepsInRegion.filter(s => s.processStepId === drilledDownStepId) : stepsInRegion,
    [stepsInRegion, drilledDownStepId]
  );

  // Resolved against the FULL, unfiltered `steps` array - row-number-based
  // DependsOn tokens only make sense against original load order, not a
  // filtered/reordered subset. SwimlaneCanvas takes care of only drawing
  // the edges whose endpoints are currently visible.
  const edges = React.useMemo(() => resolveDependencyEdges(steps), [steps]);

  const handleLabelEdge = (toRowId: string, token: string, label: string): void => {
    const step = steps.find(s => s.id === toRowId);
    if (!step) return;
    const updated: IProcessStep = {
      ...step,
      edgeLabels: { ...(step.edgeLabels || {}), [token]: label }
    };
    setSteps(prev => prev.map(s => (s.id === toRowId ? updated : s)));
    dataService.updateProcessStep(updated).catch((err: Error) => setError(err.message));
  };

  // updateProcessStep returns void (unlike addProcessStep, which hands
  // back a freshly-stamped object) - GraphDataService can't know the true
  // native lastModifiedBy/lastModifiedDateTime until the next reload
  // re-fetches it, so the caller has to stamp its own optimistic
  // modifiedBy/modifiedAt here instead. Never touches createdBy/createdAt
  // - only who/when CREATED a step changes that, not an edit.
  const withModifiedStamp = (step: IProcessStep): IProcessStep => ({
    ...step,
    modifiedBy: currentUserName,
    modifiedAt: new Date().toISOString()
  });

  const handleEditStep = (updated: IProcessStep): void => {
    if (activeLock) return; // defense in depth - SwimlaneCanvas's isLocked prop already keeps its Save button from calling this
    const previous = steps.find(s => s.id === updated.id);
    const stamped = withModifiedStamp(updated);
    setSteps(prev => prev.map(s => (s.id === stamped.id ? stamped : s)));
    dataService.updateProcessStep(stamped).catch((err: Error) => setError(err.message));
    if (previous) setLastAction({ type: 'edit', previous });
  };

  const handleDeleteStep = (stepId: string): void => {
    if (activeLock) return; // defense in depth - SwimlaneCanvas's isLocked prop already hides the delete affordance
    const step = steps.find(s => s.id === stepId);
    setSteps(prev => prev.filter(s => s.id !== stepId));
    dataService.deleteProcessStep(stepId).catch((err: Error) => setError(err.message));
    if (step) setLastAction({ type: 'delete', step });
  };

  // Deletes every step currently visible - the whole selected Process Step
  // ID group's flow, or the whole Progress ID if "All" is selected (see
  // visibleSteps) - so removing a whole mistaken flow doesn't mean
  // deleting each of its steps one at a time via the edit panel.
  const handleBulkDelete = (): void => {
    if (activeLock) return; // defense in depth - the menu item that opens this is already disabled while locked
    const deleted = visibleSteps.slice();
    const idsToDelete = deleted.map(s => s.id);
    setSteps(prev => prev.filter(s => !idsToDelete.includes(s.id)));
    idsToDelete.forEach(id => {
      dataService.deleteProcessStep(id).catch((err: Error) => setError(err.message));
    });
    setLastAction({ type: 'bulkDelete', steps: deleted });
    setBulkDeleteOpen(false);
  };

  const handleUndo = (): void => {
    if (!lastAction) return;
    if (lastAction.type === 'edit') {
      // Undoing IS itself a real modification event (someone just acted,
      // right now) even though the CONTENT reverts to old values - so this
      // gets a fresh modifiedBy/modifiedAt too, same as any other edit,
      // rather than silently reverting the audit trail along with the
      // content.
      const stamped = withModifiedStamp(lastAction.previous);
      setSteps(prev => prev.map(s => (s.id === stamped.id ? stamped : s)));
      dataService.updateProcessStep(stamped).catch((err: Error) => setError(err.message));
    } else if (lastAction.type === 'delete') {
      const { id: _id, ...rest } = lastAction.step;
      dataService.addProcessStep(rest)
        .then(created => setSteps(prev => [...prev, created]))
        .catch((err: Error) => setError(err.message));
    } else {
      lastAction.steps.forEach(step => {
        const { id: _id, ...rest } = step;
        dataService.addProcessStep(rest)
          .then(created => setSteps(prev => [...prev, created]))
          .catch((err: Error) => setError(err.message));
      });
    }
    setLastAction(undefined);
  };

  const handleImported = (created: IProcessStep[]): void => {
    setSteps(prev => [...prev, ...created]);
  };

  const handleGroupLabelCreated = (created: IProcessGroupLabel): void => {
    setProcessGroupLabels(prev => [...prev, created]);
  };

  // Lands the user straight in the empty Process Group or Progress ID
  // they just named, same "go straight to what you made" treatment as
  // handleProcessCreated gets for a full step - the difference here is
  // there's no step to select underneath it, so drilling in shows an
  // empty picker/canvas ready for "Add a step".
  const handleShellCreated = (created: IProcessGroupLabel | IProgressIdLabel): void => {
    if ('groupId' in created) {
      setProcessGroupLabels(prev => [...prev, created]);
      setSelectedProcessGroupId(created.groupId);
    } else {
      setProgressIdLabels(prev => [...prev, created]);
      setSelectedProgressId(created.progressId);
      setSelectedFlowRegion(undefined);
    }
    setAddShellState(undefined);
  };

  const openRename = (level: 'processGroup' | 'progressId', id: string, currentLabel: string): void => {
    setRenameTarget({ level, id, currentLabel });
    setRenameValue(currentLabel);
  };

  // Renaming something that already has a custom label updates that same
  // record; renaming one that's still showing its static apqcHierarchy.ts
  // name (or a step-derived/generic fallback) creates a new custom label
  // instead, which then wins over that fallback the same way it already
  // does for a brand-new group/progress ID (see customGroupNames /
  // customProgressIdNames) - this is the ONLY way to correct a wrong
  // value that ended up baked into real step data (e.g. a real case: a
  // Progress ID showing "yes" as its name because that's literally what
  // ended up in some step's Process Description field).
  const handleRenameSave = (): void => {
    if (!renameTarget) return;
    const trimmed = renameValue.trim();
    if (!trimmed) return;
    if (renameTarget.level === 'processGroup') {
      const existingLabel = processGroupLabels.find(l => l.groupId === renameTarget.id);
      if (existingLabel) {
        setProcessGroupLabels(prev => prev.map(l => (l.id === existingLabel.id ? { ...l, name: trimmed } : l)));
        dataService.updateProcessGroupLabel(existingLabel.id, trimmed).catch((err: Error) => setError(err.message));
      } else {
        dataService.addProcessGroupLabel(renameTarget.id, trimmed)
          .then(created => setProcessGroupLabels(prev => [...prev, created]))
          .catch((err: Error) => setError(err.message));
      }
    } else {
      const existingLabel = progressIdLabels.find(l => l.progressId === renameTarget.id);
      if (existingLabel) {
        setProgressIdLabels(prev => prev.map(l => (l.id === existingLabel.id ? { ...l, name: trimmed } : l)));
        dataService.updateProgressIdLabel(existingLabel.id, trimmed).catch((err: Error) => setError(err.message));
      } else {
        dataService.addProgressIdLabel(renameTarget.id, trimmed)
          .then(created => setProgressIdLabels(prev => [...prev, created]))
          .catch((err: Error) => setError(err.message));
      }
    }
    setRenameTarget(undefined);
  };

  // Confirmed design rule: v1 doesn't restrict who can lock/unlock to
  // specific people - anyone signed in can do either. Accountability comes
  // from every action being attributed (currentUserName) and permanently
  // logged (see IProgressIdLock's append-only shape), not from a
  // technical permission barrier.
  const handleLockConfirm = (): void => {
    if (!selectedProgressId || lockDialogMode !== 'lock') return;
    const region = selectedFlowRegion || '';
    dataService.lockProgressId(selectedProgressId, region, currentUserName, lockReasonValue.trim())
      .then(created => setProgressIdLocks(prev => [...prev, created]))
      .catch((err: Error) => setError(err.message));
    setLockDialogMode(undefined);
    setLockReasonValue('');
  };

  const handleUnlockConfirm = (): void => {
    if (!activeLock || lockDialogMode !== 'unlock') return;
    const reason = lockReasonValue.trim();
    const unlockedAt = new Date().toISOString();
    setProgressIdLocks(prev => prev.map(l => (l.id === activeLock.id
      ? { ...l, unlockedBy: currentUserName, unlockedAt, unlockReason: reason }
      : l)));
    dataService.unlockProgressId(activeLock.id, currentUserName, reason).catch((err: Error) => setError(err.message));
    setLockDialogMode(undefined);
    setLockReasonValue('');
  };

  // Never gated by activeLock - see ISwimlaneComment for why leaving a
  // comment is deliberately independent of the lock/edit flow entirely.
  const handleCommentConfirm = (): void => {
    const text = commentValue.trim();
    if (!selectedProgressId || !text) return;
    const region = selectedFlowRegion || '';
    dataService.addSwimlaneComment(selectedProgressId, region, currentUserName, text)
      .then(created => setSwimlaneComments(prev => [...prev, created]))
      .catch((err: Error) => setError(err.message));
    setCommentDialogOpen(false);
    setCommentValue('');
  };

  // Lands the user straight in the swimlane they just created, the same
  // place they'd be if they'd clicked all the way down through an
  // already-populated area - drilling back through empty pickers to find
  // what was just added would be a pointless extra step.
  const handleProcessCreated = (created: IProcessStep): void => {
    setSteps(prev => [...prev, created]);
    setSelectedCategoryId(getCategoryId(created.processStepId));
    setSelectedProcessGroupId(getProcessGroupId(created.processStepId));
    setSelectedProgressId(getProgressId(created.processStepId));
    // NewProcessModal has no region context to inherit (it's often used to
    // start an entirely new area from scratch), so the created step is
    // unregioned - "All" is the only view it's guaranteed to show up in.
    setSelectedFlowRegion(undefined);
    // Selects the new step's own Process Step ID tab, not "All" - without
    // this, "Add a step" right afterward defaulted to the bare Progress ID
    // as its processStepId (drilledDownStepId || selectedProgressId, with
    // drilledDownStepId unset), landing new steps in a second, separate
    // column group instead of continuing the one the user just started.
    setDrilledDownStepId(created.processStepId);
    setNewProcessOpen(false);
  };

  // Lands the user straight in the section they just created, same "go
  // straight to what you made" treatment as handleProcessCreated - the
  // new section already belongs to the current Progress ID/region (see
  // AddStepSectionModal), so only its own Process Step ID tab needs
  // selecting, nothing else about the current context changes.
  const handleSectionCreated = (created: IProcessStep): void => {
    setSteps(prev => [...prev, created]);
    setDrilledDownStepId(created.processStepId);
    setAddSectionOpen(false);
  };

  // Jumps straight to a step found via GlobalSearch - lands on the
  // Progress ID's swimlane (Category -> Process Group -> Progress ID),
  // same as clicking all the way down by hand, but stops there rather
  // than also auto-selecting that step's own Process Step ID tab - the
  // swimlane for the whole Progress ID is the useful landing spot, since
  // it shows the found step in context next to everything around it.
  const handleSearchNavigate = (step: IProcessStep): void => {
    setActiveTab('flows');
    setSelectedCategoryId(getCategoryId(step.processStepId));
    setSelectedProcessGroupId(getProcessGroupId(step.processStepId));
    setSelectedProgressId(getProgressId(step.processStepId));
    // Guarantees the found step is actually visible - without this, a
    // stale region selection from wherever the user was browsing before
    // could hide the very step search just landed them on.
    setSelectedFlowRegion(step.region || undefined);
    setDrilledDownStepId(undefined);
  };

  const newProcessPrefix = selectedProcessGroupId ? `${selectedProcessGroupId}.` : selectedCategoryId ? `${selectedCategoryId}.` : '';

  const handleAddStep = (): void => {
    if (!selectedProgressId || !newStepDraft.actionDescription.trim() || activeLock) return;
    // The very first step in a brand-new region has no reference step IN
    // THAT REGION to inherit from (stepsInRegion is empty) - it used to
    // fall back straight to blank/bare defaults there, leaving Process
    // Description empty and APQC Title as just the bare Progress ID
    // number. Falls back to ANY step in the Progress ID instead: Process
    // Description/Step Name describe the same underlying business process
    // regardless of which region's specific procedure this is, so
    // inheriting them from another region's step is far more useful than
    // leaving them blank. APQC Title alone gets rebuilt for the new
    // region (mirroring the real data's own "9.6.1 - UK" convention)
    // rather than inherited verbatim, since copying e.g. "9.6.1 - US"
    // onto a brand-new UK step would mislabel it.
    const regionReferenceStep = stepsInRegion[0];
    const anyReferenceStep = stepsInProgressId[0];
    const { dependsOnStepIds, ...fields } = newStepDraft;
    setSaving(true);
    dataService.addProcessStep({
      apqcTitle: regionReferenceStep
        ? regionReferenceStep.apqcTitle
        : selectedFlowRegion
          ? `${selectedProgressId} - ${selectedFlowRegion}`
          : anyReferenceStep ? anyReferenceStep.apqcTitle : selectedProgressId,
      // A custom rename (if one's been set) wins over whatever's actually
      // sitting in existing steps' processDescription - otherwise a new
      // step would keep perpetuating a wrong value a rename was supposed
      // to have already corrected.
      processDescription: (selectedProgressId && customProgressIdNames[selectedProgressId])
        || regionReferenceStep?.processDescription || anyReferenceStep?.processDescription || '',
      processStepId: drilledDownStepId || selectedProgressId,
      processStepName: regionReferenceStep?.processStepName || anyReferenceStep?.processStepName || '',
      region: selectedFlowRegion || '',
      ...fields,
      actionDescription: fields.actionDescription.trim(),
      dependsOn: stepIdsToDependsOnTokens(steps, dependsOnStepIds)
    })
      .then(created => {
        setSteps(prev => [...prev, created]);
        setNewStepDraft(emptyStepDraft());
        setSaving(false);
      })
      .catch((err: Error) => {
        setError(err.message);
        setSaving(false);
      });
  };

  // Drag-from-legend creation (see ShapeLegend's draggable swatches and
  // SwimlaneCanvas's onCreateStep) - unlike handleAddStep above, there's
  // no form to fill in first. The step is created immediately, positioned
  // right where it was dropped (columnStep's Process Step ID/group, the
  // lane it landed in, taking over columnStep's own column and shifting
  // it one place right - see computeInsertOrderBefore), with a
  // placeholder description, and its edit panel opens automatically right
  // after (see autoOpenStepId) so the real details get filled in on the
  // spot - the same "drop a shape,
  // then type into it" flow a real diagramming tool would give you.
  // apqcTitle/processDescription/processStepName/region are inherited
  // straight from columnStep, the concrete step actually sitting at the
  // dropped position, rather than handleAddStep's own reference-step
  // fallback chain (unnecessary here - there's always a real columnStep,
  // that's what was dropped onto).
  const handleCreateStepFromShape = (shapeOverride: string, lane: string, columnStep: IProcessStep): void => {
    if (activeLock) return;
    dataService.addProcessStep({
      apqcTitle: columnStep.apqcTitle,
      processDescription: columnStep.processDescription,
      processStepId: columnStep.processStepId,
      processStepName: columnStep.processStepName,
      region: columnStep.region || '',
      action: '',
      actionType: '',
      actionDescription: 'New step',
      shapeOverride,
      responsibleJobTitle: lane === 'Unassigned' ? '' : lane,
      manualOrder: computeInsertOrderBefore(steps, columnStep.id),
      dependsOn: [],
      linkedRisks: []
    })
      .then(created => {
        setSteps(prev => [...prev, created]);
        setAutoOpenStepId(created.id);
      })
      .catch((err: Error) => setError(err.message));
  };

  if (loading) {
    // Keeps the real header (logo, title) visible instead of a bare
    // Spinner with no layout at all - on the real Graph data path this can
    // sit on screen for a couple of seconds, long enough that an unstyled
    // corner spinner reads as the app being broken rather than loading.
    return (
      <div className={styles.page}>
        <header className={styles.appBar}>
          <img src={qleLogo} className={styles.appBarLogo} alt="Quantum Leap Energy" />
          <div>
            <h2 className={styles.title}>Swimlane Studio</h2>
            <p className={styles.breadcrumb}>Business process visualisation</p>
          </div>
        </header>
        <div className={styles.loadingState}>
          <Spinner label="Loading process data..." />
        </div>
      </div>
    );
  }

  return (
    <div className={styles.page}>
      <header className={styles.appBar}>
        <img src={qleLogo} className={styles.appBarLogo} alt="Quantum Leap Energy" />
        <div>
          <h2 className={styles.title}>Swimlane Studio</h2>
          {selectedCategoryId ? (
            <p className={styles.breadcrumb}>
              {[selectedCategoryId, selectedProcessGroupId, selectedProgressId, drilledDownStepId].filter(Boolean).join(' / ')}
              {currentLevelName ? ` — ${currentLevelName}` : ''}
            </p>
          ) : (
            <p className={styles.breadcrumb}>Business process visualisation</p>
          )}
        </div>
        <div className={styles.appBarSearch}>
          <GlobalSearch steps={steps} onNavigate={handleSearchNavigate} />
        </div>
        <div className={styles.appBarActions}>
          {onSignOut && (
            <DefaultButton
              text={signOutLabel || 'Sign out'}
              onClick={onSignOut}
              styles={{
                root: {
                  background: 'rgba(255,255,255,0.08)',
                  border: '1px solid rgba(255,255,255,0.3)',
                  borderRadius: 8,
                  marginRight: 10
                },
                rootHovered: { background: 'rgba(255,255,255,0.18)', border: '1px solid rgba(255,255,255,0.45)' },
                rootPressed: { background: 'rgba(255,255,255,0.24)' },
                label: { color: '#fff', fontWeight: 600 }
              }}
            />
          )}
        </div>
      </header>

      <ImportCsvModal
        isOpen={importOpen}
        dataService={dataService}
        insertionIndex={steps.length}
        onDismiss={() => setImportOpen(false)}
        onImported={handleImported}
      />

      <NewProcessModal
        isOpen={newProcessOpen}
        steps={steps}
        employees={employees}
        riskStatements={riskStatements}
        dataService={dataService}
        processStepIdPrefix={newProcessPrefix}
        knownProcessGroupIds={new Set([...Object.keys(APQC_PROCESS_GROUP_NAMES), ...Object.keys(customGroupNames)])}
        onDismiss={() => setNewProcessOpen(false)}
        onCreated={handleProcessCreated}
        onGroupLabelCreated={handleGroupLabelCreated}
      />

      <AddHierarchyShellModal
        isOpen={!!addShellState}
        level={addShellState?.level || 'processGroup'}
        idPrefix={addShellState?.idPrefix || ''}
        dataService={dataService}
        onDismiss={() => setAddShellState(undefined)}
        onCreated={handleShellCreated}
      />

      <AddStepSectionModal
        isOpen={addSectionOpen}
        idPrefix={selectedProgressId ? `${selectedProgressId}.` : ''}
        referenceStep={stepsInProgressId[0]}
        region={selectedFlowRegion}
        dataService={dataService}
        onDismiss={() => setAddSectionOpen(false)}
        onCreated={handleSectionCreated}
      />

      <Dialog
        hidden={!bulkDeleteOpen}
        onDismiss={() => setBulkDeleteOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete this flow?',
          subText: `This removes all ${visibleSteps.length} step${visibleSteps.length === 1 ? '' : 's'} currently shown - not just one. You'll have a short window to undo it afterward.`
        }}
      >
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => setBulkDeleteOpen(false)} />
          <PrimaryButton
            text={`Delete ${visibleSteps.length} step${visibleSteps.length === 1 ? '' : 's'}`}
            onClick={handleBulkDelete}
            styles={{ root: { background: 'var(--risk-high)', border: 'none' }, rootHovered: { background: '#b02419' } }}
          />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!renameTarget}
        onDismiss={() => setRenameTarget(undefined)}
        dialogContentProps={{ type: DialogType.normal, title: `Rename ${renameTarget?.id || ''}` }}
      >
        <TextField
          label={renameTarget?.level === 'progressId' ? 'Progress ID name' : 'Process Group name'}
          value={renameValue}
          onChange={(_e, v) => setRenameValue(v || '')}
          onKeyDown={e => { if (e.key === 'Enter') handleRenameSave(); }}
        />
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => setRenameTarget(undefined)} />
          <PrimaryButton text="Save" onClick={handleRenameSave} disabled={!renameValue.trim()} />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!lockDialogMode}
        onDismiss={() => { setLockDialogMode(undefined); setLockReasonValue(''); }}
        dialogContentProps={{
          type: DialogType.normal,
          title: lockDialogMode === 'unlock' ? 'Unlock this swimlane?' : 'Lock this swimlane?',
          subText: lockDialogMode === 'unlock'
            ? 'Reopens it for editing. This is recorded permanently, same as the original lock.'
            : "Blocks further edits until it's unlocked again. Who locked it, when, and why is recorded permanently."
        }}
      >
        <TextField
          label="Reason (optional)"
          placeholder={lockDialogMode === 'unlock' ? 'e.g. Reopening to fix an error found in review' : 'e.g. Approved for FY26 audit'}
          value={lockReasonValue}
          onChange={(_e, v) => setLockReasonValue(v || '')}
          onKeyDown={e => { if (e.key === 'Enter') (lockDialogMode === 'unlock' ? handleUnlockConfirm() : handleLockConfirm()); }}
        />
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => { setLockDialogMode(undefined); setLockReasonValue(''); }} />
          <PrimaryButton
            text={lockDialogMode === 'unlock' ? 'Unlock' : 'Lock'}
            onClick={lockDialogMode === 'unlock' ? handleUnlockConfirm : handleLockConfirm}
          />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!commentDialogOpen}
        onDismiss={() => { setCommentDialogOpen(false); setCommentValue(''); }}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Leave a comment',
          subText: "Posted under your name, permanently - visible here and on the Improvements tab. Not blocked by a lock, and doesn't change anything on the swimlane itself."
        }}
      >
        <TextField
          label="Comment"
          multiline
          rows={4}
          placeholder="e.g. This step should route to the Regional Finance Manager instead"
          value={commentValue}
          onChange={(_e, v) => setCommentValue(v || '')}
        />
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => { setCommentDialogOpen(false); setCommentValue(''); }} />
          <PrimaryButton text="Post comment" onClick={handleCommentConfirm} disabled={!commentValue.trim()} />
        </DialogFooter>
      </Dialog>

      <section className={styles.swimlaneStudio}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
            {error}
          </MessageBar>
        )}

        {lastAction && (
          <MessageBar
            messageBarType={MessageBarType.info}
            onDismiss={() => setLastAction(undefined)}
            actions={<MessageBarButton onClick={handleUndo}>Undo</MessageBarButton>}
          >
            {lastAction.type === 'delete' && `Deleted "${lastAction.step.actionDescription}".`}
            {lastAction.type === 'bulkDelete' && `Deleted ${lastAction.steps.length} step${lastAction.steps.length === 1 ? '' : 's'}.`}
            {lastAction.type === 'edit' && `Updated "${lastAction.previous.actionDescription}".`}
          </MessageBar>
        )}

        <Pivot
          className={styles.mainTabs}
          selectedKey={activeTab}
          onLinkClick={(item?: PivotItem) => {
            const key = item?.props.itemKey;
            setActiveTab(key === 'employees' || key === 'risks' || key === 'audit' || key === 'improvements' ? key : 'flows');
          }}
        >
          <PivotItem headerText="Process Flows" itemKey="flows" />
          <PivotItem headerText="Employees" itemKey="employees" />
          <PivotItem headerText="Risk Register" itemKey="risks" />
          <PivotItem headerText="Audit" itemKey="audit" />
          <PivotItem headerText="Improvements" itemKey="improvements" />
        </Pivot>

        {activeTab === 'employees' ? (
          <EmployeesList employees={employees} />
        ) : activeTab === 'risks' ? (
          <RiskRegisterList riskStatements={riskStatements} steps={steps} />
        ) : activeTab === 'audit' ? (
          <AuditView steps={steps} progressIdLocks={progressIdLocks} />
        ) : activeTab === 'improvements' ? (
          <ImprovementsView comments={swimlaneComments} />
        ) : (
          <>
            {!selectedCategoryId && (
              <div className={styles.intro}>
                <h3>Select a category to explore</h3>
                <p>Pick an APQC process category, then a process group, then a Progress ID to open its swimlane.</p>
              </div>
            )}

            {!selectedCategoryId ? (
              <>
                <div className={styles.toolbar}>
                  <DefaultButton text="Import CSV" iconProps={{ iconName: 'Upload' }} onClick={() => setImportOpen(true)} />
                </div>
                <HierarchyPicker
                  steps={steps}
                  getGroupId={getCategoryId}
                  getLabel={id => getCategoryName(id)}
                  onSelect={setSelectedCategoryId}
                  allGroupIds={Object.keys(APQC_CATEGORY_NAMES)}
                  countBy={getProcessGroupId}
                  countLabel="process group"
                  allSubGroupIds={[...Object.keys(APQC_PROCESS_GROUP_NAMES), ...Object.keys(customGroupNames)]}
                />
              </>
            ) : !selectedProcessGroupId ? (
              <>
                <div className={styles.toolbar}>
                  <DefaultButton text="Back to Categories" onClick={() => setSelectedCategoryId(undefined)} />
                  <DefaultButton text="Import CSV" iconProps={{ iconName: 'Upload' }} onClick={() => setImportOpen(true)} />
                </div>
                <HierarchyPicker
                  steps={stepsInCategory}
                  getGroupId={getProcessGroupId}
                  getLabel={id => customGroupNames[id] || getProcessGroupName(id)}
                  onSelect={setSelectedProcessGroupId}
                  allGroupIds={[...Object.keys(APQC_PROCESS_GROUP_NAMES), ...Object.keys(customGroupNames)].filter(id => getCategoryId(id) === selectedCategoryId)}
                  onAddNew={() => setAddShellState({ level: 'processGroup', idPrefix: `${selectedCategoryId}.` })}
                  addNewLabel="+ Add new process group"
                  onRename={(id, label) => openRename('processGroup', id, label)}
                  countBy={getProgressId}
                  countLabel="process"
                  countLabelPlural="processes"
                  allSubGroupIds={[...Object.keys(APQC_PROGRESS_ID_NAMES), ...Object.keys(customProgressIdNames)].filter(id => getCategoryId(id) === selectedCategoryId)}
                />
              </>
            ) : !selectedProgressId ? (
              <>
                <div className={styles.toolbar}>
                  <DefaultButton text="Back to Process Groups" onClick={() => setSelectedProcessGroupId(undefined)} />
                  <DefaultButton text="Import CSV" iconProps={{ iconName: 'Upload' }} onClick={() => setImportOpen(true)} />
                </div>
                <HierarchyPicker
                  steps={stepsInProcessGroup}
                  getGroupId={getProgressId}
                  getLabel={(id, sampleStep) => customProgressIdNames[id] || sampleStep?.processDescription || getProgressIdName(id)}
                  onSelect={setSelectedProgressId}
                  allGroupIds={[...Object.keys(APQC_PROGRESS_ID_NAMES), ...Object.keys(customProgressIdNames)].filter(id => getProcessGroupId(id) === selectedProcessGroupId)}
                  onAddNew={() => setAddShellState({ level: 'progressId', idPrefix: `${selectedProcessGroupId}.` })}
                  addNewLabel="+ Add new progress ID"
                  onRename={(id, label) => openRename('progressId', id, label)}
                />
              </>
            ) : (
              <>
                <div className={styles.toolbar}>
                  {activeLock && (
                    <MessageBar messageBarType={MessageBarType.warning} className={styles.lockBanner}>
                      Locked by {activeLock.lockedBy} on {new Date(activeLock.lockedAt).toLocaleString()}
                      {activeLock.reason ? ` — ${activeLock.reason}` : ''}. Read-only until it's unlocked.
                    </MessageBar>
                  )}
                  <div className={styles.toolbarRow}>
                    <DefaultButton text="Back to Progress IDs" onClick={() => { setSelectedProgressId(undefined); setDrilledDownStepId(undefined); setSelectedFlowRegion(undefined); }} />
                    {activeLock ? (
                      <DefaultButton
                        text="Unlock swimlane"
                        iconProps={{ iconName: 'Unlock' }}
                        onClick={() => setLockDialogMode('unlock')}
                      />
                    ) : (
                      <DefaultButton
                        text="Lock swimlane"
                        iconProps={{ iconName: 'Lock' }}
                        disabled={visibleSteps.length === 0}
                        onClick={() => setLockDialogMode('lock')}
                      />
                    )}
                    <DefaultButton
                      text="Leave a comment"
                      iconProps={{ iconName: 'Comment' }}
                      onClick={() => setCommentDialogOpen(true)}
                    />
                    <IconButton
                      menuIconProps={{ iconName: 'More' }}
                      title="More actions"
                      ariaLabel="More actions"
                      styles={{ root: { border: '1px solid var(--border)', borderRadius: 4 } }}
                      menuProps={{
                        items: [
                          { key: 'addNew', text: '+ Add new process', iconProps: { iconName: 'Add' }, onClick: () => { setNewProcessOpen(true); }, disabled: !!activeLock },
                          { key: 'importCsv', text: 'Import CSV', iconProps: { iconName: 'Upload' }, onClick: () => { setImportOpen(true); }, disabled: !!activeLock },
                          {
                            key: 'deleteFlow',
                            text: `Delete flow (${visibleSteps.length})`,
                            iconProps: { iconName: 'Delete', styles: { root: { color: 'var(--risk-high)' } } },
                            disabled: visibleSteps.length === 0 || !!activeLock,
                            onClick: () => { setBulkDeleteOpen(true); }
                          }
                        ]
                      }}
                    />
                  </div>
                  <FlowRegionTabs
                    steps={stepsInProgressId}
                    selectedRegion={selectedFlowRegion}
                    onSelect={region => {
                      // A Process Step ID tab selected in one region may not
                      // exist (or mean the same thing) in another - clear it
                      // so switching regions never leaves the canvas
                      // showing a stale, unrelated tab's worth of nothing.
                      setSelectedFlowRegion(region);
                      setDrilledDownStepId(undefined);
                    }}
                  />
                  <ProcessStepTabs
                    steps={stepsInRegion}
                    selectedStepId={drilledDownStepId}
                    onSelect={setDrilledDownStepId}
                    onAddNew={activeLock ? undefined : () => setAddSectionOpen(true)}
                  />
                </div>

                <SwimlaneCanvas
                  steps={visibleSteps}
                  allSteps={steps}
                  swimlaneSteps={stepsInRegion}
                  edges={edges}
                  riskStatements={riskStatements}
                  drilledDownStepId={drilledDownStepId}
                  employees={employees}
                  isLocked={!!activeLock}
                  onLabelEdge={handleLabelEdge}
                  onEditStep={handleEditStep}
                  onDeleteStep={handleDeleteStep}
                  onMoveStep={handleEditStep}
                  onCreateStep={handleCreateStepFromShape}
                  autoOpenStepId={autoOpenStepId}
                  onAutoOpenHandled={() => setAutoOpenStepId(undefined)}
                />

                {!activeLock && (
                <div className={styles.addStepForm}>
                  <h3 className={styles.cardTitle}>Add a step</h3>
                  <ProcessStepForm
                    value={newStepDraft}
                    onChange={setNewStepDraft}
                    employees={employees}
                    dependsOnOptions={addStepDependsOnOptions}
                    riskStatements={riskStatements}
                  />
                  <PrimaryButton
                    text={saving ? 'Adding...' : 'Add step'}
                    disabled={saving || !newStepDraft.actionDescription.trim()}
                    onClick={handleAddStep}
                  />
                </div>
                )}

                <SwimlaneComments comments={commentsForSwimlane} />
              </>
            )}
          </>
        )}
      </section>
    </div>
  );
};

export default SwimlaneStudio;

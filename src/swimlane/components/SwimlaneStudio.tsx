import * as React from 'react';
import {
  Spinner, MessageBar, MessageBarType, MessageBarButton, DefaultButton, PrimaryButton, Pivot, PivotItem,
  Dialog, DialogType, DialogFooter, TextField
} from '@fluentui/react';
import styles from './SwimlaneStudio.module.scss';
import type { ISwimlaneStudioProps } from './ISwimlaneStudioProps';
import { IProcessStep, getProgressId } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';
import { resolveDependencyEdges, stepIdsToDependsOnTokens, buildDependsOnOptions } from '../utils/dependencyResolution';
import {
  getCategoryId, getProcessGroupId, getCategoryName, getProcessGroupName, getProgressIdName,
  APQC_CATEGORY_NAMES, APQC_PROCESS_GROUP_NAMES, APQC_PROGRESS_ID_NAMES
} from '../utils/apqcHierarchy';
import HierarchyPicker from './HierarchyPicker';
import GlobalSearch from './GlobalSearch';
import ProcessStepTabs from './ProcessStepTabs';
import RegionFilter from './RegionFilter';
import EmployeesList from './EmployeesList';
import RiskRegisterList from './RiskRegisterList';
import SwimlaneCanvas from './SwimlaneCanvas';
import ImportCsvModal from './ImportCsvModal';
import NewProcessModal from './NewProcessModal';
import ProcessStepForm, { IProcessStepFormValue } from './ProcessStepForm';
import qleLogo from '../../assets/qle-logo.svg';

type MainTab = 'flows' | 'employees';

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
  riskLevelOverride: '',
  responsibleJobTitle: '',
  dependsOnStepIds: []
});

const SwimlaneStudio: React.FC<ISwimlaneStudioProps> = (props) => {
  const { dataService, onSignOut, signOutLabel } = props;

  const [steps, setSteps] = React.useState<IProcessStep[]>([]);
  const [employees, setEmployees] = React.useState<IEmployee[]>([]);
  const [riskStatements, setRiskStatements] = React.useState<IRiskStatement[]>([]);
  const [processGroupLabels, setProcessGroupLabels] = React.useState<IProcessGroupLabel[]>([]);
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
  const [selectedRegion, setSelectedRegion] = React.useState<string | undefined>(undefined);

  const [newStepDraft, setNewStepDraft] = React.useState<IProcessStepFormValue>(emptyStepDraft());
  const [saving, setSaving] = React.useState(false);
  const [importOpen, setImportOpen] = React.useState(false);
  const [newProcessOpen, setNewProcessOpen] = React.useState(false);
  const [bulkDeleteOpen, setBulkDeleteOpen] = React.useState(false);
  const [renameTarget, setRenameTarget] = React.useState<{ groupId: string; currentLabel: string } | undefined>(undefined);
  const [renameValue, setRenameValue] = React.useState('');
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
      dataService.getProcessGroupLabels()
    ])
      .then(([stepsResult, employeesResult, risksResult, groupLabelsResult]) => {
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

        setError(errors.length > 0 ? errors.join(' | ') : undefined);
        setLoading(false);
      });
  }, [dataService]);

  React.useEffect(() => { loadAll(); }, [loadAll]);

  const regions = React.useMemo(
    () => Array.from(new Set(employees.map(e => e.region).filter((r): r is string => !!r))),
    [employees]
  );

  // User-added names for Process Groups the static apqcHierarchy.ts table
  // doesn't already cover (see IProcessGroupLabel) - merged in wherever a
  // Process Group name is looked up or listed, same as the static table.
  const customGroupNames = React.useMemo(
    () => Object.fromEntries(processGroupLabels.map(l => [l.groupId, l.name])) as Record<string, string>,
    [processGroupLabels]
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

  // Scoped to the current swimlane (this Progress ID's own steps), not the
  // full cross-progress-ID dataset - a step realistically only ever
  // depends on something in its own flow, and listing all ~40 steps from
  // every unrelated flow made the real option buried in noise. No
  // excludeStepId - a brand-new step has no "self" to leave out, unlike
  // the edit panel's version of this same list.
  const addStepDependsOnOptions = React.useMemo(() => buildDependsOnOptions(stepsInProgressId), [stepsInProgressId]);

  const visibleSteps = React.useMemo(
    () => drilledDownStepId ? stepsInProgressId.filter(s => s.processStepId === drilledDownStepId) : stepsInProgressId,
    [stepsInProgressId, drilledDownStepId]
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

  const handleEditStep = (updated: IProcessStep): void => {
    const previous = steps.find(s => s.id === updated.id);
    setSteps(prev => prev.map(s => (s.id === updated.id ? updated : s)));
    dataService.updateProcessStep(updated).catch((err: Error) => setError(err.message));
    if (previous) setLastAction({ type: 'edit', previous });
  };

  const handleDeleteStep = (stepId: string): void => {
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
      const { previous } = lastAction;
      setSteps(prev => prev.map(s => (s.id === previous.id ? previous : s)));
      dataService.updateProcessStep(previous).catch((err: Error) => setError(err.message));
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

  const openRename = (groupId: string, currentLabel: string): void => {
    setRenameTarget({ groupId, currentLabel });
    setRenameValue(currentLabel);
  };

  // Renaming a group that already has a custom label updates that same
  // record; renaming one that's still showing its static apqcHierarchy.ts
  // name (or the generic "Process Group X.Y" fallback) creates a new
  // custom label instead, which then wins over the static name the same
  // way it already does for a brand-new group (see customGroupNames).
  const handleRenameSave = (): void => {
    if (!renameTarget) return;
    const trimmed = renameValue.trim();
    if (!trimmed) return;
    const existingLabel = processGroupLabels.find(l => l.groupId === renameTarget.groupId);
    if (existingLabel) {
      setProcessGroupLabels(prev => prev.map(l => (l.id === existingLabel.id ? { ...l, name: trimmed } : l)));
      dataService.updateProcessGroupLabel(existingLabel.id, trimmed).catch((err: Error) => setError(err.message));
    } else {
      dataService.addProcessGroupLabel(renameTarget.groupId, trimmed)
        .then(created => setProcessGroupLabels(prev => [...prev, created]))
        .catch((err: Error) => setError(err.message));
    }
    setRenameTarget(undefined);
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
    // Selects the new step's own Process Step ID tab, not "All" - without
    // this, "Add a step" right afterward defaulted to the bare Progress ID
    // as its processStepId (drilledDownStepId || selectedProgressId, with
    // drilledDownStepId unset), landing new steps in a second, separate
    // column group instead of continuing the one the user just started.
    setDrilledDownStepId(created.processStepId);
    setNewProcessOpen(false);
  };

  // Jumps straight to a step found via GlobalSearch - same drill-down
  // state a user would end up in by clicking all the way down through
  // Category -> Process Group -> Progress ID -> that step's own tab by
  // hand.
  const handleSearchNavigate = (step: IProcessStep): void => {
    setActiveTab('flows');
    setSelectedCategoryId(getCategoryId(step.processStepId));
    setSelectedProcessGroupId(getProcessGroupId(step.processStepId));
    setSelectedProgressId(getProgressId(step.processStepId));
    setDrilledDownStepId(step.processStepId);
  };

  const newProcessPrefix = selectedProcessGroupId ? `${selectedProcessGroupId}.` : selectedCategoryId ? `${selectedCategoryId}.` : '';

  const handleAddStep = (): void => {
    if (!selectedProgressId || !newStepDraft.actionDescription.trim()) return;
    const referenceStep = stepsInProgressId[0];
    const { dependsOnStepIds, ...fields } = newStepDraft;
    setSaving(true);
    dataService.addProcessStep({
      apqcTitle: referenceStep ? referenceStep.apqcTitle : selectedProgressId,
      processDescription: referenceStep ? referenceStep.processDescription : '',
      processStepId: drilledDownStepId || selectedProgressId,
      processStepName: referenceStep ? referenceStep.processStepName : '',
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
            <p className={styles.breadcrumb}>Finance process visualization</p>
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
              {selectedRegion ? ` — ${selectedRegion}` : ''}
            </p>
          ) : (
            <p className={styles.breadcrumb}>Finance process visualization</p>
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
        selectedRegion={selectedRegion}
        dataService={dataService}
        processStepIdPrefix={newProcessPrefix}
        knownProcessGroupIds={new Set([...Object.keys(APQC_PROCESS_GROUP_NAMES), ...Object.keys(customGroupNames)])}
        onDismiss={() => setNewProcessOpen(false)}
        onCreated={handleProcessCreated}
        onGroupLabelCreated={handleGroupLabelCreated}
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
        dialogContentProps={{ type: DialogType.normal, title: `Rename ${renameTarget?.groupId || ''}` }}
      >
        <TextField
          label="Process Group name"
          value={renameValue}
          onChange={(_e, v) => setRenameValue(v || '')}
          onKeyDown={e => { if (e.key === 'Enter') handleRenameSave(); }}
        />
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => setRenameTarget(undefined)} />
          <PrimaryButton text="Save" onClick={handleRenameSave} disabled={!renameValue.trim()} />
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
          onLinkClick={(item?: PivotItem) => setActiveTab(item?.props.itemKey === 'employees' ? 'employees' : 'flows')}
        >
          <PivotItem headerText="Process Flows" itemKey="flows" />
          <PivotItem headerText="Employees" itemKey="employees" />
        </Pivot>

        {activeTab === 'employees' ? (
          <EmployeesList employees={employees} />
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
                  onAddNew={() => setNewProcessOpen(true)}
                  addNewLabel="+ Add new process group"
                  onRename={openRename}
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
                  getLabel={(id, sampleStep) => sampleStep?.processDescription || getProgressIdName(id)}
                  onSelect={setSelectedProgressId}
                  allGroupIds={Object.keys(APQC_PROGRESS_ID_NAMES).filter(id => getProcessGroupId(id) === selectedProcessGroupId)}
                  onAddNew={() => setNewProcessOpen(true)}
                  addNewLabel="+ Add new progress ID"
                />
              </>
            ) : (
              <>
                <div className={styles.toolbar}>
                  <DefaultButton text="Back to Progress IDs" onClick={() => { setSelectedProgressId(undefined); setDrilledDownStepId(undefined); }} />
                  <PrimaryButton text="+ Add new process" onClick={() => setNewProcessOpen(true)} />
                  <DefaultButton text="Import CSV" iconProps={{ iconName: 'Upload' }} onClick={() => setImportOpen(true)} />
                  <DefaultButton
                    text={`Delete flow (${visibleSteps.length})`}
                    iconProps={{ iconName: 'Delete' }}
                    onClick={() => setBulkDeleteOpen(true)}
                    disabled={visibleSteps.length === 0}
                    styles={{ root: { borderColor: 'var(--risk-high)' }, label: { color: 'var(--risk-high)' }, icon: { color: 'var(--risk-high)' } }}
                  />
                  <ProcessStepTabs steps={stepsInProgressId} selectedStepId={drilledDownStepId} onSelect={setDrilledDownStepId} />
                  <RegionFilter regions={regions} selectedRegion={selectedRegion} onChange={setSelectedRegion} />
                </div>

                <SwimlaneCanvas
                  steps={visibleSteps}
                  allSteps={steps}
                  swimlaneSteps={stepsInProgressId}
                  edges={edges}
                  riskStatements={riskStatements}
                  drilledDownStepId={drilledDownStepId}
                  employees={employees}
                  selectedRegion={selectedRegion}
                  onLabelEdge={handleLabelEdge}
                  onEditStep={handleEditStep}
                  onDeleteStep={handleDeleteStep}
                  onMoveStep={handleEditStep}
                />

                <div className={styles.addStepForm}>
                  <h3 className={styles.cardTitle}>Add a step</h3>
                  <ProcessStepForm
                    value={newStepDraft}
                    onChange={setNewStepDraft}
                    employees={employees}
                    selectedRegion={selectedRegion}
                    dependsOnOptions={addStepDependsOnOptions}
                  />
                  <PrimaryButton
                    text={saving ? 'Adding...' : 'Add step'}
                    disabled={saving || !newStepDraft.actionDescription.trim()}
                    onClick={handleAddStep}
                  />
                </div>
              </>
            )}

            <RiskRegisterList riskStatements={riskStatements} />
          </>
        )}
      </section>
    </div>
  );
};

export default SwimlaneStudio;

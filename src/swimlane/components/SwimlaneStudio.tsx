import * as React from 'react';
import { Spinner, MessageBar, MessageBarType, DefaultButton, PrimaryButton, Pivot, PivotItem } from '@fluentui/react';
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
  const [activeTab, setActiveTab] = React.useState<MainTab>('flows');

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
    setSteps(prev => prev.map(s => (s.id === updated.id ? updated : s)));
    dataService.updateProcessStep(updated).catch((err: Error) => setError(err.message));
  };

  const handleDeleteStep = (stepId: string): void => {
    setSteps(prev => prev.filter(s => s.id !== stepId));
    dataService.deleteProcessStep(stepId).catch((err: Error) => setError(err.message));
  };

  const handleImported = (created: IProcessStep[]): void => {
    setSteps(prev => [...prev, ...created]);
  };

  const handleGroupLabelCreated = (created: IProcessGroupLabel): void => {
    setProcessGroupLabels(prev => [...prev, created]);
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

      <section className={styles.swimlaneStudio}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
            {error}
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

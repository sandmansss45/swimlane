import * as React from 'react';
import { Spinner, MessageBar, MessageBarType, DefaultButton, PrimaryButton, TextField, Pivot, PivotItem } from '@fluentui/react';
import styles from './SwimlaneStudio.module.scss';
import type { ISwimlaneStudioProps } from './ISwimlaneStudioProps';
import { IProcessStep, getProgressId } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';
import { resolveDependencyEdges } from '../utils/dependencyResolution';
import ProgressIdPicker from './ProgressIdPicker';
import ProcessStepTabs from './ProcessStepTabs';
import RegionFilter from './RegionFilter';
import EmployeePicker from './EmployeePicker';
import EmployeesList from './EmployeesList';
import RiskRegisterList from './RiskRegisterList';
import SwimlaneCanvas from './SwimlaneCanvas';
import ImportCsvModal from './ImportCsvModal';
import qleLogo from '../../assets/qle-logo.svg';

type MainTab = 'flows' | 'employees';

const SwimlaneStudio: React.FC<ISwimlaneStudioProps> = (props) => {
  const { dataService } = props;

  const [steps, setSteps] = React.useState<IProcessStep[]>([]);
  const [employees, setEmployees] = React.useState<IEmployee[]>([]);
  const [riskStatements, setRiskStatements] = React.useState<IRiskStatement[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const [selectedProgressId, setSelectedProgressId] = React.useState<string | undefined>(undefined);
  const [drilledDownStepId, setDrilledDownStepId] = React.useState<string | undefined>(undefined);
  const [selectedRegion, setSelectedRegion] = React.useState<string | undefined>(undefined);

  const [newActionDescription, setNewActionDescription] = React.useState('');
  const [newResponsibleJobTitle, setNewResponsibleJobTitle] = React.useState<string | undefined>(undefined);
  const [saving, setSaving] = React.useState(false);
  const [importOpen, setImportOpen] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<MainTab>('flows');

  const loadAll = React.useCallback(() => {
    setLoading(true);
    setError(undefined);
    Promise.all([dataService.getProcessSteps(), dataService.getEmployees(), dataService.getRiskStatements()])
      .then(([loadedSteps, loadedEmployees, loadedRisks]) => {
        setSteps(loadedSteps);
        setEmployees(loadedEmployees);
        setRiskStatements(loadedRisks);
        setLoading(false);
      })
      .catch((err: Error) => {
        setError(err.message || 'Failed to load data.');
        setLoading(false);
      });
  }, [dataService]);

  React.useEffect(() => { loadAll(); }, [loadAll]);

  const regions = React.useMemo(
    () => Array.from(new Set(employees.map(e => e.region).filter((r): r is string => !!r))),
    [employees]
  );

  const stepsInProgressId = React.useMemo(
    () => selectedProgressId ? steps.filter(s => getProgressId(s.processStepId) === selectedProgressId) : [],
    [steps, selectedProgressId]
  );

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

  const handleAddStep = (): void => {
    if (!selectedProgressId || !newActionDescription.trim()) return;
    const referenceStep = stepsInProgressId[0];
    setSaving(true);
    dataService.addProcessStep({
      apqcTitle: referenceStep ? referenceStep.apqcTitle : selectedProgressId,
      processDescription: referenceStep ? referenceStep.processDescription : '',
      processStepId: drilledDownStepId || selectedProgressId,
      processStepName: referenceStep ? referenceStep.processStepName : '',
      actionType: 'Execute (Within Limits)',
      action: '',
      actionDescription: newActionDescription.trim(),
      responsibleJobTitle: newResponsibleJobTitle || '',
      shapeOverride: '',
      dependsOn: []
    })
      .then(created => {
        setSteps(prev => [...prev, created]);
        setNewActionDescription('');
        setNewResponsibleJobTitle(undefined);
        setSaving(false);
      })
      .catch((err: Error) => {
        setError(err.message);
        setSaving(false);
      });
  };

  if (loading) {
    return <Spinner label="Loading process data..." />;
  }

  return (
    <div className={styles.page}>
      <header className={styles.appBar}>
        <img src={qleLogo} className={styles.appBarLogo} alt="Quantum Leap Energy" />
        <div>
          <h2 className={styles.title}>Swimlane Studio</h2>
          {selectedProgressId ? (
            <p className={styles.breadcrumb}>
              {selectedProgressId}
              {drilledDownStepId ? ` / ${drilledDownStepId}` : ''}
              {selectedRegion ? ` — ${selectedRegion}` : ''}
            </p>
          ) : (
            <p className={styles.breadcrumb}>Finance process visualization</p>
          )}
        </div>
        <div className={styles.appBarActions}>
          <DefaultButton
            text="Import CSV"
            iconProps={{ iconName: 'Upload' }}
            onClick={() => setImportOpen(true)}
            styles={{
              root: {
                background: 'rgba(255,255,255,0.08)',
                border: '1px solid rgba(255,255,255,0.3)',
                borderRadius: 8
              },
              rootHovered: { background: 'rgba(255,255,255,0.18)', border: '1px solid rgba(255,255,255,0.45)' },
              rootPressed: { background: 'rgba(255,255,255,0.24)' },
              label: { color: '#fff', fontWeight: 600 },
              icon: { color: '#fff' }
            }}
          />
        </div>
      </header>

      <ImportCsvModal
        isOpen={importOpen}
        dataService={dataService}
        insertionIndex={steps.length}
        onDismiss={() => setImportOpen(false)}
        onImported={handleImported}
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
            {!selectedProgressId && (
              <div className={styles.intro}>
                <h3>Select a process to explore</h3>
                <p>Pick a Progress ID below to open its swimlane - drill into an individual step, filter by region, or add and edit tasks directly.</p>
              </div>
            )}

            {!selectedProgressId ? (
            <ProgressIdPicker steps={steps} onSelect={setSelectedProgressId} />
          ) : (
            <>
              <div className={styles.toolbar}>
                <DefaultButton text="Back to Progress IDs" onClick={() => { setSelectedProgressId(undefined); setDrilledDownStepId(undefined); }} />
                <ProcessStepTabs steps={stepsInProgressId} selectedStepId={drilledDownStepId} onSelect={setDrilledDownStepId} />
                <RegionFilter regions={regions} selectedRegion={selectedRegion} onChange={setSelectedRegion} />
              </div>

              <SwimlaneCanvas
                steps={visibleSteps}
                allSteps={steps}
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
                <TextField
                  label="New step - action description"
                  value={newActionDescription}
                  onChange={(_e, v) => setNewActionDescription(v || '')}
                />
                <EmployeePicker
                  employees={employees}
                  selectedRegion={selectedRegion}
                  actionType="Execute (Within Limits)"
                  value={newResponsibleJobTitle}
                  onChange={setNewResponsibleJobTitle}
                />
                <PrimaryButton text={saving ? 'Adding...' : 'Add step'} disabled={saving} onClick={handleAddStep} />
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

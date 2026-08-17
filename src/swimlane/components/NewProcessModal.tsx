import * as React from 'react';
import { Modal, PrimaryButton, DefaultButton, TextField, MessageBar, MessageBarType, Spinner } from '@fluentui/react';
import { IProcessStep } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IDataService } from '../services/IDataService';
import { getCategoryId, getProcessGroupId, getCategoryName, getProcessGroupName } from '../utils/apqcHierarchy';
import { getProgressId } from '../models/IProcessStep';
import { stepIdsToDependsOnTokens, buildDependsOnOptions } from '../utils/dependencyResolution';
import ProcessStepForm, { IProcessStepFormValue } from './ProcessStepForm';
import styles from './NewProcessModal.module.scss';

export interface INewProcessModalProps {
  isOpen: boolean;
  steps: IProcessStep[];
  employees: IEmployee[];
  selectedRegion: string | undefined;
  dataService: IDataService;
  // Wherever the user opened this from (a category or process group
  // already drilled into) - prefills the Process Step ID so continuing
  // that context doesn't mean retyping it.
  processStepIdPrefix: string;
  onDismiss: () => void;
  onCreated: (created: IProcessStep) => void;
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

// Every other way of getting a step into the app (Add a step, CSV import)
// needs an existing Progress ID to attach to - there was no way to start
// a brand new one (e.g. Category 7, which has zero steps today) short of
// a CSV. This asks for the handful of identity fields a first step in a
// new area needs (Process Step ID plus the two labels nothing else can
// infer) and otherwise reuses the same ProcessStepForm as everywhere else.
const NewProcessModal: React.FC<INewProcessModalProps> = ({
  isOpen, steps, employees, selectedRegion, dataService, processStepIdPrefix, onDismiss, onCreated
}) => {
  const [processStepId, setProcessStepId] = React.useState('');
  const [processDescription, setProcessDescription] = React.useState('');
  const [processStepName, setProcessStepName] = React.useState('');
  const [stepValue, setStepValue] = React.useState<IProcessStepFormValue>(emptyStepDraft());
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Re-seed the prefix whenever the modal opens - reopening it from a
  // different picker level shouldn't leave the previous attempt's ID sitting
  // in the field.
  React.useEffect(() => {
    if (isOpen) {
      setProcessStepId(processStepIdPrefix);
      setProcessDescription('');
      setProcessStepName('');
      setStepValue(emptyStepDraft());
      setError(undefined);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen]);

  const dependsOnOptions = React.useMemo(() => buildDependsOnOptions(steps), [steps]);

  const trimmedId = processStepId.trim();
  const idLooksValid = /^\d+(\.\d+){3,}$/.test(trimmedId);
  const preview = idLooksValid
    ? `${getCategoryName(getCategoryId(trimmedId))} · ${getProcessGroupName(getProcessGroupId(trimmedId))} · Progress ID ${getProgressId(trimmedId)}`
    : undefined;

  const canSubmit = idLooksValid
    && processDescription.trim().length > 0
    && processStepName.trim().length > 0
    && stepValue.actionDescription.trim().length > 0
    && !saving;

  const handleCreate = (): void => {
    if (!canSubmit) return;
    setSaving(true);
    setError(undefined);
    const { dependsOnStepIds, ...fields } = stepValue;
    dataService.addProcessStep({
      apqcTitle: trimmedId,
      processDescription: processDescription.trim(),
      processStepId: trimmedId,
      processStepName: processStepName.trim(),
      ...fields,
      actionDescription: fields.actionDescription.trim(),
      dependsOn: stepIdsToDependsOnTokens(steps, dependsOnStepIds)
    })
      .then(created => {
        setSaving(false);
        onCreated(created);
      })
      .catch((err: Error) => {
        setSaving(false);
        setError(err.message);
      });
  };

  return (
    <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking={false} containerClassName={styles.modal}>
      <div className={styles.header}>
        <h3>Add a new process</h3>
        <p>Starts a brand new Process Step ID - use this for an area with nothing in it yet (e.g. a different APQC category). To add another step to an existing flow, open that flow and use "Add a step" instead.</p>
      </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.banner}>
          {error}
        </MessageBar>
      )}

      <TextField
        label="Process Step ID"
        placeholder="e.g. 7.1.1.1"
        value={processStepId}
        onChange={(_e, v) => setProcessStepId(v || '')}
        errorMessage={trimmedId.length > 0 && !idLooksValid ? 'Needs at least 4 dot-separated numbers, e.g. 7.1.1.1' : undefined}
      />
      {preview && <p className={styles.preview}>{preview}</p>}

      <TextField
        label="Process description"
        placeholder="What this flow does, e.g. Recruit and select employees"
        value={processDescription}
        onChange={(_e, v) => setProcessDescription(v || '')}
      />
      <TextField
        label="Process step name"
        placeholder="Column header for this step, e.g. Screen candidates"
        value={processStepName}
        onChange={(_e, v) => setProcessStepName(v || '')}
      />

      <div className={styles.divider} />

      <ProcessStepForm
        value={stepValue}
        onChange={setStepValue}
        employees={employees}
        selectedRegion={selectedRegion}
        dependsOnOptions={dependsOnOptions}
      />

      <div className={styles.footer}>
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        <PrimaryButton text={saving ? 'Creating...' : 'Create process'} onClick={handleCreate} disabled={!canSubmit} />
        {saving && <Spinner />}
      </div>
    </Modal>
  );
};

export default NewProcessModal;

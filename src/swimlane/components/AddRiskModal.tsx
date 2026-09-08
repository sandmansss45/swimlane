import * as React from 'react';
import { Modal, PrimaryButton, DefaultButton, TextField } from '@fluentui/react';
import { IRiskStatement } from '../models/IRiskStatement';
import { IDataService } from '../services/IDataService';
import styles from './AddHierarchyShellModal.module.scss';

export interface IAddRiskModalProps {
  isOpen: boolean;
  dataService: IDataService;
  onDismiss: () => void;
  onCreated: (created: IRiskStatement) => void;
}

// Writes a real row into "risk register data" - CONFIRMED 2026-08-21, an
// explicit user choice despite that list otherwise being a standing
// enterprise register this app doesn't own (see the schema comment on
// IDataService.addRiskStatement). All eight real columns are editable
// here, same set getRiskStatements/RiskRegisterList already read.
const AddRiskModal: React.FC<IAddRiskModalProps> = ({ isOpen, dataService, onDismiss, onCreated }) => {
  const [riskId, setRiskId] = React.useState('');
  const [category, setCategory] = React.useState('');
  const [apqcProcessArea, setApqcProcessArea] = React.useState('');
  const [process, setProcess] = React.useState('');
  const [riskStatement, setRiskStatement] = React.useState('');
  const [rootCause, setRootCause] = React.useState('');
  const [riskResponse, setRiskResponse] = React.useState('');
  const [riskOwner, setRiskOwner] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    if (isOpen) {
      setRiskId('');
      setCategory('');
      setApqcProcessArea('');
      setProcess('');
      setRiskStatement('');
      setRootCause('');
      setRiskResponse('');
      setRiskOwner('');
      setError(undefined);
    }
  }, [isOpen]);

  const trimmedStatement = riskStatement.trim();
  const canSubmit = trimmedStatement.length > 0 && !saving;

  const handleCreate = (): void => {
    if (!canSubmit) return;
    setSaving(true);
    setError(undefined);
    dataService.addRiskStatement({
      riskId: riskId.trim(),
      category: category.trim(),
      apqcProcessArea: apqcProcessArea.trim(),
      process: process.trim(),
      riskStatement: trimmedStatement,
      rootCause: rootCause.trim(),
      riskResponse: riskResponse.trim(),
      riskOwner: riskOwner.trim()
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
        <h3>Add a risk</h3>
        <p>Writes a new row into the risk register - Risk Statement is the only required field.</p>
      </div>

      {error && <p className={styles.error}>{error}</p>}

      <TextField label="Risk ID" placeholder="e.g. OP-042" value={riskId} onChange={(_e, v) => setRiskId(v || '')} />
      <TextField label="Category" placeholder="e.g. Operational / Financial Controls" value={category} onChange={(_e, v) => setCategory(v || '')} />
      <TextField
        label="APQC process area"
        placeholder="e.g. Manage Financial Resources"
        value={apqcProcessArea}
        onChange={(_e, v) => setApqcProcessArea(v || '')}
      />
      <TextField
        label="Process"
        placeholder="e.g. Process accounts payable (AP)"
        value={process}
        onChange={(_e, v) => setProcess(v || '')}
      />
      <TextField
        label="Risk statement"
        placeholder="Describe the risk"
        value={riskStatement}
        onChange={(_e, v) => setRiskStatement(v || '')}
        multiline
        rows={3}
      />
      <TextField label="Root cause" placeholder="Why this risk exists" value={rootCause} onChange={(_e, v) => setRootCause(v || '')} multiline rows={2} />
      <TextField
        label="Risk response"
        placeholder="e.g. Mitigate"
        value={riskResponse}
        onChange={(_e, v) => setRiskResponse(v || '')}
      />
      <TextField
        label="Risk owner"
        placeholder="e.g. Finance Manager"
        value={riskOwner}
        onChange={(_e, v) => setRiskOwner(v || '')}
        onKeyDown={e => { if (e.key === 'Enter') handleCreate(); }}
      />

      <div className={styles.footer}>
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        <PrimaryButton text={saving ? 'Adding...' : 'Add'} onClick={handleCreate} disabled={!canSubmit} />
      </div>
    </Modal>
  );
};

export default AddRiskModal;

import * as React from 'react';
import { Modal, PrimaryButton, DefaultButton, TextField } from '@fluentui/react';
import { IEmployee } from '../models/IEmployee';
import { IDataService } from '../services/IDataService';
import styles from './AddHierarchyShellModal.module.scss';

export interface IAddEmployeeModalProps {
  isOpen: boolean;
  dataService: IDataService;
  onDismiss: () => void;
  onCreated: (created: IEmployee) => void;
}

// Writes a real row into "QLE Existing Organisation" - CONFIRMED
// 2026-08-19, an explicit user choice despite that list otherwise being
// read-only/owned elsewhere (see the schema comment on
// IDataService.addEmployee). Job title and Department only, same two
// fields EmployeesList/EmployeePicker already read - no name field, by
// the same explicit "job title only, never a person's name" design rule
// as everywhere else this list is used (see models/IEmployee.ts).
const AddEmployeeModal: React.FC<IAddEmployeeModalProps> = ({ isOpen, dataService, onDismiss, onCreated }) => {
  const [jobTitle, setJobTitle] = React.useState('');
  const [department, setDepartment] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    if (isOpen) {
      setJobTitle('');
      setDepartment('');
      setError(undefined);
    }
  }, [isOpen]);

  const trimmedJobTitle = jobTitle.trim();
  const canSubmit = trimmedJobTitle.length > 0 && !saving;

  const handleCreate = (): void => {
    if (!canSubmit) return;
    setSaving(true);
    setError(undefined);
    dataService.addEmployee(trimmedJobTitle, department.trim())
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
        <h3>Add an employee</h3>
        <p>Job title and department only - names aren&apos;t tracked here, same as the rest of this directory.</p>
      </div>

      {error && <p className={styles.error}>{error}</p>}

      <TextField
        label="Job title"
        placeholder="e.g. Finance Manager"
        value={jobTitle}
        onChange={(_e, v) => setJobTitle(v || '')}
      />
      <TextField
        label="Department"
        placeholder="e.g. 6002 - Finance"
        value={department}
        onChange={(_e, v) => setDepartment(v || '')}
        onKeyDown={e => { if (e.key === 'Enter') handleCreate(); }}
      />

      <div className={styles.footer}>
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        <PrimaryButton text={saving ? 'Adding...' : 'Add'} onClick={handleCreate} disabled={!canSubmit} />
      </div>
    </Modal>
  );
};

export default AddEmployeeModal;

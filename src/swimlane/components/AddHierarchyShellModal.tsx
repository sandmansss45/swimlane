import * as React from 'react';
import { Modal, PrimaryButton, DefaultButton, TextField } from '@fluentui/react';
import { IDataService } from '../services/IDataService';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';
import { IProcessIdLabel } from '../models/IProcessIdLabel';
import { ICategoryLabel } from '../models/ICategoryLabel';
import styles from './AddHierarchyShellModal.module.scss';

export type HierarchyShellLevel = 'category' | 'processGroup' | 'processId';

export interface IAddHierarchyShellModalProps {
  isOpen: boolean;
  level: HierarchyShellLevel;
  // Already includes the trailing dot, e.g. "13." - the user completes the
  // rest (one more segment for a Process Group, two more for a Process ID).
  idPrefix: string;
  dataService: IDataService;
  onDismiss: () => void;
  // The caller can tell which was created from the shape alone (a
  // Category label has categoryId, a Process Group label has groupId, a
  // Process ID label has processId) - lets it both cache the new label
  // and navigate straight into it, same as every other "created
  // something, land in it" flow in this app.
  onCreated: (created: ICategoryLabel | IProcessGroupLabel | IProcessIdLabel) => void;
}

const LEVEL_COPY: Record<HierarchyShellLevel, { title: string; idLabel: string; nameLabel: string; namePlaceholder: string; segments: number }> = {
  category: {
    title: 'Add a new Category',
    idLabel: 'Category ID',
    nameLabel: 'Category name',
    namePlaceholder: 'e.g. Manage Supply Chain',
    segments: 1
  },
  processGroup: {
    title: 'Add a new Process Group',
    idLabel: 'Process Group ID',
    nameLabel: 'Process Group name',
    namePlaceholder: 'e.g. Manage petty cash',
    segments: 2
  },
  processId: {
    title: 'Add a new Process ID',
    idLabel: 'Process ID',
    nameLabel: 'Process ID name',
    namePlaceholder: 'e.g. Manage portfolio project',
    segments: 3
  }
};

// Creating a process previously always meant going all the way down to a
// full 4-segment Activity with a complete step form - there was no way to
// just reserve and name an empty Process Group or Process ID first and
// fill in real steps later, which is exactly backwards from how someone
// actually plans out a framework. This creates just the shell (an ID plus
// a name, persisted the same way a rename does - see IProcessGroupLabel/
// IProcessIdLabel) with no step content at all; "Add a step" handles
// populating it afterward.
const AddHierarchyShellModal: React.FC<IAddHierarchyShellModalProps> = ({
  isOpen, level, idPrefix, dataService, onDismiss, onCreated
}) => {
  const [id, setId] = React.useState('');
  const [name, setName] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const copy = LEVEL_COPY[level];

  React.useEffect(() => {
    if (isOpen) {
      setId(idPrefix);
      setName('');
      setError(undefined);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, idPrefix]);

  const trimmedId = id.trim();
  const idLooksValid = new RegExp(`^\\d+(\\.\\d+){${copy.segments - 1}}$`).test(trimmedId);
  const canSubmit = idLooksValid && name.trim().length > 0 && !saving;

  const handleCreate = (): void => {
    if (!canSubmit) return;
    setSaving(true);
    setError(undefined);
    const trimmedName = name.trim();
    const request = level === 'category'
      ? dataService.addCategoryLabel(trimmedId, trimmedName)
      : level === 'processGroup'
        ? dataService.addProcessGroupLabel(trimmedId, trimmedName)
        : dataService.addProcessIdLabel(trimmedId, trimmedName);
    request
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
        <h3>{copy.title}</h3>
        <p>Just names and reserves this ID - add real steps to it afterward with "Add a step", whenever you're ready.</p>
      </div>

      {error && <p className={styles.error}>{error}</p>}

      <TextField
        label={copy.idLabel}
        value={id}
        onChange={(_e, v) => setId(v || '')}
        errorMessage={trimmedId.length > 0 && !idLooksValid
          ? (copy.segments === 1
            ? `Needs to be a single number, e.g. ${Array(copy.segments - idPrefix.split('.').filter(Boolean).length).fill('1').join('.')}`
            : `Needs exactly ${copy.segments} dot-separated numbers, e.g. ${idPrefix}${Array(copy.segments - idPrefix.split('.').filter(Boolean).length).fill('1').join('.')}`)
          : undefined}
      />
      <TextField
        label={copy.nameLabel}
        placeholder={copy.namePlaceholder}
        value={name}
        onChange={(_e, v) => setName(v || '')}
        onKeyDown={e => { if (e.key === 'Enter') handleCreate(); }}
      />

      <div className={styles.footer}>
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        <PrimaryButton text={saving ? 'Creating...' : 'Create'} onClick={handleCreate} disabled={!canSubmit} />
      </div>
    </Modal>
  );
};

export default AddHierarchyShellModal;

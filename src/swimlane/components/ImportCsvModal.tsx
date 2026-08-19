import * as React from 'react';
import { Modal, DefaultButton, PrimaryButton, MessageBar, MessageBarType, Spinner } from '@fluentui/react';
import { IProcessStep } from '../models/IProcessStep';
import { IDataService } from '../services/IDataService';
import { buildImportPreview, prepareStepsForImport, ICsvImportPreview } from '../utils/csvImport';
import styles from './ImportCsvModal.module.scss';

export interface IImportCsvModalProps {
  isOpen: boolean;
  dataService: IDataService;
  insertionIndex: number;
  // Steps already in the diagram - purely to warn about Process Step ID
  // collisions before committing (see duplicateStepIds below). Import
  // itself never touches these; addProcessSteps only ever creates new
  // items, so a colliding row lands as a genuine duplicate sitting
  // alongside the original, not an overwrite.
  existingSteps: IProcessStep[];
  onDismiss: () => void;
  onImported: (created: IProcessStep[]) => void;
}

const MAX_PREVIEW_ROWS = 12;
const MAX_DUPLICATE_IDS_SHOWN = 5;

const ImportCsvModal: React.FC<IImportCsvModalProps> = ({ isOpen, dataService, insertionIndex, existingSteps, onDismiss, onImported }) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = React.useState<string | undefined>(undefined);
  const [preview, setPreview] = React.useState<ICsvImportPreview | undefined>(undefined);
  const [readError, setReadError] = React.useState<string | undefined>(undefined);
  const [importing, setImporting] = React.useState(false);

  const reset = (): void => {
    setFileName(undefined);
    setPreview(undefined);
    setReadError(undefined);
    setImporting(false);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleClose = (): void => {
    reset();
    onDismiss();
  };

  const handleFileChosen = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setReadError(undefined);
    setPreview(undefined);

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = typeof reader.result === 'string' ? reader.result : '';
        setPreview(buildImportPreview(text));
      } catch (err) {
        setReadError(err instanceof Error ? err.message : 'Could not parse this file.');
      }
    };
    reader.onerror = () => setReadError('Could not read this file.');
    reader.readAsText(file);
  };

  const handleImport = (): void => {
    if (!preview || preview.rows.length === 0) return;
    setImporting(true);
    const prepared = prepareStepsForImport(preview.rows, insertionIndex);
    dataService.addProcessSteps(prepared)
      .then(created => {
        onImported(created);
        handleClose();
      })
      .catch((err: Error) => {
        setReadError(err.message || 'Import failed.');
        setImporting(false);
      });
  };

  const distinctStepIds = React.useMemo(() => {
    if (!preview) return 0;
    return new Set(preview.rows.map(r => r.step.processStepId)).size;
  }, [preview]);

  // Process Step IDs this file shares with steps already in the diagram -
  // addProcessSteps only ever creates new items (see IDataService), so a
  // collision here means the import is about to add a second, separate
  // copy sitting alongside the original rather than replacing or merging
  // it. Surfaced as a warning rather than blocked outright - re-importing
  // an updated version of an already-loaded file is a legitimate thing to
  // want to do (e.g. after deleting the old rows first), just not silently.
  const duplicateStepIds = React.useMemo(() => {
    if (!preview) return [];
    const existingIds = new Set(existingSteps.map(s => s.processStepId).filter(Boolean));
    return Array.from(new Set(preview.rows.map(r => r.step.processStepId).filter(id => id && existingIds.has(id))));
  }, [preview, existingSteps]);

  return (
    <Modal isOpen={isOpen} onDismiss={handleClose} isBlocking={false} containerClassName={styles.modal}>
      <div className={styles.header}>
        <h3>Import from CSV</h3>
        <p>
          Upload a process export - <code>APQC Title, Process Description, Process Step ID, Process Step Name,
          Action Type, Action, Action Description, ResponsibleJobTitle, ShapeOverride, DependsOn, Region</code>.
          Column order doesn&apos;t matter; each is matched by its heading. Region is optional - if the file
          doesn&apos;t have that column but APQC Title already ends "- UK" (or "- US"/"- SA"), that's picked up
          automatically.
        </p>
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,text/csv"
        className={styles.hiddenInput}
        onChange={handleFileChosen}
      />

      <div className={styles.pickRow}>
        <DefaultButton text={fileName ? 'Choose a different file' : 'Choose CSV file'} onClick={() => fileInputRef.current?.click()} />
        {fileName && <span className={styles.fileName}>{fileName}</span>}
      </div>

      {readError && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.banner}>
          {readError}
        </MessageBar>
      )}

      {preview && preview.warnings.map((w, i) => (
        <MessageBar key={i} messageBarType={MessageBarType.warning} className={styles.banner}>
          {w}
        </MessageBar>
      ))}

      {duplicateStepIds.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} className={styles.banner}>
          {duplicateStepIds.length} process step{duplicateStepIds.length === 1 ? '' : 's'} in this file
          ({duplicateStepIds.slice(0, MAX_DUPLICATE_IDS_SHOWN).join(', ')}
          {duplicateStepIds.length > MAX_DUPLICATE_IDS_SHOWN ? `, +${duplicateStepIds.length - MAX_DUPLICATE_IDS_SHOWN} more` : ''})
          {' '}{duplicateStepIds.length === 1 ? 'already exists' : 'already exist'} in this diagram. Importing
          won&apos;t replace or merge {duplicateStepIds.length === 1 ? 'it' : 'them'} - {duplicateStepIds.length === 1 ? 'this row' : 'these rows'} will
          be added as {duplicateStepIds.length === 1 ? 'its own duplicate step' : 'their own duplicate steps'}.
          Delete the existing one{duplicateStepIds.length === 1 ? '' : 's'} first if you meant to replace {duplicateStepIds.length === 1 ? 'it' : 'them'}.
        </MessageBar>
      )}

      {preview && preview.autoLinkedCount > 0 && (
        <MessageBar messageBarType={MessageBarType.info} className={styles.banner}>
          {preview.autoLinkedCount} row{preview.autoLinkedCount === 1 ? '' : 's'} had no Depends On value, so
          {preview.autoLinkedCount === 1 ? " it's" : " they've"} been linked to the previous step in the same
          flow by default. Adjust any of these afterward from the step&apos;s edit panel - corrections save
          normally.
        </MessageBar>
      )}

      {preview && preview.autoRegionedCount > 0 && (
        <MessageBar messageBarType={MessageBarType.info} className={styles.banner}>
          {preview.autoRegionedCount} row{preview.autoRegionedCount === 1 ? '' : 's'} had no Region column, so
          {preview.autoRegionedCount === 1 ? " it's" : " they've"} been tagged automatically from the "- UK" /
          "- US" / "- SA" suffix already on APQC Title.
        </MessageBar>
      )}

      {preview && preview.rows.length > 0 && (
        <>
          <p className={styles.summary}>
            <strong>{preview.rows.length}</strong> row{preview.rows.length === 1 ? '' : 's'} ready to import
            across <strong>{distinctStepIds}</strong> process step{distinctStepIds === 1 ? '' : 's'}.
          </p>

          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Process Step ID</th>
                  <th>Step Name</th>
                  <th>Action</th>
                  <th>Responsible</th>
                  <th>Region</th>
                  <th>Depends On</th>
                </tr>
              </thead>
              <tbody>
                {preview.rows.slice(0, MAX_PREVIEW_ROWS).map((row, i) => (
                  <tr key={i}>
                    <td>
                      {row.step.processStepId}
                      {duplicateStepIds.includes(row.step.processStepId) && (
                        <span className={styles.duplicateTag} title="Already exists in this diagram">dup</span>
                      )}
                    </td>
                    <td>{row.step.processStepName}</td>
                    <td>{row.step.action}</td>
                    <td>{row.step.responsibleJobTitle}</td>
                    <td>
                      {row.step.region || <span className={styles.muted}>-</span>}
                      {row.autoRegioned && <span className={styles.autoTag}>auto</span>}
                    </td>
                    <td>
                      {row.step.dependsOn.join(', ') || <span className={styles.muted}>-</span>}
                      {row.autoLinked && <span className={styles.autoTag}>auto</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {preview.rows.length > MAX_PREVIEW_ROWS && (
              <p className={styles.moreRows}>+ {preview.rows.length - MAX_PREVIEW_ROWS} more row{preview.rows.length - MAX_PREVIEW_ROWS === 1 ? '' : 's'}</p>
            )}
          </div>
        </>
      )}

      <div className={styles.footer}>
        <DefaultButton text="Cancel" onClick={handleClose} disabled={importing} />
        <PrimaryButton
          text={importing ? 'Importing...' : preview ? `Import ${preview.rows.length} row${preview.rows.length === 1 ? '' : 's'}` : 'Import'}
          onClick={handleImport}
          disabled={importing || !preview || preview.rows.length === 0}
        />
        {importing && <Spinner />}
      </div>
    </Modal>
  );
};

export default ImportCsvModal;

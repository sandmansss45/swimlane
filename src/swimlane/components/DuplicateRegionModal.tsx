import * as React from 'react';
import { Modal, PrimaryButton, DefaultButton, ComboBox, IComboBoxOption, MessageBar, MessageBarType } from '@fluentui/react';
import { IProcessStep, KNOWN_FLOW_REGIONS } from '../models/IProcessStep';
import { IDataService } from '../services/IDataService';
import { buildDuplicatedSteps } from '../utils/duplicateRegion';
import styles from './AddHierarchyShellModal.module.scss';

export interface IDuplicateRegionModalProps {
  isOpen: boolean;
  dataService: IDataService;
  processId: string;
  sourceRegion: string;
  // The steps being duplicated (this Process ID, this region only), in
  // original dataset order - see buildDuplicatedSteps for why order matters.
  sourceSteps: IProcessStep[];
  // Full unfiltered dataset - needed both to resolve what a DependsOn
  // token on a source step originally pointed to, and as the insertion
  // point the new batch will land at (see buildDuplicatedSteps).
  allSteps: IProcessStep[];
  onDismiss: () => void;
  onDuplicated: (created: IProcessStep[]) => void;
}

// Lets someone stand up e.g. the US swimlane for a Process ID by copying
// an existing one (UK, say) wholesale rather than rebuilding every step by
// hand - added at a real user's request. Every field carries over as-is
// except Region; DependsOn links within the duplicated set are rewired to
// the new copies, links pointing outside the source region are dropped
// (see buildDuplicatedSteps) since there's nothing sensible to point them at.
const DuplicateRegionModal: React.FC<IDuplicateRegionModalProps> = ({
  isOpen, dataService, processId, sourceRegion, sourceSteps, allSteps, onDismiss, onDuplicated
}) => {
  const [targetRegion, setTargetRegion] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    if (isOpen) {
      setTargetRegion('');
      setError(undefined);
    }
  }, [isOpen]);

  // Regions this Process ID already has steps in, so the picker steers
  // toward a genuinely new one - offered as suggestions, not enforced,
  // since duplicating on top of an existing region (adding more steps
  // alongside what's already there) is unusual but not wrong.
  const existingRegions = React.useMemo(
    () => new Set(allSteps.filter(s => s.processStepId && s.region).map(s => s.region as string)),
    [allSteps]
  );
  const regionOptions: IComboBoxOption[] = KNOWN_FLOW_REGIONS
    .filter(r => r !== sourceRegion)
    .map(r => ({ key: r, text: r }));

  const trimmedTarget = targetRegion.trim();
  const canSubmit = trimmedTarget.length > 0 && trimmedTarget !== sourceRegion && sourceSteps.length > 0 && !saving;
  const targetCollides = trimmedTarget.length > 0 && existingRegions.has(trimmedTarget);

  // Doesn't depend on the target region - only on which DependsOn links
  // point outside the set being duplicated - so it's shown as soon as the
  // modal opens rather than waiting for a region to be chosen.
  const droppedDependencyCount = React.useMemo(
    () => buildDuplicatedSteps(allSteps, sourceSteps, '', allSteps.length).droppedDependencyCount,
    [allSteps, sourceSteps]
  );

  const handleDuplicate = (): void => {
    if (!canSubmit) return;
    setSaving(true);
    setError(undefined);
    const { steps } = buildDuplicatedSteps(allSteps, sourceSteps, trimmedTarget, allSteps.length);
    dataService.addProcessSteps(steps)
      .then(result => {
        setSaving(false);
        // Same "keep whatever succeeded" contract as CSV import - a
        // partial failure still hands up every step that's really there.
        if (result.created.length > 0) onDuplicated(result.created);
        if (result.failed.length === 0) {
          onDismiss();
          return;
        }
        setError(
          `${result.created.length} of ${steps.length} step${steps.length === 1 ? '' : 's'} duplicated. ` +
          `${result.failed.length} failed: ${result.failed[0].error}`
        );
      })
      .catch((err: Error) => {
        setSaving(false);
        setError(err.message);
      });
  };

  return (
    <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking={false} containerClassName={styles.modal}>
      <div className={styles.header}>
        <h3>Duplicate to another region</h3>
        <p>
          Copies all {sourceSteps.length} step{sourceSteps.length === 1 ? '' : 's'} of {sourceRegion
            ? <>the <strong>{sourceRegion}</strong> swimlane</>
            : <>the steps with <strong>no region set</strong></>}
          {' '}for {processId} into a new region - everything carries over (Depends On links included, rewired to the new copies),
          only Region changes. Edit the copies afterward for anything that genuinely differs by location.
        </p>
      </div>

      {error && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 12 } }}>{error}</MessageBar>}

      {targetCollides && (
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12 } }}>
          {processId} already has steps tagged "{trimmedTarget}" - duplicating won't replace or merge with them,
          these will be added as their own separate steps alongside what's already there.
        </MessageBar>
      )}

      {droppedDependencyCount > 0 && (
        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: 12 } }}>
          {droppedDependencyCount} Depends On link{droppedDependencyCount === 1 ? '' : 's'} point{droppedDependencyCount === 1 ? 's' : ''} outside
          {sourceRegion ? ` the ${sourceRegion} swimlane` : ' the steps being duplicated'} and won&apos;t carry over - reconnect {droppedDependencyCount === 1 ? 'it' : 'them'} afterward
          from the copies' edit panel if needed.
        </MessageBar>
      )}

      <ComboBox
        label="Duplicate to region"
        placeholder="Choose or type a region"
        text={targetRegion}
        allowFreeform
        autoComplete="on"
        options={regionOptions}
        onChange={(_e, option, _index, freeformValue) => setTargetRegion(option ? String(option.key) : (freeformValue || ''))}
      />

      <div className={styles.footer}>
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={saving} />
        <PrimaryButton
          text={saving ? 'Duplicating...' : `Duplicate ${sourceSteps.length} step${sourceSteps.length === 1 ? '' : 's'}`}
          onClick={handleDuplicate}
          disabled={!canSubmit}
        />
      </div>
    </Modal>
  );
};

export default DuplicateRegionModal;

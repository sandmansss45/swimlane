import * as React from 'react';
import { Dropdown, IDropdownOption, DefaultButton, IconButton } from '@fluentui/react';
import { IRiskStatement, IRiskLink, RiskLevel } from '../models/IRiskStatement';
import styles from './RiskLinkPicker.module.scss';

export interface IRiskLinkPickerProps {
  riskStatements: IRiskStatement[];
  value: IRiskLink[];
  onChange: (value: IRiskLink[]) => void;
}

const SEVERITY_OPTIONS: IDropdownOption[] = [
  { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' },
  { key: 'Low', text: 'Low' }
];

const SEVERITY_CLASS: Record<RiskLevel, string> = {
  High: styles.badgeHigh,
  Medium: styles.badgeMedium,
  Low: styles.badgeLow
};

// Two-step picker (category, then a risk within it) plus a manually-chosen
// severity for THIS step - confirmed design rule: severity isn't derived
// from the register's own numbers, it's a judgment call made at the point
// of tying a risk to a step, so it's asked for explicitly every time
// rather than defaulted or computed. Used inside ProcessStepForm, both for
// "Add a step" and the shape edit panel.
const RiskLinkPicker: React.FC<IRiskLinkPickerProps> = ({ riskStatements, value, onChange }) => {
  const [category, setCategory] = React.useState<string | undefined>(undefined);
  const [riskId, setRiskId] = React.useState<string | undefined>(undefined);
  const [severity, setSeverity] = React.useState<RiskLevel | undefined>(undefined);

  const risksById = React.useMemo(() => new Map(riskStatements.map(r => [r.id, r])), [riskStatements]);

  const categoryOptions: IDropdownOption[] = React.useMemo(
    () => Array.from(new Set(riskStatements.map(r => r.category).filter(Boolean))).sort().map(c => ({ key: c, text: c })),
    [riskStatements]
  );

  const riskOptions: IDropdownOption[] = React.useMemo(
    () => riskStatements
      .filter(r => r.category === category)
      // Already-linked risks aren't offered again - the picker is for
      // adding new links, not editing severity of an existing one
      // (remove and re-add covers that, same as everywhere else in this
      // form re-uses "remove chip, add fresh" instead of in-place edit).
      .filter(r => !value.some(link => link.riskId === r.id))
      .map(r => ({ key: r.id, text: r.riskId ? `${r.riskId} — ${r.riskStatement}` : r.riskStatement })),
    [riskStatements, category, value]
  );

  const canAdd = !!category && !!riskId && !!severity;

  const handleAdd = (): void => {
    if (!canAdd || !riskId || !severity) return;
    onChange([...value, { riskId, severity }]);
    setCategory(undefined);
    setRiskId(undefined);
    setSeverity(undefined);
  };

  const handleRemove = (targetRiskId: string): void => {
    onChange(value.filter(link => link.riskId !== targetRiskId));
  };

  return (
    <div className={styles.picker}>
      <label className={styles.label}>Linked risks</label>

      {value.length > 0 && (
        <ul className={styles.linkedList}>
          {value.map(link => {
            const risk = risksById.get(link.riskId);
            return (
              <li key={link.riskId} className={styles.linkedRow}>
                <span className={`${styles.badge} ${SEVERITY_CLASS[link.severity]}`}>{link.severity}</span>
                <span className={styles.linkedText}>
                  {risk ? risk.riskStatement : <em>Risk no longer in the register (id {link.riskId})</em>}
                </span>
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  title="Remove this linked risk"
                  ariaLabel="Remove this linked risk"
                  className={styles.removeButton}
                  onClick={() => handleRemove(link.riskId)}
                />
              </li>
            );
          })}
        </ul>
      )}

      {riskStatements.length === 0 ? (
        <p className={styles.empty}>No risks loaded from the Risk Register yet.</p>
      ) : (
        <div className={styles.addRow}>
          <Dropdown
            placeholder="Category"
            selectedKey={category}
            options={categoryOptions}
            onChange={(_e, option) => { setCategory(option ? String(option.key) : undefined); setRiskId(undefined); }}
            className={styles.addField}
          />
          <Dropdown
            placeholder="Risk"
            selectedKey={riskId}
            options={riskOptions}
            disabled={!category}
            onChange={(_e, option) => setRiskId(option ? String(option.key) : undefined)}
            className={styles.addField}
          />
          <Dropdown
            placeholder="Severity"
            selectedKey={severity}
            options={SEVERITY_OPTIONS}
            disabled={!riskId}
            onChange={(_e, option) => setSeverity(option ? (String(option.key) as RiskLevel) : undefined)}
            className={styles.addFieldNarrow}
          />
          <DefaultButton text="+ Link risk" onClick={handleAdd} disabled={!canAdd} />
        </div>
      )}
    </div>
  );
};

export default RiskLinkPicker;

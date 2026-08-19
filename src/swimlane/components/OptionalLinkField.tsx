import * as React from 'react';
import { TextField, DefaultButton, IconButton } from '@fluentui/react';
import styles from './OptionalLinkField.module.scss';

export interface IOptionalLinkFieldProps {
  label: string;
  addButtonText: string;
  placeholder?: string;
  value: string;
  onChange: (value: string) => void;
}

// Collapses to a small "+ Add..." button when empty rather than always
// showing an empty text box - most steps won't have a link filled in, and
// two permanently-visible blank fields on every single step reads as
// clutter for something only a minority will ever use. Once opened (by
// clicking Add, or because a value already exists on load), stays open
// for the rest of this edit session even if the text is cleared back to
// empty, so fixing a typo doesn't fight the user by collapsing the field
// out from under them mid-edit.
const OptionalLinkField: React.FC<IOptionalLinkFieldProps> = ({ label, addButtonText, placeholder, value, onChange }) => {
  const [open, setOpen] = React.useState(!!value);

  if (!open) {
    return <DefaultButton text={addButtonText} onClick={() => setOpen(true)} className={styles.addButton} />;
  }

  return (
    <div className={styles.row}>
      <TextField
        label={label}
        placeholder={placeholder}
        value={value}
        onChange={(_e, v) => onChange(v || '')}
        className={styles.field}
      />
      <IconButton
        iconProps={{ iconName: 'OpenInNewWindow' }}
        title="Open this link in a new tab"
        ariaLabel="Open this link in a new tab"
        disabled={!value.trim()}
        className={styles.iconButton}
        onClick={() => window.open(value.trim(), '_blank', 'noopener,noreferrer')}
      />
      <IconButton
        iconProps={{ iconName: 'Cancel' }}
        title={`Remove ${label.toLowerCase()}`}
        ariaLabel={`Remove ${label.toLowerCase()}`}
        className={styles.iconButton}
        onClick={() => { onChange(''); setOpen(false); }}
      />
    </div>
  );
};

export default OptionalLinkField;

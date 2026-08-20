import * as React from 'react';
import { IProcessStep } from '../models/IProcessStep';
import styles from './GlobalSearch.module.scss';

export interface IGlobalSearchProps {
  steps: IProcessStep[];
  onNavigate: (step: IProcessStep) => void;
}

const MAX_RESULTS = 8;

// Finding a specific step meant navigating Category -> Process Group ->
// Process ID by hand even if you already knew roughly what you were
// looking for - this jumps straight there from a description, step name,
// or Process Step ID typed anywhere in the app.
const GlobalSearch: React.FC<IGlobalSearchProps> = ({ steps, onNavigate }) => {
  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState(false);
  const containerRef = React.useRef<HTMLDivElement>(null);

  const results = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    return steps
      .filter(s =>
        s.actionDescription.toLowerCase().includes(q)
        || s.processStepName.toLowerCase().includes(q)
        || s.processStepId.toLowerCase().includes(q)
      )
      .slice(0, MAX_RESULTS);
  }, [steps, query]);

  React.useEffect(() => {
    const onClickOutside = (e: MouseEvent): void => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', onClickOutside);
    return () => document.removeEventListener('mousedown', onClickOutside);
  }, []);

  const handleSelect = (step: IProcessStep): void => {
    onNavigate(step);
    setQuery('');
    setOpen(false);
  };

  return (
    <div className={styles.search} ref={containerRef}>
      <input
        className={styles.input}
        type="text"
        placeholder="Search process steps..."
        value={query}
        onChange={e => { setQuery(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
      />
      {open && query.trim() && (
        <div className={styles.results}>
          {results.length === 0 ? (
            <div className={styles.empty}>No matching steps</div>
          ) : (
            results.map(step => (
              <div className={styles.result} key={step.id} onClick={() => handleSelect(step)}>
                <span className={styles.resultId}>{step.processStepId}</span>
                <span className={styles.resultText}>{step.actionDescription}</span>
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
};

export default GlobalSearch;

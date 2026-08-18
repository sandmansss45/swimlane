import * as React from 'react';
import { ISwimlaneComment } from '../models/ISwimlaneComment';
import styles from './ImprovementsView.module.scss';

export interface IImprovementsViewProps {
  comments: ISwimlaneComment[];
}

const formatDate = (iso: string | undefined): string => (iso ? new Date(iso).toLocaleString() : '—');

// Read-only report of every comment ever left on a swimlane via "Leave a
// comment" in the toolbar - one place to see what changes/issues have
// been raised across the whole app, not just while looking at one
// swimlane at a time. See ISwimlaneComment for why this is append-only.
const ImprovementsView: React.FC<IImprovementsViewProps> = ({ comments }) => {
  const sorted = React.useMemo(
    () => comments.slice().sort((a, b) => (b.postedAt || '').localeCompare(a.postedAt || '')),
    [comments]
  );

  return (
    <div className={styles.card}>
      <h3 className={styles.title}>Improvements ({comments.length})</h3>
      {comments.length === 0 ? (
        <p className={styles.empty}>No comments have been left on any swimlane yet.</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Progress ID</th>
                <th>Region</th>
                <th>Author</th>
                <th>Comment</th>
                <th>Posted at</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map(c => (
                <tr key={c.id}>
                  <td>{c.progressId}</td>
                  <td>{c.region || <span className={styles.muted}>All</span>}</td>
                  <td>{c.author}</td>
                  <td className={styles.commentCell}>{c.comment}</td>
                  <td className={styles.muted}>{formatDate(c.postedAt)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default ImprovementsView;

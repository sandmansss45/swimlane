import * as React from 'react';
import { ISwimlaneComment } from '../models/ISwimlaneComment';
import styles from './SwimlaneComments.module.scss';

export interface ISwimlaneCommentsProps {
  comments: ISwimlaneComment[]; // already filtered to this one swimlane (progressId + region) by the caller
}

const formatDate = (iso: string): string => new Date(iso).toLocaleString();

// Read-only feed for THIS swimlane specifically - posting still happens
// via "Leave a comment" in the toolbar (see SwimlaneStudio.tsx), this is
// just where the result shows up without having to leave the page and
// find it on the separate cross-swimlane Improvements tab. That tab still
// exists for the "everything, everywhere" view; this is the "what's been
// said about the thing I'm looking at right now" view.
const SwimlaneComments: React.FC<ISwimlaneCommentsProps> = ({ comments }) => {
  const sorted = React.useMemo(
    () => comments.slice().sort((a, b) => (b.postedAt || '').localeCompare(a.postedAt || '')),
    [comments]
  );

  return (
    <div className={styles.card}>
      <h3 className={styles.title}>Comments ({comments.length})</h3>
      {comments.length === 0 ? (
        <p className={styles.empty}>No comments yet on this swimlane - use "Leave a comment" above to add one.</p>
      ) : (
        <div className={styles.list}>
          {sorted.map(c => (
            <div className={styles.comment} key={c.id}>
              <div className={styles.commentHeader}>
                <span className={styles.author}>{c.author}</span>
                <span className={styles.postedAt}>{formatDate(c.postedAt)}</span>
              </div>
              <p className={styles.text}>{c.comment}</p>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default SwimlaneComments;

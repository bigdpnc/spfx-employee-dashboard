import * as React from 'react';
import styles from './QuickLinks.module.scss';
import { IDashboardNotice, IQuickLink } from './IEmployeeDashboardProps';
import { ArrowIcon, DashboardIcon } from './icons';

export interface IQuickLinksProps {
  links: IQuickLink[];
  /** Columns at full width. Narrow containers automatically drop to 2, then 1. */
  columns: number;
  loading?: boolean;
  notice?: IDashboardNotice;
  manageUrl?: string;
}

export const MIN_COLUMNS = 1;
export const MAX_COLUMNS = 6;

export function clampColumns(value: number): number {
  if (!isFinite(value)) return 4;
  return Math.min(MAX_COLUMNS, Math.max(MIN_COLUMNS, Math.round(value)));
}

export const QuickLinks: React.FC<IQuickLinksProps> = ({ links, columns, loading = false, notice, manageUrl }) => {
  const cols = clampColumns(columns);

  // CSS custom properties drive the grid, so changing columns never re-mounts cards.
  const gridStyle = {
    '--ql-cols': cols,
    '--ql-cols-md': Math.min(cols, 2),
  } as React.CSSProperties;

  return (
    <section className={styles.section} aria-labelledby="ed-quicklinks-heading">
      <div className={styles.heading}>
        <h3 id="ed-quicklinks-heading">Quick Links</h3>
        <span className={styles.headingMeta}>
          {!loading && !notice ? (
            <span className={styles.count}>
              {links.length} {links.length === 1 ? 'link' : 'links'}
            </span>
          ) : null}
          {manageUrl ? (
            <a className={styles.manage} href={manageUrl} target="_blank" rel="noopener noreferrer">
              Manage links
              <span className={styles.srOnly}> (opens the Quick Links list in a new tab)</span>
            </a>
          ) : null}
        </span>
      </div>

      {loading ? (
        <ul className={styles.grid} style={gridStyle} aria-busy="true" aria-label="Loading quick links">
          {[0, 1, 2, 3].slice(0, Math.max(cols, 2)).map((i) => (
            <li key={i}>
              <span className={`${styles.card} ${styles.skeletonCard}`} aria-hidden="true">
                <span className={`${styles.iconTile} ${styles.skeleton}`} />
                <span className={styles.body}>
                  <span className={`${styles.skeleton} ${styles.skeletonLine}`} />
                  <span className={`${styles.skeleton} ${styles.skeletonLine} ${styles.skeletonShort}`} />
                </span>
              </span>
            </li>
          ))}
        </ul>
      ) : notice ? (
        <div className={`${styles.notice} ${notice.tone === 'error' ? styles.noticeError : ''}`} role={notice.tone === 'error' ? 'alert' : undefined}>
          <p>{notice.message}</p>
          {notice.actionLabel && notice.onAction ? (
            <button type="button" className={styles.noticeButton} onClick={notice.onAction} disabled={notice.busy}>
              {notice.busy ? notice.busyLabel || 'Working…' : notice.actionLabel}
            </button>
          ) : null}
        </div>
      ) : links.length === 0 ? (
        <p className={styles.empty}>
          No links to show yet.{manageUrl ? ' Add one with Manage links, and set Show to Yes.' : ''}
        </p>
      ) : (
        <ul className={styles.grid} style={gridStyle}>
          {links.map((link) => {
            const newTab = link.openInNewTab !== false;
            return (
              <li key={link.id}>
                <a
                  className={styles.card}
                  href={link.url}
                  target={newTab ? '_blank' : undefined}
                  rel={newTab ? 'noopener noreferrer' : undefined}
                  style={{ '--ql-accent': link.accent || 'var(--ed-primary)' } as React.CSSProperties}
                >
                  <span className={styles.iconTile} aria-hidden="true">
                    <DashboardIcon name={link.icon} size={22} />
                  </span>
                  <span className={styles.body}>
                    <span className={styles.title}>{link.title}</span>
                    {link.description ? <span className={styles.description}>{link.description}</span> : null}
                  </span>
                  <ArrowIcon className={styles.arrow} size={16} />
                  {newTab ? <span className={styles.srOnly}> (opens in a new tab)</span> : null}
                </a>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

import * as React from 'react';
import styles from './QuickLinks.module.scss';
import { IQuickLink } from './IEmployeeDashboardProps';
import { ArrowIcon, DashboardIcon } from './icons';

export interface IQuickLinksProps {
  links: IQuickLink[];
  /** Columns at full width. Narrow containers automatically drop to 2, then 1. */
  columns: number;
}

export const MIN_COLUMNS = 1;
export const MAX_COLUMNS = 6;

export function clampColumns(value: number): number {
  if (!isFinite(value)) return 4;
  return Math.min(MAX_COLUMNS, Math.max(MIN_COLUMNS, Math.round(value)));
}

export const QuickLinks: React.FC<IQuickLinksProps> = ({ links, columns }) => {
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
        <span className={styles.count}>{links.length} apps</span>
      </div>

      {links.length === 0 ? (
        <p className={styles.empty}>No quick links configured yet.</p>
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

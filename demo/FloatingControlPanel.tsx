import * as React from 'react';
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneFieldType } from './mockPropertyPane';

export type CanvasLayout = 'full' | 'one' | 'twoThirds' | 'oneThird';

/** Demo-only: simulate what the To Do connection can do in SharePoint. */
export type TaskState = 'loaded' | 'loading' | 'error' | 'saveFails';

export const TASK_STATES: { key: TaskState; label: string }[] = [
  { key: 'loaded', label: 'Loaded' },
  { key: 'loading', label: 'Loading' },
  { key: 'error', label: 'Load error' },
  { key: 'saveFails', label: 'Save fails' },
];

export const CANVAS_LAYOUTS: { key: CanvasLayout; label: string; width: string }[] = [
  { key: 'full', label: 'Full width', width: '100%' },
  { key: 'one', label: 'One column', width: '1200px' },
  { key: 'twoThirds', label: '2/3 column', width: '800px' },
  { key: 'oneThird', label: '1/3 column', width: '380px' },
];

interface IFloatingControlPanelProps {
  config: IPropertyPaneConfiguration;
  properties: Record<string, unknown>;
  /** Same signature as `BaseClientSideWebPart.onPropertyPaneFieldChanged`. */
  onPropertyPaneFieldChanged: (propertyPath: string, oldValue: unknown, newValue: unknown) => void;
  onReset: () => void;
  canvasLayout: CanvasLayout;
  onCanvasLayoutChange: (layout: CanvasLayout) => void;
  taskState: TaskState;
  onTaskStateChange: (state: TaskState) => void;
}

const GearIcon: React.FC = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" aria-hidden="true">
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z" />
  </svg>
);

const CloseIcon: React.FC = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
    <path d="M6 6l12 12M18 6 6 18" />
  </svg>
);

const Field: React.FC<{
  field: IPropertyPaneField;
  value: unknown;
  onChange: (newValue: unknown) => void;
}> = ({ field, value, onChange }) => {
  const id = `pp-${field.targetProperty}`;

  if (field.type === PropertyPaneFieldType.Toggle) {
    const checked = !!value;
    const { label, onText = 'On', offText = 'Off' } = field.properties;
    return (
      <div className="pp-field pp-toggle-row">
        <label htmlFor={id}>{label}</label>
        <span className="pp-toggle-wrap">
          <span className="pp-toggle-text">{checked ? onText : offText}</span>
          <button
            id={id}
            type="button"
            role="switch"
            aria-checked={checked}
            className={`pp-toggle ${checked ? 'is-on' : ''}`}
            onClick={() => onChange(!checked)}
          >
            <span className="pp-toggle-thumb" />
          </button>
        </span>
      </div>
    );
  }

  const { label, min, max, step = 1, showValue } = field.properties;
  const num = typeof value === 'number' ? value : min;
  return (
    <div className="pp-field">
      <div className="pp-slider-head">
        <label htmlFor={id}>{label}</label>
        {showValue ? <output htmlFor={id}>{num}</output> : null}
      </div>
      <input
        id={id}
        type="range"
        min={min}
        max={max}
        step={step}
        value={num}
        onChange={(e) => onChange(Number(e.target.value))}
        style={{ '--pp-fill': `${((num - min) / (max - min)) * 100}%` } as React.CSSProperties}
      />
      <div className="pp-slider-scale" aria-hidden="true">
        <span>{min}</span>
        <span>{max}</span>
      </div>
    </div>
  );
};

/**
 * A floating stand-in for the SharePoint property pane. It renders whatever
 * `getPropertyPaneConfiguration()` returns, so new fields show up automatically.
 */
export const FloatingControlPanel: React.FC<IFloatingControlPanelProps> = ({
  config,
  properties,
  onPropertyPaneFieldChanged,
  onReset,
  canvasLayout,
  onCanvasLayoutChange,
  taskState,
  onTaskStateChange,
}) => {
  // Start open on desktop; on phones start closed so the dashboard is visible first.
  const [open, setOpen] = React.useState(() => window.innerWidth > 700);
  const toggleRef = React.useRef<HTMLButtonElement>(null);

  React.useEffect(() => {
    if (!open) return undefined;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        setOpen(false);
        if (toggleRef.current) toggleRef.current.focus();
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open]);

  return (
    <div className={`pp-root ${open ? 'is-open' : ''}`}>
      {open ? (
        <aside className="pp-panel" aria-label="Web part properties (demo)">
          <div className="pp-header">
            <div>
              <p className="pp-eyebrow">Property pane · mock</p>
              <h2>Employee Dashboard</h2>
            </div>
            <button type="button" className="pp-icon-btn" onClick={() => setOpen(false)} aria-label="Close property pane">
              <CloseIcon />
            </button>
          </div>

          <div className="pp-body">
            {config.pages.map((page, pi) => (
              <div key={pi}>
                <p className="pp-description">{page.header.description}</p>
                {page.groups.map((group) => (
                  <fieldset key={group.groupName} className="pp-group">
                    <legend>{group.groupName}</legend>
                    {group.groupFields.map((field) => (
                      <Field
                        key={field.targetProperty}
                        field={field}
                        value={properties[field.targetProperty]}
                        onChange={(newValue) =>
                          onPropertyPaneFieldChanged(field.targetProperty, properties[field.targetProperty], newValue)
                        }
                      />
                    ))}
                  </fieldset>
                ))}
              </div>
            ))}

            <fieldset className="pp-group pp-group-demo">
              <legend>Demo only · section width</legend>
              <div className="pp-segment" role="radiogroup" aria-label="Simulated SharePoint section width">
                {CANVAS_LAYOUTS.map((l) => (
                  <button
                    key={l.key}
                    type="button"
                    role="radio"
                    aria-checked={canvasLayout === l.key}
                    className={canvasLayout === l.key ? 'is-active' : ''}
                    onClick={() => onCanvasLayoutChange(l.key)}
                  >
                    {l.label}
                  </button>
                ))}
              </div>
              <p className="pp-hint">Shrinks the web part to test the responsive Quick Links grid.</p>
            </fieldset>

            <fieldset className="pp-group pp-group-demo">
              <legend>Demo only · To Do connection</legend>
              <div className="pp-segment" role="radiogroup" aria-label="Simulated To Do connection state">
                {TASK_STATES.map((t) => (
                  <button
                    key={t.key}
                    type="button"
                    role="radio"
                    aria-checked={taskState === t.key}
                    className={taskState === t.key ? 'is-active' : ''}
                    onClick={() => onTaskStateChange(t.key)}
                  >
                    {t.label}
                  </button>
                ))}
              </div>
              <p className="pp-hint">Shows how My Tasks behaves while loading, if loading fails, or if ticking a task can't be saved.</p>
            </fieldset>
          </div>

          <div className="pp-footer">
            <button type="button" className="pp-link-btn" onClick={onReset}>
              Reset to defaults
            </button>
            <span className="pp-hint">Esc to close</span>
          </div>
        </aside>
      ) : null}

      <button
        ref={toggleRef}
        type="button"
        className="pp-fab"
        aria-expanded={open}
        aria-label={open ? 'Hide property pane' : 'Show property pane'}
        onClick={() => setOpen((o) => !o)}
      >
        <GearIcon />
      </button>
    </div>
  );
};

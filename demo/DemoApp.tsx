import * as React from 'react';
import { EmployeeDashboard, ITask, clampColumns } from '../src/components';
import { CANVAS_LAYOUTS, CanvasLayout, FloatingControlPanel } from './FloatingControlPanel';
import { DEFAULT_WEB_PART_PROPERTIES, IEmployeeDashboardWebPartProps, getPropertyPaneConfiguration } from './mockPropertyPane';
import { mockTasks, mockUser } from './mockContext';

const STORAGE_KEY = 'ed-demo-properties';

interface IDemoState extends IEmployeeDashboardWebPartProps {
  canvasLayout: CanvasLayout;
}

const DEFAULT_STATE: IDemoState = { ...DEFAULT_WEB_PART_PROPERTIES, canvasLayout: 'one' };

// localStorage is a convenience only; the demo works without it (private mode, etc.).
function loadState(): IDemoState {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...DEFAULT_STATE, ...JSON.parse(raw) };
  } catch {
    /* ignore */
  }
  return DEFAULT_STATE;
}

function saveState(state: IDemoState): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    /* ignore */
  }
}

export const DemoApp: React.FC = () => {
  const [state, setState] = React.useState<IDemoState>(loadState);
  const paneConfig = React.useMemo(getPropertyPaneConfiguration, []);

  React.useEffect(() => saveState(state), [state]);

  // Mirrors BaseClientSideWebPart: update the property bag, then re-render.
  const onPropertyPaneFieldChanged = (propertyPath: string, _oldValue: unknown, newValue: unknown): void => {
    setState((s) => ({ ...s, [propertyPath]: newValue }));
  };

  const onTaskToggle = React.useCallback((task: ITask) => {
    // In SharePoint this is where you'd PATCH Planner / To Do / a list item.
    console.info('[demo] task toggled', task);
  }, []);

  const width = CANVAS_LAYOUTS.filter((l) => l.key === state.canvasLayout)[0].width;

  return (
    <div className={`demo-page ${state.isDarkTheme ? 'demo-dark' : ''}`}>
      <div className="demo-suitebar">
        <span className="demo-waffle" aria-hidden="true" />
        <span className="demo-site">Intranet</span>
        <span className="demo-badge">Local workbench</span>
      </div>

      <main className="demo-canvas" style={{ maxWidth: width }}>
        <EmployeeDashboard
          user={mockUser}
          isDarkTheme={state.isDarkTheme}
          quickLinksColumns={clampColumns(state.quickLinksColumns)}
          tasks={mockTasks}
          onTaskToggle={onTaskToggle}
        />
      </main>

      <FloatingControlPanel
        config={paneConfig}
        properties={state as unknown as Record<string, unknown>}
        onPropertyPaneFieldChanged={onPropertyPaneFieldChanged}
        onReset={() => setState(DEFAULT_STATE)}
        canvasLayout={state.canvasLayout}
        onCanvasLayoutChange={(canvasLayout) => setState((s) => ({ ...s, canvasLayout }))}
      />
    </div>
  );
};

import * as React from 'react';
import { EmployeeDashboard, ITask, SAMPLE_TASKS, clampColumns } from '../src/components';
import { CANVAS_LAYOUTS, CanvasLayout, FloatingControlPanel, TaskState } from './FloatingControlPanel';
import { DEFAULT_WEB_PART_PROPERTIES, IEmployeeDashboardWebPartProps, getPropertyPaneConfiguration } from './mockPropertyPane';
import { mockUser } from './mockContext';

const STORAGE_KEY = 'ed-demo-properties';

interface IDemoState extends IEmployeeDashboardWebPartProps {
  canvasLayout: CanvasLayout;
}

const DEFAULT_STATE: IDemoState = { ...DEFAULT_WEB_PART_PROPERTIES, canvasLayout: 'one' };

const LOAD_ERROR =
  "We couldn't load your tasks from Microsoft To Do. If this keeps happening, ask your SharePoint admin to approve the Tasks.ReadWrite permission under API access.";

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

  // Not persisted: the demo always opens with real data showing.
  const [taskState, setTaskState] = React.useState<TaskState>('loaded');

  const onTaskToggle = React.useCallback(
    (task: ITask): Promise<void> => {
      // In SharePoint this PATCHes the task in Microsoft To Do via Graph.
      console.info('[demo] task toggled', task);
      return new Promise((resolve, reject) => {
        window.setTimeout(() => (taskState === 'saveFails' ? reject(new Error('Simulated save failure')) : resolve()), 600);
      });
    },
    [taskState],
  );

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
          tasks={SAMPLE_TASKS}
          tasksLoading={taskState === 'loading'}
          tasksError={taskState === 'error' ? LOAD_ERROR : undefined}
          onRetryTasks={() => setTaskState('loaded')}
          tasksSourceName="Microsoft To Do"
          tasksSourceUrl="https://to-do.office.com/tasks/"
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
        taskState={taskState}
        onTaskStateChange={setTaskState}
      />
    </div>
  );
};

# spfx-employee-dashboard

A **Personal Employee Dashboard** web part for SharePoint Framework (SPFx), plus a standalone `/demo` workbench for building and testing it without a tenant.

## What's inside

```
src/components/              ← the web part's React UI (SPFx-ready: React 17, *.module.scss)
  EmployeeDashboard.tsx        root component + theme tokens (light / dark)
  GreetingHeader.tsx           time-of-day greeting, avatar, date, task summary
  QuickLinks.tsx               responsive card grid with a configurable column count
  MyTasks.tsx                  Pending / Completed tabs (accessible tabs pattern)
  IEmployeeDashboardProps.ts   props contract (user context + web part properties)
  defaultQuickLinks.ts         default links (Email, HR Portal, Expense Reports, …)
  icons.tsx, dateUtils.ts      inline SVG icons and date helpers (no extra deps)

demo/                        ← local workbench (Vite)
  mockPropertyPane.ts          SPFx-shaped PropertyPaneToggle / PropertyPaneSlider + getPropertyPaneConfiguration()
  mockContext.ts               mock pageContext.user and sample tasks
  FloatingControlPanel.tsx     floating property pane that renders the mock config
  DemoApp.tsx                  wires the property bag to the dashboard
```

## Run the demo

```bash
npm install
npm run dev          # http://localhost:5173
```

Use the floating ⚙ button (bottom right) to:

- toggle **Dark mode** (`isDarkTheme`)
- change **Visible columns** in Quick Links (`quickLinksColumns`, 1–6)
- *(demo only)* change the simulated SharePoint section width to check the responsive grid

Settings are saved in `localStorage`, so they survive a reload.

Other scripts: `npm run typecheck`, `npm run build:demo`.

## Responsive behaviour

The dashboard uses **container queries** against its own width, not the viewport, so it adapts to SharePoint's full-width, one-, two- and three-column sections:

| Web part width | Quick Links columns           |
| -------------- | ----------------------------- |
| > 720px        | `quickLinksColumns` (1–6)     |
| 421–720px      | `min(quickLinksColumns, 2)`   |
| ≤ 420px        | 1                             |

## Using it in the SPFx web part

```ts
public render(): void {
  const element = React.createElement(EmployeeDashboard, {
    user: {
      displayName: this.context.pageContext.user.displayName,
      email: this.context.pageContext.user.email,
    },
    isDarkTheme: this.properties.isDarkTheme,
    quickLinksColumns: this.properties.quickLinksColumns,
    tasks: [], // e.g. from Microsoft Graph: /me/todo or /me/planner/tasks
  });
  ReactDom.render(element, this.domElement);
}
```

`getPropertyPaneConfiguration()` in `demo/mockPropertyPane.ts` already has the shape the web part needs. Import the real `PropertyPaneToggle` / `PropertyPaneSlider` from `@microsoft/sp-property-pane` and it can be used as-is.

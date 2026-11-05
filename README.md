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

Other scripts: `npm run typecheck:demo`, `npm run build:demo`.

## Responsive behaviour

The dashboard uses **container queries** against its own width, not the viewport, so it adapts to SharePoint's full-width, one-, two- and three-column sections:

| Web part width | Quick Links columns           |
| -------------- | ----------------------------- |
| > 720px        | `quickLinksColumns` (1–6)     |
| 421–720px      | `min(quickLinksColumns, 2)`   |
| ≤ 420px        | 1                             |

## Build the SharePoint package

Requires Node 22 (SPFx 1.23.2).

```bash
npm install
npm run build        # → sharepoint/solution/spfx-employee-dashboard.sppkg
```

## Deploy to your tenant

You need to be a SharePoint admin, or have rights to the tenant App Catalog.

1. Open the App Catalog: **SharePoint admin center → More features → Apps → Open**, or go to
   `https://<your-tenant>.sharepoint.com/sites/appcatalog`.
2. Upload `spfx-employee-dashboard.sppkg`.
3. In the dialog, choose **Enable this app and add it to all sites** (or **Only enable this app** to
   add it site by site under *Site contents → New → App*), then **Enable**.
4. Edit any modern page, click **+**, search for **Employee Dashboard**, add it and **Republish**.
5. Click the web part's ✏️ edit icon to open the property pane: Dark mode, Visible columns, Show sample tasks.

To ship an update, bump `version` in `config/package-solution.json`, rebuild, and upload the new
`.sppkg` over the old one.

## Project layout for SPFx

- `src/webparts/employeeDashboard/EmployeeDashboardWebPart.ts` reads the signed-in user from
  `this.context.pageContext.user`, loads their profile photo, and defines the real property pane.
- `src/components/` is shared by the web part and the demo.
- `src/components/sampleTasks.ts` holds the example tasks shown until a real task source
  (Microsoft To Do / Planner via Graph) is connected. Turn them off with **Show sample tasks**.

declare interface IEmployeeDashboardWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  DarkModeLabel: string;
  QuickLinksGroupName: string;
  ColumnsLabel: string;
  TasksGroupName: string;
  TaskSourceLabel: string;
  TodoSourceName: string;
  SampleSourceName: string;
  TodoLoadError: string;
  On: string;
  Off: string;
}

declare module 'EmployeeDashboardWebPartStrings' {
  const strings: IEmployeeDashboardWebPartStrings;
  export = strings;
}

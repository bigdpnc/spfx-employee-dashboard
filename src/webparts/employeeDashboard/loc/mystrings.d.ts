declare interface IEmployeeDashboardWebPartStrings {
  PropertyPaneDescription: string;
  AppearanceGroupName: string;
  DarkModeLabel: string;
  QuickLinksGroupName: string;
  ColumnsLabel: string;
  TasksGroupName: string;
  SampleTasksLabel: string;
  On: string;
  Off: string;
}

declare module 'EmployeeDashboardWebPartStrings' {
  const strings: IEmployeeDashboardWebPartStrings;
  export = strings;
}

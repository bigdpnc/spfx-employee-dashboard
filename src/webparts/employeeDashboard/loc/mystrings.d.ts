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
  TodoPermissionError: string;
  TodoNoMailboxError: string;
  LinksListLabel: string;
  LinksListDescription: string;
  LinksListMissingOwner: string;
  LinksListMissingVisitor: string;
  LinksLoadError: string;
  CreateList: string;
  CreatingList: string;
  CreateListError: string;
  TryAgain: string;
  On: string;
  Off: string;
}

declare module 'EmployeeDashboardWebPartStrings' {
  const strings: IEmployeeDashboardWebPartStrings;
  export = strings;
}

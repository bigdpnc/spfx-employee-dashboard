/**
 * Mock of the SPFx property pane API (`@microsoft/sp-property-pane`).
 *
 * The factories and `getPropertyPaneConfiguration()` below have the same shape the
 * real web part class will use, so this file can be copied into
 * `EmployeeDashboardWebPart.ts` almost verbatim. Swap the imports for the real
 * `PropertyPaneToggle` / `PropertyPaneSlider` and delete the local factories.
 */

// ---- Minimal re-creation of the SPFx types ---------------------------------

export enum PropertyPaneFieldType {
  Toggle = 'Toggle',
  Slider = 'Slider',
}

export interface IPropertyPaneToggleProps {
  label: string;
  onText?: string;
  offText?: string;
}

export interface IPropertyPaneSliderProps {
  label: string;
  min: number;
  max: number;
  step?: number;
  showValue?: boolean;
}

export type IPropertyPaneField =
  | { type: PropertyPaneFieldType.Toggle; targetProperty: string; properties: IPropertyPaneToggleProps }
  | { type: PropertyPaneFieldType.Slider; targetProperty: string; properties: IPropertyPaneSliderProps };

export interface IPropertyPaneGroup {
  groupName: string;
  groupFields: IPropertyPaneField[];
}

export interface IPropertyPanePage {
  header: { description: string };
  groups: IPropertyPaneGroup[];
}

export interface IPropertyPaneConfiguration {
  pages: IPropertyPanePage[];
}

export function PropertyPaneToggle(targetProperty: string, properties: IPropertyPaneToggleProps): IPropertyPaneField {
  return { type: PropertyPaneFieldType.Toggle, targetProperty, properties };
}

export function PropertyPaneSlider(targetProperty: string, properties: IPropertyPaneSliderProps): IPropertyPaneField {
  return { type: PropertyPaneFieldType.Slider, targetProperty, properties };
}

// ---- The web part's property bag and pane config ---------------------------

/** What will live in `this.properties` on the real web part (and in its manifest's `preconfiguredEntries`). */
export interface IEmployeeDashboardWebPartProps {
  isDarkTheme: boolean;
  quickLinksColumns: number;
}

export const DEFAULT_WEB_PART_PROPERTIES: IEmployeeDashboardWebPartProps = {
  isDarkTheme: false,
  quickLinksColumns: 4,
};

export function getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: { description: 'Personalise how the Employee Dashboard looks on this page.' },
        groups: [
          {
            groupName: 'Appearance',
            groupFields: [
              PropertyPaneToggle('isDarkTheme', {
                label: 'Dark mode',
                onText: 'On',
                offText: 'Off',
              }),
            ],
          },
          {
            groupName: 'Quick Links',
            groupFields: [
              PropertyPaneSlider('quickLinksColumns', {
                label: 'Visible columns',
                min: 1,
                max: 6,
                step: 1,
                showValue: true,
              }),
            ],
          },
        ],
      },
    ],
  };
}

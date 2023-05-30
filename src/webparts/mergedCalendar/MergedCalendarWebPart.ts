import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {SPHttpClient} from '@microsoft/sp-http';

import * as strings from 'MergedCalendarWebPartStrings';
import MergedCalendar from './components/MergedCalendar';
import { IMergedCalendarProps } from './components/IMergedCalendarProps';

export interface IMergedCalendarWebPartProps {
  description: string;  
  showWeekends: boolean;
  calSettingsList: string;
  roomsList: string;
  periodsList: string;
  guidelinesList: string;
  isPeriods: boolean;

  isListView: boolean;
  listViewType: string;
  listViewNavBtns: boolean;
  listViewLegend: boolean;
  listViewErrors: boolean;
  listViewMonthTitle: boolean;
  listViewViews: boolean;
  listViewHeight: number;
  listViewTitle: string;
  listViewRoomsFilter: boolean;
  listViewRoomsLegend: boolean;
}

export default class MergedCalendarWebPart extends BaseClientSideWebPart<IMergedCalendarWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMergedCalendarProps> = React.createElement(
      MergedCalendar,
      {
        description: this.properties.description,
        showWeekends: this.properties.showWeekends,
        context: this.context,
        calSettingsList: this.properties.calSettingsList,
        roomsList: this.properties.roomsList,
        periodsList: this.properties.periodsList,
        guidelinesList: this.properties.guidelinesList,
        dpdOptions : [
          { key: 'E1Day', text: '1 Day Cycle' },
          { key: 'E2Day', text: '2 Day Cycle' },
          { key: 'E3Day', text: '3 Day Cycle' },
          { key: 'E4Day', text: '4 Day Cycle' },
          { key: 'E5Day', text: '5 Day Cycle' },
          { key: 'E6Day', text: '6 Day Cycle' },
          { key: 'E7Day', text: '7 Day Cycle' },
          { key: 'E8Day', text: '8 Day Cycle' },
          { key: 'E9Day', text: '9 Day Cycle' },
          { key: 'E10Day', text: '10 Day Cycle' },
        ],
        isPeriods: this.properties.isPeriods,

        isListView: this.properties.isListView,
        listViewType: this.properties.listViewType,
        listViewNavBtns: this.properties.listViewNavBtns,
        listViewLegend: this.properties.listViewLegend,
        listViewErrors: this.properties.listViewErrors,
        listViewMonthTitle: this.properties.listViewMonthTitle,
        listViewViews: this.properties.listViewViews,
        listViewHeight: this.properties.listViewHeight,
        listViewTitle: this.properties.listViewTitle,
        listViewRoomsFilter: this.properties.listViewRoomsFilter,
        listViewRoomsLegend: this.properties.listViewRoomsLegend
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // protected get disableReactivePropertyChanges(): boolean {
  //   return true;
  // }
  // private validateListName(value: string): string {
  //   if (value === null || value.trim().length === 0) {
  //     return 'Provide a list name';
  //   }
  //   if (value.length > 40) {
  //     return 'List name should not be longer than 40 characters';
  //   }
  //   return '';
  // }

  /* Loading Dpd with list names - Start */
  private lists: IPropertyPaneDropdownOption[];
  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {    
    let listsTitle : any = [];
    try {
      let response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=Title&$filter=BaseType eq 0 and BaseTemplate eq 100 and Hidden eq false`, SPHttpClient.configurations.v1);
      if (response.ok) {
        const results = await response.json();
        if(results){
          console.log('results', results);
          results.value.map((result: any)=>{
            listsTitle.push({
              key: result.Title,
              text: result.Title
            });
          });
          return listsTitle;
        }
      }
    } catch (error) {
      return error.message;
    }
  }
  protected onPropertyPaneConfigurationStart(): void {
    if (this.lists) {
      this.render();  
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');
    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();       
      });
  } 
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listName' && newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // re-render the web part as clearing the loading indicator removes the web part body
      this.render();      
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, oldValue);
    }
  }
  /* Loading Dpd with list names - End */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                /*PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),*/
                // PropertyPaneTextField('calSettingsList', {
                //   label: 'Calendar Settings List',
                //   onGetErrorMessage: this.validateListName.bind(this)
                // }),
                PropertyPaneDropdown('calSettingsList', {
                  label : 'Calendar Settings List',
                  options: this.lists,
                  selectedKey : 'CalendarSettings'
                }),
                PropertyPaneCheckbox('showWeekends', {
                  text: "Show Weekends"
                }),
                PropertyPaneDropdown('roomsList', {
                  label : 'Rooms List',
                  options: this.lists,
                  selectedKey : 'Rooms'
                }),
                PropertyPaneToggle('isPeriods', {
                  label: 'Use School Periods',
                  checked: this.properties.isPeriods !== undefined ? this.properties.isPeriods : true,
                  onText: 'Yes',
                  offText: 'No',
                }),
                PropertyPaneDropdown('periodsList', {
                  label : 'Periods List',
                  options: this.lists,
                  selectedKey : 'Periods'
                }),
                PropertyPaneDropdown('guidelinesList', {
                  label : 'Guidelines List',
                  options: this.lists,
                  selectedKey : 'Guidelines'
                }),
              ]
            },
            {
              groupName: 'Events View',
              groupFields: [
                PropertyPaneToggle('isListView', {
                  label: 'List View',
                  onText: 'On',
                  offText: 'Off',
                  checked : false
                }),
                PropertyPaneTextField('listViewTitle', {
                  label: 'Title',
                  value: this.properties.listViewTitle,
                }),
                PropertyPaneDropdown('listViewType', {
                  label: 'List View Type',                  
                  disabled: !this.properties.isListView,
                  options: [
                    {key: 'listDay', text: 'Day List'},
                    {key: 'listWeek', text: 'Week List'},
                    {key: 'listMonth', text: 'Month List'}
                  ],
                  selectedKey : this.properties.listViewType
                }),
                PropertyPaneLabel('listViewOptions', {
                  text: 'Header & Footer Options',
                }),
                PropertyPaneCheckbox('listViewMonthTitle', {
                  text: "Month Name",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewMonthTitle
                }),
                PropertyPaneCheckbox('listViewNavBtns', {
                  text: "Navigation Buttons (previous, next, today)",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewNavBtns
                }),
                PropertyPaneCheckbox('listViewLegend', {
                  text: "Legend",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewLegend
                  
                }),
                PropertyPaneCheckbox('listViewRoomsLegend', {
                  text: "Rooms Legend",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewRoomsLegend
                }),
                PropertyPaneCheckbox('listViewErrors', {
                  text: "Errors",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewErrors
                }),
                PropertyPaneCheckbox('listViewView', {
                  text: "Views",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewViews
                }),
                PropertyPaneCheckbox('listViewRoomsFilter', {
                  text: "Rooms Filter",
                  disabled: !this.properties.isListView,
                  checked: this.properties.listViewRoomsFilter
                }),
                PropertyPaneSlider('listViewHeight', {
                  label: 'Height',
                  min: 200,
                  max: 1000,
                  value: this.properties.listViewHeight,
                  disabled: !this.properties.isListView,
                  step : 10,
                  showValue: true,
                })
              ]
            },
          ]
        }
      ]
    };
  }
}

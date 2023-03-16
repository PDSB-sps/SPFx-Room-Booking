import { WebPartContext } from "@microsoft/sp-webpart-base";
import {IDropdownOption} from "@fluentui/react";

export interface IMergedCalendarProps {
  description: string;
  showWeekends: boolean;
  context: WebPartContext;  
  calSettingsList: string;
  dpdOptions: IDropdownOption[];
  roomsList: string;
  periodsList: string;
  guidelinesList: string;

  isPeriods: boolean;

  isListView: boolean;
  listViewType: any;
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

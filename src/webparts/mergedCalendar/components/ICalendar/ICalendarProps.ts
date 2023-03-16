import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICalendarProps{
    showWeekends: boolean;
    eventSources: {}[];
    openPanel: any;
    handleDateClick: (args:any) => void;
    context: WebPartContext;
    listGUID: string;
    passCurrentDate: (args:any) => void;

    isListView: boolean;
    listViewType: any;
    listViewNavBtns: boolean;
    listViewMonthTitle: boolean;
    listViewViews: boolean;
    listViewHeight: number;
}
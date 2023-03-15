import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRoomsManageProps{
    context: WebPartContext;
    roomsList: string;
    periodsList: string;
    guidelinesList: string;
    onRoomsManage: any;   
    iframeState: string; 
    openMultiBook: any;
    isPeriods: any;
}
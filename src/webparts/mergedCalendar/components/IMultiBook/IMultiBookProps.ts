import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMultiBookProps {
    formField: any;
    errorMsgField: any;
    onChangeFormField: any;

    schoolCategory: string;
    schoolNum: string;
    schoolCycleOptions: any;
    schoolCycleDayOptions: any;
    periodOptions: any;
    roomOptions: any;

    checkBookingClick: any;
    cancelMultiBook: any;

    context: WebPartContext;
    bookingsGridVisible: boolean;
}
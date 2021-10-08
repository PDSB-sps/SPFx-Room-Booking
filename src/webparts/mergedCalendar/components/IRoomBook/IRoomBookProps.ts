export interface IRoomBookProps{
    formField: any;
    errorMsgField: any;
    onChangeFormField: any;
    periodOptions: any;
    children: any;
    
    roomInfo: any;
    eventDetailsRoom: any;
    dismissPanelBook: any;
    bookFormMode: string;
    onNewBookingClick: any;
    onEditBookingClick: any;
    onDeleteBookingClick: any;
    onUpdateBookingClick: any;
    isCreator: boolean;
}
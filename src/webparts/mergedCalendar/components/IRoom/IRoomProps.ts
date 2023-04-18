import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRoomProps{
    onCheckAvailClick: any;
    onViewDetailsClick: any;
    onBookClick: any;
    roomInfo: any;
    onEditClick: any;
    onDeleteClick: any;
    context: WebPartContext;
}
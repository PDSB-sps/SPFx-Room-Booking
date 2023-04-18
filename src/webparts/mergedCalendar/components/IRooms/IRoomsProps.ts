import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRoomsProps{
    rooms: any;
    onCheckAvailClick: any;
    onViewDetailsClick: any;
    onBookClick: any;
    onEditClick: any;
    onDeleteClick: any;
    context: WebPartContext;
}
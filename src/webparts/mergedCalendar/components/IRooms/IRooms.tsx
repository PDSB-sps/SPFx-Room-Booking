import * as React from 'react';
import { IRoomsProps } from './IRoomsProps';
import IRoom from '../IRoom/IRoom';

export default function IRooms (props:IRoomsProps) {
    
    return(
        props.rooms.map((room: any)=>{
            return(
                <IRoom key = {room.Id}
                    roomInfo = {room}
                    onCheckAvailClick = {props.onCheckAvailClick(room.Id)}
                    onViewDetailsClick = {props.onViewDetailsClick(room)}
                    onBookClick = {props.onBookClick(room)}
                    onEditClick={props.onEditClick(room.Id)}
                    onDeleteClick={props.onDeleteClick(room.Id)}
                    context={props.context}
                />
            );
        })
    );
}
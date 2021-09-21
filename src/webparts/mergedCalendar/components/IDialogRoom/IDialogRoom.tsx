import * as React from 'react';
import {Dialog, DialogType, DialogFooter, DefaultButton} from '@fluentui/react';

import {IDialogRoomProps} from './IDialogRoomProps';
import IEventDetailsRoom from  '../IEventDetailsRoom/IEventDetailsRoom';

export default function IDialogRoom(props:IDialogRoomProps){
  
  const modelProps = {
    isBlocking: false,
    className: props.eventDetails ? 'modalColor'+props.eventDetails.Color :'',
  };
  const dialogContentProps = {
    type: DialogType.close,
    title: props.eventDetails.Title,
    subText: '',
  };

      return (
        <>
          <Dialog 
            hidden={props.hideDialog}
            onDismiss={props.toggleHideDialog}
            dialogContentProps={dialogContentProps}
            modalProps={modelProps}
            minWidth="25%">
            
            <IEventDetailsRoom 
                Title ={props.eventDetails.Title} 
                Start ={props.eventDetails.Start}
                End = {props.eventDetails.End}
                AllDay = {props.eventDetails.AllDay}
                Body = {props.eventDetails.Body}
                Location = {props.eventDetails.Location}       
                handleAddtoCal = {props.handleAddtoCal}   
                Room={props.eventDetails.Room}    
                Status={props.eventDetails.Status}
                Period={props.eventDetails.Period}  
                Color={props.eventDetails.Color}
            />
            
            <DialogFooter>
              <DefaultButton onClick={props.toggleHideDialog} text="Close" />
            </DialogFooter>
          </Dialog>
        </>
      );
}
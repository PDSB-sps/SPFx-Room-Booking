import * as React from 'react';
import {Dialog, DialogType, DialogFooter, DefaultButton} from '@fluentui/react';
import styles from '../MergedCalendar.module.scss';

import {IDialogProps} from './IDialogProps';
import IEventDetails from  '../IEventDetails/IEventDetails';

export default function IDialog(props:IDialogProps){
  
  const dlgTitleMkp = <span><span className={styles.evTitleDlg} style={{backgroundColor: props.eventDetails.Color }}></span>{props.eventDetails.Calendar}</span> ;

  const modelProps = {
    isBlocking: false,
  };
  const dialogContentProps = {
    type: DialogType.close,
    title: dlgTitleMkp,
    subText: '',
  };
  
  return (
    <>
      <Dialog
        hidden={props.hideDialog}
        onDismiss={props.toggleHideDialog}
        dialogContentProps={dialogContentProps}
        modalProps={modelProps}
        minWidth="35%" >

        <IEventDetails 
            Title ={props.eventDetails.Title} 
            Start ={props.eventDetails.Start}
            End = {props.eventDetails.End}
            AllDay = {props.eventDetails.AllDay}
            Body = {props.eventDetails.Body}
            Location = {props.eventDetails.Location}       
            handleAddtoCal = {props.handleAddtoCal}         
        />
        <DialogFooter>
          <DefaultButton onClick={props.toggleHideDialog} text="Close" />
        </DialogFooter>
      </Dialog>
    </>
  );
}
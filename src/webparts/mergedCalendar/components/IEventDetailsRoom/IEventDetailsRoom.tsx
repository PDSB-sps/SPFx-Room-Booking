import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import roomStyles from '../Room.module.scss';
import {IEventDetailsRoomProps} from './IEventDetailsRoomProps';
// import {DefaultButton} from '@fluentui/react';

export default function IEventDetailsRoom (props: IEventDetailsRoomProps){

    return(
        <div className={styles.eventDetails}>

            {/* <h1 style={{backgroundColor: props.Color}} className={props.Color ? roomStyles.dialogHDR : roomStyles.dialogHDRNoColor}>{props.Title}</h1> */}

            {props.Room &&
                <h3>
                    <span className={roomStyles.roomBullet} style={{backgroundColor: props.Color}}></span>&nbsp;
                    {props.Room}
                </h3>
            }

            {/* <div className={styles.evFld}>
                <label className={styles.evLbl}>Title</label>
                <div className={styles.evIp}>{props.Title}</div>
            </div> */}
            <div className={styles.evFld}>
                <label className={styles.evLbl}>Start time</label>
                <div className={styles.evIp}>
                    {props.Start}
                    {props.AllDay &&
                        <i> (All Day Event)</i>
                    }
                </div>
            </div>
            {props.End && 
                <div className={styles.evFld}>
                    <label className={styles.evLbl}>End Time</label>
                    <div className={styles.evIp}>{props.End}</div>
                </div>
            }
            {props.Period && 
                <div className={styles.evFld}>
                    <label className={styles.evLbl}>Period</label>
                    <div className={styles.evIp}>{props.Period}</div>
                </div>
            }
            {props.Status && 
                <div className={styles.evFld}>
                    <label className={styles.evLbl}>Status</label>
                    <div className={styles.evIp}>{props.Status}</div>
                </div>
            }
            {props.Location &&
                <div className={styles.evFld}>
                    <label className={styles.evLbl}>Location</label>
                    <div className={styles.evIp}>{props.Location}</div>
                </div>
            }
            {props.Body &&
                <div className={styles.evFld}>
                    <label className={styles.evLbl +" "+ styles.nonFL}>Description</label>
                    <div><p dangerouslySetInnerHTML={{__html: props.Body}}></p></div>
                </div>
            }

            {/* <DefaultButton onClick={props.handleAddtoCal}>Add dummy event to my Calendar</DefaultButton> */}

        </div>
    );
}
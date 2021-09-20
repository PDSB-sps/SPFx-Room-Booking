import * as React from 'react';
import '../ILegend/ILegend.scss';
import styles from '../MergedCalendar.module.scss';
import roomStyles from '../Room.module.scss';
import { ILegendRoomsProps } from './ILegendRoomsProps';

export default function ILegendRooms(props:ILegendRoomsProps){
    return(
        <div className={styles.calendarLegend}>
            <ul>
            {
                props.calSettings.map((value:any)=>{
                    return(
                        <React.Fragment>
                            {value.ShowCal && value.CalType === 'Room' &&
                                props.rooms.map((room: any)=>{
                                    return(
                                        <li key={value.Id} className={roomStyles.roomLegendItem}>
                                            <a>
                                                <span style={{backgroundColor: room.Colour}} className={styles.legendBullet}></span>
                                                <span className={styles.legendText}>{room.Title}</span>
                                            </a>
                                        </li>
                                    );
                                })
                            }
                        </React.Fragment>
                    );
                })
            }
            </ul>
        </div>
    );
}






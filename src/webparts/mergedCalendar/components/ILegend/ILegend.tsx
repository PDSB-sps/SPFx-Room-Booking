import * as React from 'react';
import './ILegend.scss';
import styles from '../MergedCalendar.module.scss';
import roomStyles from '../Room.module.scss';
import {Checkbox, Link} from '@fluentui/react';
import { ILegendProps } from './ILegendProps';
import { initializeIcons } from '@uifabric/icons';

export default function ILegend(props:ILegendProps){

    initializeIcons();
    const _renderLabelWithLink = (calTitle: string, hrefVal: string) => {
        return (
            <Link href={hrefVal} target="_blank" underline>
                {calTitle}
            </Link>
        );
    };
    return(
        <div className={styles.calendarLegend}>
            <ul>
            {
                props.calSettings.map((value:any)=>{
                    return(
                        <li key={value.Id}>
                            {value.ShowCal && value.CalType !== 'Room' &&
                                <Checkbox 
                                    className={'chkboxLegend chkbox_'+value.BgColor}
                                    label={value.Title} 
                                    defaultChecked={false}
                                    // checked={props.legendChked}
                                    onChange={props.onLegendChkChange(value.Id)} 
                                    onRenderLabel={() => _renderLabelWithLink(value.Title, value.LegendURL)}
                                />
                            }
                        </li>
                    );
                })
            }
            </ul>
        </div>
    );
}






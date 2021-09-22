import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import { IPreloaderProps } from './IPreloaderProps';

import {Spinner, SpinnerSize, Overlay} from '@fluentui/react';

export default function IPreloader (props:IPreloaderProps) {

    return(
        <React.Fragment>
            {props.isDataLoading &&
                <div className={styles.preloader}>
                    <Spinner className={styles.preloaderTxt} size={SpinnerSize.medium} label={props.text} ariaLive="assertive" labelPosition="right" />
                    <Overlay className={styles.preloaderOverlay}></Overlay>
                </div>
            }
        </React.Fragment>
    );
}
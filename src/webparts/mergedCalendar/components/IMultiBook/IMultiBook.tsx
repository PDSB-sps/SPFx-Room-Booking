import * as React from 'react';
import roomStyles from '../Room.module.scss';
import styles from '../MergedCalendar.module.scss';

import { IMultiBookProps } from './IMultiBookProps';
import {Stack, TextField, Dropdown, DatePicker, IDatePickerStrings, DayOfWeek, IComboBoxOption, Toggle, PrimaryButton, DefaultButton, Dialog, DialogType, DialogFooter} from '@fluentui/react';

export default function IMultiBook(props: IMultiBookProps) {

    const stackTokens = { childrenGap: 10 };
    const firstDayOfWeek = DayOfWeek.Sunday;
    const DayPickerStrings: IDatePickerStrings = {
        months: [
          'January',
          'February',
          'March',
          'April',
          'May',
          'June',
          'July',
          'August',
          'September',
          'October',
          'November',
          'December',
        ],
        shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
        days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
        shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
        goToToday: 'Go to today',
        prevMonthAriaLabel: 'Go to previous month',
        nextMonthAriaLabel: 'Go to next month',
        prevYearAriaLabel: 'Go to previous year',
        nextYearAriaLabel: 'Go to next year',
        closeButtonAriaLabel: 'Close date picker',
        monthPickerHeaderAriaLabel: '{0}, select to change the year',
        yearPickerHeaderAriaLabel: '{0}, select to change the month',
    };

    return(
        <React.Fragment>
        <div className={roomStyles.bookingForm}>

            <Stack tokens={stackTokens}>
                <TextField 
                    label="Title" 
                    required 
                    value={props.formField.titleField} 
                    onChange={props.onChangeFormField('titleField')}
                />  
                <TextField 
                    label="Description"
                    multiline rows={3}
                    value={props.formField.descpField} 
                    onChange={props.onChangeFormField('descpField')}
                />   
                <DatePicker
                    isRequired={true}
                    firstDayOfWeek={firstDayOfWeek}
                    strings={DayPickerStrings}
                    label="Start Date"
                    ariaLabel="Select a date"
                    onSelectDate={props.onChangeFormField('startDateField')}
                    value={props.formField.startDateField}
                />
                <DatePicker
                    isRequired={true}
                    firstDayOfWeek={firstDayOfWeek}
                    strings={DayPickerStrings}
                    label="End Date"
                    ariaLabel="Select a date"
                    onSelectDate={props.onChangeFormField('endDateField')}
                    value={props.formField.endDateField}
                />
                <Dropdown 
                    placeholder="Select the school cycle 5-10" 
                    label="School Cycle" 
                    required
                    selectedKey = {props.schoolCategory === 'Sec' ? props.schoolNum : undefined}
                    options={props.schoolCycleOptions} 
                    onChange={props.onChangeFormField('schoolCycleField')} 
                    errorMessage={props.errorMsgField.schoolCycleField} 
                /> 
                <Dropdown 
                    placeholder="Select the day of the school cycle" 
                    label="Day of the School Cycle" 
                    required
                    options={props.schoolCycleDayOptions} 
                    onChange={props.onChangeFormField('schoolCycleDayField')} 
                    errorMessage={props.errorMsgField.schoolCycleDayField} 
                />   
                <Dropdown 
                    placeholder="Select a room" 
                    label="Room" 
                    required
                    options={props.roomOptions} 
                    onChange={props.onChangeFormField('roomField')} 
                    errorMessage={props.errorMsgField.roomField} 
                />
                <Dropdown 
                    placeholder="Select a period" 
                    label="Period" 
                    required
                    options={props.periodOptions} 
                    onChange={props.onChangeFormField('periodField')} 
                    errorMessage={props.errorMsgField.periodField} 
                />                      
                <Toggle 
                    label="Add this event's booking to my Calendar" 
                    onText="Yes" 
                    offText="No" 
                    checked={props.formField.addToCalField}
                    onChange={props.onChangeFormField('addToCalField')}
                />
                                
            </Stack>
        </div>
        <div>
            <PrimaryButton text="Check Bookings" onClick={props.checkBookingClick} className={styles.marginR10}/>
            <DefaultButton text="Cancel" onClick={props.dismissPanelMultiBook}  />
        </div>
        </React.Fragment>
    );
}
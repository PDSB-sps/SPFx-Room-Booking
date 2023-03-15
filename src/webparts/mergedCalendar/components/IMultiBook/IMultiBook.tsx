import * as React from 'react';
import roomStyles from '../Room.module.scss';
import styles from '../MergedCalendar.module.scss';
import { IMultiBookProps } from './IMultiBookProps';
import {Stack, IStackStyles, IStackProps, TextField, Dropdown, DatePicker, IDatePickerStrings, DayOfWeek, Toggle, PrimaryButton, DefaultButton, IIconProps, initializeIcons, Icon} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export default function IMultiBook(props: IMultiBookProps) {

    initializeIcons();
    const stackTokens = { childrenGap: 10 };
    const stackStyles: Partial<IStackStyles> = { root: { width: '100%' } };
    const columnProps: Partial<IStackProps> = {
        tokens: { childrenGap: 15 },
        styles: { root: { width: '50%' } },
    };

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
        isOutOfBoundsErrorMessage: props.errorMsgField.endDateField,
    };
    // console.log("multibook props.formField", props.formField);
    // console.log("props.schoolNum", props.schoolNum);
    // console.log("props.schoolCategory", props.schoolCategory);

    return(
        <React.Fragment>
        <div className={roomStyles.bookingForm}>
            <Stack tokens={stackTokens}>
                <TextField 
                    label="Title" 
                    required 
                    value={props.formField.titleField} 
                    onChange={props.onChangeFormField('titleField')}
                    errorMessage={props.errorMsgField.titleField} 
                />  
                <TextField 
                    label="Description"
                    multiline rows={3}
                    value={props.formField.descpField} 
                    onChange={props.onChangeFormField('descpField')}
                />   
            </Stack>
            <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                <Stack {...columnProps}>
                    <DatePicker
                        isRequired
                        firstDayOfWeek={firstDayOfWeek}
                        strings={DayPickerStrings}
                        label="Start Date"
                        ariaLabel="Select a date"
                        onSelectDate={props.onChangeFormField('startDateField')}
                        value={props.formField.startDateField}
                        minDate={new Date()}
                    />
                    <Dropdown 
                        placeholder="Select the school cycle 5-10" 
                        label="School Cycle" 
                        required
                        selectedKey = {props.schoolCategory === 'Sec' ? props.schoolNum : undefined}
                        // defaultSelectedKey = {props.schoolCategory === 'Sec' ? props.schoolNum : undefined}
                        options={props.schoolCycleOptions} 
                        onChange={props.onChangeFormField('schoolCycleField')} 
                        errorMessage={props.errorMsgField.schoolCycleField} 
                        // disabled = {props.schoolCategory === 'Sec'}
                    />
                    <Dropdown 
                        placeholder="Select a room" 
                        label="Room" 
                        required
                        options={props.roomOptions} 
                        onChange={props.onChangeFormField('roomField')} 
                        errorMessage={props.errorMsgField.roomField} 
                    />
                </Stack>
                <Stack {...columnProps}>
                    <DatePicker
                        isRequired
                        firstDayOfWeek={firstDayOfWeek}
                        strings={DayPickerStrings}
                        label="End Date"
                        ariaLabel="Select a date"
                        onSelectDate={props.onChangeFormField('endDateField')}
                        value={props.formField.endDateField}
                        minDate={new Date(props.formField.startDateField)}
                        textField={{errorMessage: props.errorMsgField.endDateField}}
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
                        placeholder="Select a period" 
                        label="Period" 
                        required
                        options={props.periodOptions} 
                        onChange={props.onChangeFormField('periodField')} 
                        errorMessage={props.errorMsgField.periodField} 
                    />   
                </Stack>
            </Stack>
            {false &&
                <Stack tokens={stackTokens}>
                    <Toggle 
                        label="Add this event's booking to my Calendar" 
                        onText="Yes" 
                        offText="No" 
                        checked={props.formField.addToCalField}
                        onChange={props.onChangeFormField('addToCalField')}
                    />
                    {props.formField.addToCalField &&
                        <>
                            <PeoplePicker
                                context={props.context}
                                titleText="Invite Attendees"
                                groupName={''} // Leave this blank in case you want to filter from all users
                                showtooltip={false}
                                required={false}
                                onChange={props.onChangeFormField('attendees')}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.DistributionList, PrincipalType.SecurityGroup]}
                                resolveDelay={1000} 
                                personSelectionLimit={50}
                                defaultSelectedUsers = {props.invitedAttendees}
                            />
                            <p className={roomStyles.eventWarning}>
                                <Icon className={roomStyles.eventWarningIcon} iconName='Info'/> 
                                <span>Only board employees</span>
                            </p>
                        </>

                    }
                </Stack>
            }
        </div>
        <div>
            {!props.bookingsGridVisible ?
                <>
                    <PrimaryButton text="Check Bookings" onClick={props.checkBookingClick} className={styles.marginR10}/>
                    <DefaultButton text="Cancel" onClick={props.cancelMultiBook}  />
                </>
                :
                <>
                    <PrimaryButton iconProps={{iconName: 'Refresh'}} text="Reload Bookings" onClick={props.checkBookingClick} className={styles.marginR10}/>
                </>
            }
            
        </div>
        </React.Fragment>
    );
}
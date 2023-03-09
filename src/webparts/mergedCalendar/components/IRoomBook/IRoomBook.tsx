import * as React from 'react';
import {Stack, IStackStyles, TextField, Dropdown, DatePicker, IDatePickerStrings, DayOfWeek, ComboBox, IComboBox, IComboBoxOption, Toggle, PrimaryButton, DefaultButton, Dialog, DialogType, DialogFooter} from '@fluentui/react';
import styles from '../MergedCalendar.module.scss';
import roomStyles from '../Room.module.scss';
import { IRoomBookProps } from './IRoomBookProps';
import { useBoolean } from '@fluentui/react-hooks';
import {isUserManage} from '../../Services/RoomOperations';
import { IIconProps, initializeIcons, Icon } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export default function IRoomBook (props:IRoomBookProps) {
    
    initializeIcons();
    const deleteIcon: IIconProps = { iconName: 'Delete' };
    const editIcon: IIconProps = { iconName: 'Edit' };
    const checkIcon: IIconProps = { iconName: 'Accept' };
    const saveIcon: IIconProps = { iconName: 'Save' };

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

    const stackTokens = { childrenGap: 10 };
    const stackStyles: Partial<IStackStyles> = { root: { width: '100%' } };

    const [firstDayOfWeek, setFirstDayOfWeek] = React.useState(DayOfWeek.Monday);
    
    const hours: IComboBoxOption[] = [
        { key: '12 AM', text: '12 AM' },
        { key: '1 AM', text: '1 AM' },
        { key: '2 AM', text: '2 AM' },
        { key: '3 AM', text: '3 AM' },
        { key: '4 AM', text: '4 AM' },
        { key: '5 AM', text: '5 AM' },
        { key: '6 AM', text: '6 AM' },
        { key: '7 AM', text: '7 AM' },
        { key: '8 AM', text: '8 AM' },
        { key: '9 AM', text: '9 AM' },
        { key: '10 AM', text: '10 AM' },
        { key: '11 AM', text: '11 AM' },
        { key: '12 PM', text: '12 PM' },
        { key: '1 PM', text: '1 PM' },
        { key: '2 PM', text: '2 PM' },
        { key: '3 PM', text: '3 PM' },
        { key: '4 PM', text: '4 PM' },
        { key: '5 PM', text: '5 PM' },
        { key: '6 PM', text: '6 PM' },
        { key: '7 PM', text: '7 PM' },
        { key: '8 PM', text: '8 PM' },
        { key: '9 PM', text: '9 PM' },
        { key: '10 PM', text: '10 PM' },
        { key: '11 PM', text: '11 PM' },
    ];
    const minutes: IComboBoxOption[] = [
        { key: '00', text: '00' },
        { key: '05', text: '05' },
        { key: '10', text: '10' },
        { key: '15', text: '15' },
        { key: '20', text: '20' },
        { key: '25', text: '25' },
        { key: '30', text: '30' },
        { key: '35', text: '35' },
        { key: '40', text: '40' },
        { key: '45', text: '45' },
        { key: '50', text: '50' },
        { key: '55', text: '55' },
    ];


    let newStartKey = 1, newEndKey = 1;
    let initialTimeOptions: IComboBoxOption[] = [
        {key: '6:00 AM', text: '6:00 AM'},
        {key: '6:30 AM', text: '6:30 AM'},
        {key: '7:00 AM', text: '7:00 AM'},
        {key: '7:30 AM', text: '7:30 AM'},
        {key: '8:00 AM', text: '8:00 AM'},
        {key: '8:30 AM', text: '8:30 AM'},
        {key: '9:00 AM', text: '9:00 AM'},
        {key: '9:30 AM', text: '9:30 AM'},
        {key: '10:00 AM', text: '10:00 AM'},
        {key: '10:30 AM', text: '10:30 AM'},
        {key: '11:00 AM', text: '11:00 AM'},
        {key: '11:30 AM', text: '11:30 AM'},
        {key: '12:00 PM', text: '12:00 PM'},
        {key: '12:30 PM', text: '12:30 PM'},
        {key: '1:00 PM', text: '1:00 PM'},
        {key: '1:30 PM', text: '1:30 PM'},
        {key: '2:00 PM', text: '2:00 PM'},
        {key: '2:30 PM', text: '2:30 PM'},
        {key: '3:00 PM', text: '3:00 PM'},
        {key: '3:30 PM', text: '3:30 PM'},
        {key: '4:00 PM', text: '4:00 PM'},
        {key: '4:30 PM', text: '4:30 PM'},
        {key: '5:00 PM', text: '5:00 PM'},
        {key: '5:30 PM', text: '5:30 PM'},
        {key: '6:00 PM', text: '6:00 PM'},
        {key: '6:30 PM', text: '6:30 PM'},
        {key: '7:00 PM', text: '7:00 PM'},
        {key: '7:30 PM', text: '7:30 PM'},
        {key: '8:00 PM', text: '8:00 PM'},
        {key: '8:30 PM', text: '8:30 PM'},
        {key: '9:00 PM', text: '9:00 PM'},
    ];
    const [startSelectedKey, setStartSelectedKey] = React.useState<string | number | undefined>('C');
    const [endSelectedKey, setEndSelectedKey] = React.useState<string | number | undefined>('C');
    const [startTimeOptions, setStartTimeOptions] = React.useState(initialTimeOptions);
    const [endTimeOptions, setEndTimeOptions] = React.useState(initialTimeOptions);
    const onStartTimeChange = React.useCallback(
        (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void => {
          let key = option?.key;
          let text = option?.text;
          if (!option && value) {
            // If allowFreeform is true, the newly selected option might be something the user typed that
            // doesn't exist in the options list yet. So there's extra work to manually add it.
            //key = `${newStartKey++}`;
            key = value;
            text = value;
            setStartTimeOptions(prevOptions => [...prevOptions, { key: value, text: value }]);
          }
          setStartSelectedKey(key);
          props.onChangeFormTimesField('startTimeField', key, text);
        }
    ,[]);
    const onEndTimeChange = React.useCallback(
        (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void => {
            console.log("event -- in onEndTimeChange", event);
            console.log("option -- in onEndTimeChange", option);
            console.log("value -- in onEndTimeChange", value);
          let key = option?.key;
          let text = option?.text;
          console.log("key in onEndTimeChange", key);

          if (!option && value) {
            // If allowFreeform is true, the newly selected option might be something the user typed that
            // doesn't exist in the options list yet. So there's extra work to manually add it.
            //key = `${newEndKey++}`;
            key = value;
            text = value;
            setEndTimeOptions(prevOptions => [...prevOptions, { key: key, text: value }]);
          }
          setEndSelectedKey(key);
          props.onChangeFormTimesField('endTimeField', key, text);
        }
    ,[]);

    const disabledControl = props.bookFormMode === 'View' ? true: false;
    
    const modelProps = {
        isBlocking: false,
        styles: { main: { maxWidth: 450 } },
    };
    const dialogContentProps = {
        type: DialogType.largeHeader,
        title: 'Delete Booking',
        subText: 'Are you sure you want to delete this event booking?',
    };
    const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
    const panelHdrColor = props.bookFormMode === "New" ? props.roomInfo.Colour : props.eventDetailsRoom.RoomColor;

    const isPeriodKeyEmpty = props.formField.periodField.key == '' || props.formField.periodField.key == undefined;

    console.log("props.formField", props.formField);
    // console.log("props.bookFormMode", props.bookFormMode);
    // console.log("props.formField.startTimeField.key", props.formField.startTimeField.key);
    // console.log("props.formField.endTimeField.key", props.formField.endTimeField.key);
    // console.log("props.roomInfo", props.roomInfo);
    // console.log("props.eventDetailsRoom", props.eventDetailsRoom);
    
    return(
        <React.Fragment>
        <div className={roomStyles.bookingForm}>

            <div 
                style={{backgroundColor: panelHdrColor}} 
                className={roomStyles.roomColor}>
            </div>

            <div className={roomStyles.panelHdrOptions}>
                <h3>Booking Details</h3>
                {props.bookFormMode === "New" &&
                    <div className={roomStyles.editDeleteBtns}>
                        <PrimaryButton className={roomStyles.editBtn} iconProps={saveIcon} title="Save Booking" ariaLabel="Save Booking" onClick={props.onNewBookingClick} />
                    </div>
                }
                {props.bookFormMode === "Edit" &&
                    <div className={roomStyles.editDeleteBtns}>
                        <PrimaryButton className={roomStyles.editBtn} iconProps={checkIcon} title="Update Booking" ariaLabel="Update Booking" onClick={() => props.onUpdateBookingClick(props.eventDetailsRoom)} />
                    </div>
                }
                {props.bookFormMode === "View" && ( props.isCreator || isUserManage ) &&
                    <div className={roomStyles.editDeleteBtns}>
                        <PrimaryButton className={roomStyles.editBtn} iconProps={editIcon} title="Edit Booking" ariaLabel="Edit Booking" onClick={props.onEditBookingClick} />
                        <PrimaryButton className={roomStyles.deleteBtn} iconProps={deleteIcon} title="Delete Booking" ariaLabel="Delete Booking" onClick={toggleHideDialog} />
                    </div>
                }
            </div>

            {props.children}

            {props.roomInfo &&
                <div className={roomStyles.formComments}>                
                    <span>{props.roomInfo.OData__Comments}</span>
                </div>
            }

            <Stack tokens={stackTokens}>
                <TextField 
                    label="Title" 
                    required 
                    value={props.formField.titleField} 
                    onChange={props.onChangeFormField('titleField')} 
                    errorMessage={props.errorMsgField.titleField} 
                    disabled={disabledControl}
                    className={disabledControl ? roomStyles.disabledCtrl : ''}
                />  
                <TextField 
                    label="Description"
                    multiline rows={3}
                    value={props.formField.descpField} 
                    onChange={props.onChangeFormField('descpField')}
                    disabled={disabledControl}
                    className={disabledControl ? roomStyles.disabledCtrl : ''}
                />   
                <DatePicker
                    isRequired={true}
                    firstDayOfWeek={firstDayOfWeek}
                    strings={DayPickerStrings}
                    label="Date"
                    ariaLabel="Select a date"
                    onSelectDate={props.onChangeFormField('dateField')}
                    value={props.formField.dateField}
                    disabled={disabledControl}
                    className={disabledControl ? roomStyles.disabledCtrl : ''}
                    minDate={new Date()}
                />
                {((props.isPeriods && props.bookFormMode === 'New') || (props.bookFormMode !== 'New' && props.formField.isBookedByPeriods)) &&
                    <Dropdown 
                        placeholder="Select a period" 
                        label="Period" 
                        required
                        selectedKey={props.formField.periodField ? props.formField.periodField.key : undefined}
                        options={props.periodOptions} 
                        onChange={props.onChangeFormField('periodField')} 
                        errorMessage={props.errorMsgField.periodField} 
                        disabled={disabledControl}
                        className={disabledControl ? roomStyles.disabledCtrl : ''}
                    />     
                }
                {((!props.isPeriods && props.bookFormMode === 'New') || (props.bookFormMode !== 'New' && !props.formField.isBookedByPeriods) ) &&
                    <>
                        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                            <ComboBox
                                label="Start Time"
                                allowFreeform
                                autoComplete={'on'}
                                options={startTimeOptions}
                                onChange={onStartTimeChange}
                                defaultSelectedKey={props.formField.startTimeField ? props.formField.startTimeField.key : startSelectedKey}
                                text={props.formField.startTimeField ? props.formField.startTimeField.text : startSelectedKey}
                                useComboBoxAsMenuWidth
                                required
                                disabled={disabledControl}
                            />
                            <ComboBox
                                label="End Time"
                                allowFreeform
                                autoComplete={'on'}
                                options={endTimeOptions}
                                onChange={onEndTimeChange}
                                defaultSelectedKey={props.formField.endTimeField ? props.formField.endTimeField.key : endSelectedKey}
                                text={props.formField.endTimeField ? props.formField.endTimeField.text : endSelectedKey}
                                useComboBoxAsMenuWidth
                                required
                                disabled={disabledControl}
                            />
                        </Stack>
                        <p className={roomStyles.formAstrisk}>{props.errorMsgField.startEndTimeFields}</p>
                    </>
                }              
                <Toggle 
                    label="Add this event's booking to my Calendar" 
                    onText="Yes" 
                    offText="No" 
                    checked={props.formField.addToCalField}
                    onChange={props.onChangeFormField('addToCalField')}
                    disabled={disabledControl}
                />
                {props.formField.addToCalField &&
                    <>
                        <PeoplePicker
                            context={props.context}
                            titleText="Invite Attendees"
                            groupName={''} // Leave this blank in case you want to filter from all users
                            showtooltip={false}
                            required={false}
                            disabled={disabledControl}
                            onChange={props.onChangeFormField('attendees')}
                            showHiddenInUI={false}
                            principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.DistributionList, PrincipalType.SecurityGroup]}
                            resolveDelay={1000} 
                            personSelectionLimit={50}
                            defaultSelectedUsers = {props.invitedAttendees}
                        />
                        <p className={roomStyles.eventWarning}>
                            <Icon className={roomStyles.eventWarningIcon} iconName='Info'/> 
                            <span>Only internal board users</span>
                        </p>
                    </>

                }
                {/* 
                {props.bookFormMode === 'Edit' && props.formField.addToCalField &&
                    <p className={roomStyles.eventWarning}>
                        <Icon className={roomStyles.eventWarningIcon} iconName='Info'/> 
                        <span>Please note that by updating this event, this will a add new event to your <i>personal calendar</i>. You will have to manually delete the old one.</span>
                    </p>
                }                      
                */}
            </Stack>
        </div>
        <div>
            <Dialog
                hidden={hideDialog}
                onDismiss={toggleHideDialog}
                dialogContentProps={dialogContentProps}
                modalProps={modelProps}
            >
                <DialogFooter>
                    <PrimaryButton onClick={() => props.onDeleteBookingClick(props.eventDetailsRoom)} text="Yes" />
                    <DefaultButton onClick={toggleHideDialog} text="No" />
                </DialogFooter>
            </Dialog>

            {props.bookFormMode === "New" &&
                <PrimaryButton text="Book" onClick={props.onNewBookingClick} className={styles.marginR10}/>
            }            
            {props.bookFormMode === "Edit" &&
                <PrimaryButton text="Update" onClick={() => props.onUpdateBookingClick(props.eventDetailsRoom)} className={styles.marginR10}/>
            }
            <DefaultButton text="Cancel" onClick={props.dismissPanelBook}  />
        </div>
        </React.Fragment>
    );
}
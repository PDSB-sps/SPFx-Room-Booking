import * as React from 'react';
import styles from './MergedCalendar.module.scss';
import roomStyles from './Room.module.scss';
import { IMergedCalendarProps } from './IMergedCalendarProps';
import {IDropdownOption, DefaultButton, Panel, IComboBox, IComboBoxOption, MessageBar, MessageBarType, MessageBarButton, PanelType, Dialog, DialogFooter, DialogType} from '@fluentui/react';
import {useBoolean} from '@fluentui/react-hooks';

import {CalendarOperations} from '../Services/CalendarOperations';
import {updateCalSettings} from '../Services/CalendarSettingsOps';
import {addToMyGraphCal, getMySchoolCalGUID, reRenderCalendars, getLegendChksState, calsErrs} from '../Services/CalendarRequests';
import {getGraphCalsMultiBook, getAllPeriods, getSchoolCategory, getSchoolCycles, getBookedEvents, mergeBookings, addBooking} from '../Services/MultiBookOperations';
import {formatEvDetails} from '../Services/EventFormat';
import {setWpData} from '../Services/WpProperties';
import {getRooms, getPeriods, getLocationGroup, getGuidelines, getRoomsCalendarName, addEvent, deleteItem, updateEvent, isEventCreator, getRoomInfo} from '../Services/RoomOperations';
import {isUserManage} from '../Services/RoomOperations';

import ICalendar from './ICalendar/ICalendar';
import IPanel from './IPanel/IPanel';
import ILegend from './ILegend/ILegend';
import IDialog from './IDialog/IDialog';
import IRooms from './IRooms/IRooms';
import IRoomBook from './IRoomBook/IRoomBook';
import IRoomDetails from './IRoomDetails/IRoomDetails';
import IRoomDropdown from './IRoomDropdown/IRoomDropdown';
import IRoomGuidelines from './IRoomGuidelines/IRoomGuidelines';
import IRoomsManage from './IRoomsManage/IRoomsManage';
import IMultiBook from './IMultiBook/IMultiBook';
import IMultiBookList from './IMultiBookList/IMultiBookList';

import toast, { Toaster } from 'react-hot-toast';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { PrimaryButton } from 'office-ui-fabric-react';
import ILegendRooms from './ILegendRooms/ILegendRooms';
import IPreloader from './IPreloader/IPreloader';
import * as moment from 'moment';


export default function MergedCalendar (props:IMergedCalendarProps) {
  
  // Calendar states & Event details states
  const _calendarOps = new CalendarOperations();
  const [eventSources, setEventSources] = React.useState([]);
  const [calSettings, setCalSettings] = React.useState([]);
  const [eventDetailsRoom, setEventDetailsRoom] = React.useState(null);
  const [eventDetails, setEventDetails] = React.useState(null);
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
  const [isDataLoading, { toggle: toggleIsDataLoading }] = useBoolean(false);
  const [showWeekends, { toggle: toggleshowWeekends }] = useBoolean(props.showWeekends);
  const [listGUID, setListGUID] = React.useState('');
  const [calVisibility, setCalVisibility] = React.useState <{calId: string, calChk: boolean}>({calId: null, calChk: null});
  const [currentCalDate, setCurrentCalDate] = React.useState(new Date().toISOString());

  // Room Booking states
  const [rooms, setRooms] = React.useState([]);
  const [roomId, setRoomId] = React.useState(null);
  const [roomLoadedId, setRoomLoadedId] = React.useState(roomId);
  const [roomInfo, setRoomInfo] = React.useState(null);
  const [selectedPeriod, setSelectedPeriod] = React.useState('');
  const [isCreator, setIsCreator] = React.useState(false);
  const [isOpenDetails, { setTrue: openPanelDetails, setFalse: dismissPanelDetails }] = useBoolean(false);
  const [isOpenBook, { setTrue: openPanelBook, setFalse: dismissPanelBook }] = useBoolean(false);
  const [bookFormMode, setBookFormMode] = React.useState('New');
  const [filteredRooms, setFilteredRooms] = React.useState(rooms);
  const [roomSelectedKey, setRoomSelectedKey] = React.useState<string | number | undefined>('all');
  const [locationGroup, setLocationGroup] = React.useState([]);
  const [periods, setPeriods] = React.useState([]);
  const [guidelines, setGuidelines] = React.useState([]);
  const [isFiltered, { setTrue: showFilterWarning, setFalse: hideFilterWarning }] = useBoolean(false);
  const [roomsCalendar, setRoomsCalendar] = React.useState('Events');
  const [calsVisibility, setCalsVisibility] = React.useState([]);
  const [calMsgErrs, setCalMsgErrs] = React.useState([]);  

  // Multi-booking states
  const [isOpenMultiBook, { setTrue: openPanelMultiBook, setFalse: dismissPanelMultiBook }] = useBoolean(false);
  const [allPeriods, setAllPeriods] = React.useState([]);
  const {schoolNum, schoolCategory} = getSchoolCategory(window.location.href);
  const [cycleDays, setCycleDays] = React.useState([]);
  const [cycleDaysCalUrl, setCycleDaysCalUrl] = React.useState('');
  const [mergedBookings, setMergedBookings] = React.useState([]);
  const [isConflict, setIsConflict] = React.useState(false);
  const [isCheckBookingClicked, setIsCheckBookingClicked] = React.useState(false);
  const [hideConfirmMultiDlg, { toggle: toggleConfirmMultiDlg }] = useBoolean(true);
  const [isMultiBookingDataLoading, { toggle: toggleIsMultiBookingDataLoading }] = useBoolean(false);

  const ACTIONS = {
    EVENT_DETAILS_TOGGLE : "event-details-toggle",
    ROOM_DELETE_TOGGLE : "room-delete-toggle",
    ROOMS_MANAGE : "rooms-manage",
    IFRAME_DISMISS: "iframe-dimiss",
    ROOM_EDIT: "room-edit",
    LOAD_EVENTS: "load-events",
    LOAD_EVENTS_VIS : "load-events-visibility"
  };

  const dlgInitialState : any = {dlgDetails: false, dlgDelete: false};
  const dialogReducer = (state : any, action: any) =>{
    switch (action.type){
      case ACTIONS.EVENT_DETAILS_TOGGLE:
        return {...state, dlgDetails: !state.dlgDetails};
      case ACTIONS.ROOM_DELETE_TOGGLE:
        return {...state, dlgDelete: !state.dlgDelete};
      default: 
        return state;
    }
  };
  const [dialogState, dialogDispatch] = React.useReducer(dialogReducer, dlgInitialState);

  const calSettingsList = props.calSettingsList ? props.calSettingsList : "CalendarSettings";
  const roomsList = props.roomsList ? props.roomsList : "Rooms";
  const periodsList = props.periodsList ? props.periodsList : "Periods";
  const guidelinesList = props.guidelinesList ? props.guidelinesList : "Guidelines";
  
  const eventSourcesReducer = (state: any, action: any) => {
    switch (action.type){
      case ACTIONS.LOAD_EVENTS:
        return [...action.payload];
      case ACTIONS.LOAD_EVENTS_VIS:
        const prevEventSources = state;
        let tempEventSources = [];
        action.payload.map(calVis =>{
          if (calVis.calId){
            tempEventSources = reRenderCalendars(prevEventSources, calVis);
          }
        });
      return [...tempEventSources];
      default:
        return state;
    }
  };
  const [eventSourcesState, dispatchEventSources] = React.useReducer(eventSourcesReducer, []);

  const loadLatestCalendars = async (callback?: any, displayPreloader?:boolean) =>{
    console.log("loadLatestCalendars Function!");
    
    if (displayPreloader == undefined) displayPreloader = true;
    if(displayPreloader) toggleIsDataLoading();
    _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, roomId).then((results: any)=>{
      setRoomsCalendar(getRoomsCalendarName(results[0]));
      setCalSettings(results[0]);
      //setEventSources(results[1]);
      dispatchEventSources({type: ACTIONS.LOAD_EVENTS, payload: results[1] });
      if (calsVisibility.length > 1){
        dispatchEventSources({type: ACTIONS.LOAD_EVENTS_VIS, payload: calsVisibility });
      }
      if(displayPreloader) toggleIsDataLoading();
      setCalMsgErrs(calsErrs);
      callback ? callback() : null;
    });
  };

  const popToast = (toastMsg: string) =>{
    toast.success(toastMsg, {
      duration: 2000,
      style: {
        margin: '150px',
      },
      className: roomStyles.popNotif            
    });
  };

  // UseEffect(s)
  React.useEffect(()=>{
    loadLatestCalendars();
    getRooms(props.context, roomsList).then((results)=>{
      setRooms(results);
      setFilteredRooms(results);
    });    
  },[roomId]);

  React.useEffect(()=>{
    loadLatestCalendars(null, false);
  },[currentCalDate]);

  React.useEffect(()=>{
    getLocationGroup(props.context, roomsList).then((results)=>{
      setLocationGroup(results);
    });    
  }, []);

  React.useEffect(()=>{
    // setEventSources(reRenderCalendars(eventSources, calVisibility));
    const updatedEventSources = reRenderCalendars(eventSourcesState, calVisibility);
    dispatchEventSources({type: ACTIONS.LOAD_EVENTS, payload: updatedEventSources});
    setCalsVisibility(getLegendChksState(calsVisibility, calVisibility));
  },[calVisibility]);

  const chkHandleChange = (newCalSettings:{})=>{    
    return (ev: any, checked: boolean) => { 
      toggleIsDataLoading();
      updateCalSettings(props.context, calSettingsList, newCalSettings, checked).then(()=>{
        loadLatestCalendars(toggleIsDataLoading);
      });
     };
  };  
  const dpdHandleChange = (newCalSettings:any)=>{
    return (ev: any, item: IDropdownOption) => { 
      toggleIsDataLoading();
      updateCalSettings(props.context, props.calSettingsList, newCalSettings, newCalSettings.ShowCal, item.key).then(()=>{
        loadLatestCalendars(toggleIsDataLoading);
      });
     };
  };
  const chkViewHandleChange = (ev: any, checked: boolean) =>{
    toggleIsDataLoading();
    setWpData(props.context, "showWeekends", checked).then(()=>{
      toggleshowWeekends();
      toggleIsDataLoading();
    });
  };
  
  const handleAddtoCal = ()=>{
    addToMyGraphCal(props.context).then((result)=>{
      console.log('calendar updated', result);
    });
  };

  //Booking Forms states
  const [formField, setFormField] = React.useState({
    titleField: "",
    descpField: "",
    periodField : {key: '', text:'', start:new Date(), end:new Date()},
    dateField : new Date(),
    addToCalField: false
  });
  //error handeling
  const [errorMsgField , setErrorMsgField] = React.useState({
    titleField: "",
    periodField : "",
  });
  const resetFields = () =>{
    setFormField({
    titleField: "",
    descpField: "",
    periodField : {key: '', text:'', start:new Date(), end:new Date()},
    dateField : new Date(),    
    addToCalField: false
    });
    setErrorMsgField({
      titleField: "",
      periodField : "",
    });
  };
  const onChangeFormField = (formFieldParam: string) =>{
    return (event: any, newValue?: any)=>{
      //Note to self
      //(newValue === undefined && typeof event === "object") //this is for date
      //for date, there is no 2nd param, the newValue is the main one
      //typeof newValue === "boolean" //this one for toggle buttons
      setFormField({
        ...formField,
        [formFieldParam]: (newValue === undefined && typeof event === "object") ? event : (typeof newValue === "boolean" ? !!newValue : newValue || ''),
      });

      setErrorMsgField({titleField: "", periodField: ""});
      // const calculatedRoomId = roomInfo ? roomInfo.Id : eventDetailsRoom.RoomId;
      // console.log("bookFormMode", bookFormMode);
      const calculatedRoomId = bookFormMode === "New" ? roomInfo.Id : eventDetailsRoom.RoomId;

      // console.log("formField.dateField", formField.dateField);
      // console.log("event", event);
      // console.log("formField.dateField !== event", moment(formField.dateField).format('MM-DD-YYYY') !== moment(event).format('MM-DD-YYYY'));

      if(formFieldParam === 'dateField' && moment(formField.dateField).format('MM-DD-YYYY') !== moment(event).format('MM-DD-YYYY')){
        setFormField((prevState)=>{
          return {...prevState, periodField : {key: '', text:'', start:new Date(), end:new Date()}};
        });
        getPeriods(props.context, periodsList, calculatedRoomId , event, null).then((results)=>{
          setPeriods(results);
        });
      }
    };
  };

  const handleDateClick = (arg:any) =>{    
    if(arg.event._def.extendedProps.roomId){
      const evDetails: any = formatEvDetails(arg);
      setEventDetailsRoom(evDetails);
      setSelectedPeriod(evDetails.PeriodId);
      setBookFormMode('View');
      
      isEventCreator(props.context, roomsCalendar, evDetails.EventId).then((v)=>{
        setIsCreator(v);
      });
      getGuidelines(props.context, guidelinesList).then((results)=>{
        setGuidelines(results);
      });
      getPeriods(props.context, periodsList, evDetails.RoomId, new Date(evDetails.Start), evDetails.PeriodId).then((results)=>{
        setPeriods(results);
      });
      getRoomInfo(props.context, roomsList, arg.event._def.extendedProps.roomId).then(results => setRoomInfo(results));

      setFormField({
        ...formField,
        titleField: evDetails.Title,
        descpField: evDetails.Body,
        periodField : {key: evDetails.PeriodId, text:evDetails.Period, start:new Date(evDetails.Start), end:new Date(evDetails.End)},
        dateField : new Date(evDetails.Start),    
        addToCalField: evDetails.AddToMyCal
      });    

      dismissPanelDetails();
      openPanelBook();
    }
    else{
      setEventDetails(formatEvDetails(arg));
      dialogDispatch({type: ACTIONS.EVENT_DETAILS_TOGGLE});
    }
  };
  
  const handleError = (callback:any) =>{
    if (formField.titleField == "" && formField.periodField.key == ""){
      setErrorMsgField({titleField: "Title Field Required", periodField: "Period Field Required"});
    }
    else if (formField.titleField == ""){
      setErrorMsgField({titleField: "Title Field Required", periodField: ""});
    }
    else if (formField.periodField.key == ""){
      setErrorMsgField({titleField: "", periodField: "Period Field Required"});
    }
    else{
      setErrorMsgField({titleField: "", periodField: ""});
      callback();
    }
  };

  //Filter Rooms
  const onFilterChanged = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    setRoomSelectedKey(option.key);
    if(option.key === 'all'){
      setFilteredRooms(rooms);
    }else{
      setFilteredRooms(rooms.filter(room => room.LocationGroup.toLowerCase().indexOf(option.text.toLowerCase()) >= 0));
    }
  };

  //Rooms functions
  const onCheckAvailClick = (roomIdParam: number) =>{
    setRoomId(roomIdParam);
    showFilterWarning();
  };
  const onResetRoomsClick = ()=>{
    setRoomId(null);
    hideFilterWarning();
  };
  const onViewDetailsClick = (roomInfoParam: any) =>{
    setRoomInfo(roomInfoParam);
    dismissPanelBook();
    openPanelDetails();
  };
  const onBookClick = (bookingInfoParam: any) =>{
    setBookFormMode('New');
    getPeriods(props.context, periodsList, bookingInfoParam.roomInfo.Id, formField.dateField).then((results)=>{
      setPeriods(results);
    });
    getGuidelines(props.context, guidelinesList).then((results)=>{
      setGuidelines(results);
    });
    resetFields();
    setRoomInfo(bookingInfoParam.roomInfo);
    dismissPanelDetails();
    openPanelBook();
  };

  //when clicking on the book button in the panel
  const onNewBookingClickHandler = ()=>{
    handleError(()=>{
      
      getPeriods(props.context, periodsList, roomInfo.Id, formField.dateField).then((results: any)=>{
        setPeriods(results);
        
        let seletedPeriod = results.filter(item => item.key === formField.periodField.key);
        if (!seletedPeriod[0].disabled){          
          addEvent(props.context, roomsCalendar, formField, roomInfo).then(()=>{
            const callback = () =>{
              dismissPanelBook();
              popToast('A New Event Booking is successfully added!');                              
            };
            loadLatestCalendars(callback);
          });
        }else{ //Period already booked
          setErrorMsgField({titleField: "", periodField: "Looks like the period is already booked! Please choose another one."});
          setFormField({
            ...formField,
            periodField : {key: '', text:'', start:new Date(), end:new Date()}
          });
        }
      });
      
    });
  };

  const onEditBookingClickHandler = () =>{
    setBookFormMode('Edit');
  };
  const onDeleteBookingClickHandler = (eventIdParam: any) =>{
    deleteItem(props.context, roomsCalendar, eventIdParam).then(()=>{
      const callback = () =>{
        dismissPanelBook();
        popToast('The Event Booking is successfully deleted!');   
      };
      loadLatestCalendars(callback);
    });
  };
  const onUpdateBookingClickHandler = (eventIdParam: any) =>{
    getPeriods(props.context, periodsList, eventDetailsRoom.RoomId, formField.dateField, eventDetailsRoom.PeriodId).then((results: any)=>{
      setPeriods(results);
      
      let seletedPeriod = results.filter(item => item.key === formField.periodField.key);
      if (!seletedPeriod[0].disabled){          
        updateEvent(props.context, roomsCalendar, eventIdParam, formField, eventDetailsRoom).then(()=>{
          const callback = () =>{
            dismissPanelBook();
            popToast('Event Booking is successfully updated!');
          };
          loadLatestCalendars(callback);
        });
      }else{ //Period already booked
        setErrorMsgField({titleField: "", periodField: "Looks like the period is already booked! Please choose another one."});
        setFormField({
          ...formField,
          periodField : {key: '', text:'', start:new Date(), end:new Date()}
        });
      }
    });
  };

  //Rooms, Periods, Guidelines Management
  const iFrameInitialState : any = {iFrameUrl: '', iFrameShow: false, iFrameState: 'Add'};
  const iFrameReducer = (state: any, action: any) =>{
    switch (action.type){
      case ACTIONS.ROOMS_MANAGE:
        return {iFrameUrl: action.payload.iFrameUrl, iFrameShow: action.payload.iFrameShow, iFrameState: action.payload.iFrameState};
      case ACTIONS.IFRAME_DISMISS:
        return {iFrameShow: action.payload.iFrameShow};
      case ACTIONS.ROOM_EDIT:
        return {iFrameUrl: action.payload.iFrameUrl, iFrameShow: action.payload.iFrameShow};
    }
  };
  const [iFrameState, iFrameDispatch] = React.useReducer(iFrameReducer, iFrameInitialState);

  const onRoomsManage = (newFormUrl: string, roomsManageState: string) =>{
    iFrameDispatch({type: ACTIONS.ROOMS_MANAGE, payload: {iFrameUrl: newFormUrl, iFrameShow: true, iFrameState: roomsManageState}});
  };
  const onIFrameDismiss = async (event: React.MouseEvent) => {
    iFrameDispatch({type: ACTIONS.IFRAME_DISMISS, payload: {iFrameShow: false}});

    getRooms(props.context, roomsList).then((results)=>{
      setRooms(results);
      setFilteredRooms(results);
    });
    getGuidelines(props.context, guidelinesList).then((results)=>{
        setGuidelines(results);
    });
  };
  const onIFrameLoad = async (iframe: any) => {
    if(iframe.contentWindow.location.href.indexOf('AllItems.aspx') > 0 && iFrameState.iFrameState !== 'All')
      onIFrameDismiss(null);  
  };
  const onEditRoom = (editedRoomId: any) =>{
    const editRoomUrl = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${roomsList}/Editform.aspx?ID=${editedRoomId}`;
    iFrameDispatch({type: ACTIONS.ROOM_EDIT, payload: {iFrameUrl: editRoomUrl, iFrameShow: true}});

    getRooms(props.context, roomsList).then((results)=>{
      setRooms(results);
      setFilteredRooms(results);
    });
  };
  const onDeleteRoomClickHandler = (roomIdParam: any) =>{
    deleteItem(props.context, roomsList, roomIdParam).then(()=>{
      const callback = () =>{
        dismissPanelDetails();
        dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE});
        popToast('The Room is successfully deleted!');   
      };
      getRooms(props.context, roomsList).then((results)=>{
        setRooms(results);
        setFilteredRooms(results);
        callback();
      });
    });
  };
  const modelProps = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
  };
  const dialogContentProps = {
      type: DialogType.largeHeader,
      title: 'Delete Room',
      subText: 'Are you sure you want to delete this room?',
  };
  const onDeleteRoomClick = (roomIdParam: any) =>{
    setRoomLoadedId(roomIdParam);
    dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE});
  };

  const onLegendChkChange = (calId: string) =>{
    return(ev: any, checked: boolean) =>{
        setCalVisibility({calId: calId, calChk: checked});
    };
  };

  // Current date passing for window range
  const passCurrentDate = (currDate: string) => {
    console.log("passCurrentCalDate function", currDate);
    setCurrentCalDate(currDate);
  };

  /* Multiple Room Booking */
  const multiBookPanelOpenHandler = () => {
    getAllPeriods(props.context, periodsList).then(results => {
      setAllPeriods(results);
    });
    if (schoolCategory === 'Sec'){
      getSchoolCycles(props.context, schoolNum).then(results => {
        setCycleDaysCalUrl(results.calUrl);
        setCycleDays(results.cycleDays.map(item => ({key: `Day${item}`, text: `Day ${item}`})));
      });
    }
    openPanelMultiBook();    
  };
  //Booking Forms states
  const [formFieldMultiBk, setFormFieldMultiBk] = React.useState({
    titleField: "",
    descpField: "",
    schoolCycleField : {key: '', text:''}, 
    schoolCycleDayField : {key: '', text:''}, 
    periodField : {key: '', text:'', start:new Date(), end:new Date()},
    roomField : {key: '', text:''},
    startDateField : new Date(),
    endDateField : new Date(),
    addToCalField: false
  });
  //error handeling
  const [errorMsgFieldMultiBk , setErrorMsgFieldMultiBk] = React.useState({
    titleField: "",
    schoolCycleField : "",
    schoolCycleDayField : "",
    periodField : "",
    roomField : "",
    startDateField : "",    
    endDateField : "",    
  });
  const resetErrorMsgFieldMultiBk = () => {
    setErrorMsgFieldMultiBk({
      titleField: "",
      schoolCycleField : "",
      schoolCycleDayField : "",
      periodField : "",
      roomField : "",
      startDateField : "",    
      endDateField : "",    
    });
  };
  const resetFieldsMultiBk = () =>{
    setFormFieldMultiBk({
      titleField: "",
      descpField: "",
      schoolCycleField : {key: '', text:''},
      schoolCycleDayField : {key: '', text:''},
      periodField : {key: '', text:'', start:new Date(), end:new Date()},
      roomField : {key: '', text:''},
      startDateField : new Date(),    
      endDateField : new Date(),    
      addToCalField: false
    });
    resetErrorMsgFieldMultiBk();
  };
  const handleErrorMultiBk = (callback:any) =>{
    let allFieldsValid = true;
    const fieldsMultiBkMapping = [
      {"field" : "titleField", value: formFieldMultiBk.titleField, "name" : "Title"},
      {"field" : "schoolCycleField", value: formFieldMultiBk.schoolCycleField.key, "name" : "School Cycle"},
      {"field" : "schoolCycleDayField", value: formFieldMultiBk.schoolCycleDayField.key, "name" : "Day of the School Cycle"},
      {"field" : "periodField", value: formFieldMultiBk.periodField.key, "name" : "Period"},
      {"field" : "roomField", value: formFieldMultiBk.roomField.key, "name" : "Room"},
      {"field" : "startDateField", value: formFieldMultiBk.startDateField, "name" : "Start Date"},
      {"field" : "endDateField", value: formFieldMultiBk.endDateField, "name" : "End Date"}
    ];
    for (let fieldMapping of fieldsMultiBkMapping){
      if (fieldMapping.value == ""){
        setErrorMsgFieldMultiBk(prevState => {
          return {
            ...prevState,
            [fieldMapping.field] : fieldMapping.name + " Field Required"
          };
        });
        allFieldsValid = false;
      }else{
        setErrorMsgFieldMultiBk(prevState => {
          return {
            ...prevState,
            [fieldMapping.field] : ""
          };
        });
      }
    }
    if (allFieldsValid) callback();
  };
  const onChangeFormFieldMultiBk = (formFieldParam: string) =>{
    return (event: any, newValue?: any)=>{
      setFormFieldMultiBk(prevState => { 
        return{
          ...prevState,
          [formFieldParam]: (newValue === undefined && typeof event === "object") ? event : (typeof newValue === "boolean" ? !!newValue : newValue || ''),
        };
      });
      setErrorMsgFieldMultiBk(prevState => {
        return {
          ...prevState,
          [formFieldParam] : ""
        };
      });
      if (formFieldParam === 'schoolCycleField' && schoolCategory === 'Elem') selectDayCycleHandler(newValue.key);
      if (formFieldParam === 'startDateField' && formFieldMultiBk.endDateField < new Date(event)) {
        setFormFieldMultiBk(prevState => { 
          return{
            ...prevState,
            endDateField: new Date(event),
          };
        });
      }
    };
  };
  const schoolCycleOptions = schoolCategory === 'Elem' 
    ? [{key: 'E5Day', text: '5 Day'}, {key: 'E10Day', text: '10 Day'}]
    : [{key: schoolNum, text: 'School Rotary'}];

  const selectDayCycleHandler = (schoolRotary: string) => {
    getSchoolCycles(props.context, schoolRotary).then(results => {
      setCycleDaysCalUrl(results.calUrl);
      setCycleDays(results.cycleDays.map(item => ({key: `Day${item}`, text: `Day ${item}`})));
    });
  };

  const checkBookingClickHandler = () => {
    handleErrorMultiBk(async()=>{
      console.log("formFieldMultiBk OK!", formFieldMultiBk);
      
      toggleIsMultiBookingDataLoading();
      
      const multiBookings = await getGraphCalsMultiBook(props.context, {
        CalType: 'Graph', 
        Title: formFieldMultiBk.schoolCycleField.key, 
        CalName: formFieldMultiBk.schoolCycleField.key, 
        CalURL: cycleDaysCalUrl
      }, formFieldMultiBk.startDateField.toISOString(), formFieldMultiBk.endDateField.toISOString(), formFieldMultiBk.schoolCycleDayField.text);
      
      const existingBookings = await getBookedEvents(props.context, {
        CalType: 'Room', 
        Title: formFieldMultiBk.schoolCycleField.key, 
        CalName: formFieldMultiBk.schoolCycleField.key, 
        CalURL: cycleDaysCalUrl
      }, formFieldMultiBk.roomField.key, formFieldMultiBk.periodField.key, formFieldMultiBk.startDateField.toISOString(), formFieldMultiBk.endDateField.toISOString());
      
      const {isConflictBool, mergedBookingsList} = mergeBookings(existingBookings, multiBookings, formFieldMultiBk);
      
      setMergedBookings(mergedBookingsList);
      setIsConflict(isConflictBool);
      setIsCheckBookingClicked(true);
      
      toggleIsMultiBookingDataLoading();
    });
  };
  const cancelBookingClickHandler = () => {
    dismissPanelMultiBook();
    resetFieldsMultiBk();
    setMergedBookings([]);
    setIsCheckBookingClicked(false);
  };
  const updateBookings = (itemIndex, checked) => {
    setMergedBookings(prevState => {
      return prevState.map(booking => {
        if(booking.index === itemIndex){
          return {...booking, overwrite: checked}
        }else{
          return booking;
        }
      })
    });
  };
  const mulitBookEventsHandler = () =>{

    if (!hideConfirmMultiDlg) toggleConfirmMultiDlg();

    // console.log("final bookings", mergedBookings);
    const finalBookings = [];
    const conflictBookingsIds = [];
    const finalBookingPromises = []; 

    for (let booking of mergedBookings){
      if (booking.overwrite){
        finalBookings.push({
          titleField: formFieldMultiBk.titleField,
          descpField: formFieldMultiBk.descpField,
          periodField: formFieldMultiBk.periodField,
          dateField: booking.start,
          addToCalField:formFieldMultiBk.addToCalField
        });
      }
      if (booking.overwrite && booking.conflict){
        conflictBookingsIds.push(booking.conflictId)
      }
    }
    for (let finalBooking of finalBookings){
      finalBookingPromises.push(addBooking(props.context, roomsCalendar, finalBooking, {Id: formFieldMultiBk.roomField.key, Title: formFieldMultiBk.roomField.text}));
    }
    for (let conflictId of conflictBookingsIds){
      finalBookingPromises.push(deleteItem(props.context, roomsCalendar, conflictId));
    }
    Promise.all(finalBookingPromises).then(values => {
      const callback = () =>{
        cancelBookingClickHandler();
        popToast(`Hurray! Your ${finalBookings.length} booking(s) are successfully added and conflicts resolved!`);
      };
      loadLatestCalendars(callback);
    });
  };

  return(
    <div className={styles.mergedCalendar}>

      <Toaster />
      
      <IPreloader 
        isDataLoading = {isDataLoading} 
        text = "Please wait, Loading Events..."
      />

      <div className={roomStyles.roomsCalendarCntnr}>
        <div className={roomStyles.allRoomsCntnr}> 
          {filteredRooms.length !== 0 ?
              <IRoomDropdown 
              onFilterChanged={onFilterChanged}
              roomSelectedKey={roomSelectedKey}
              locationGroup = {locationGroup}
            />
            :
            <MessageBar messageBarType={MessageBarType.warning} isMultiline={true} >
              There are no Rooms created yet. Please use the "Add" and "Edit" options below to manage your Rooms, Periods and Guidelines.
            </MessageBar>
          }

          {isUserManage &&
            <React.Fragment>
              <IRoomsManage 
                context={props.context}
                roomsList={props.roomsList}
                periodsList={props.periodsList}
                guidelinesList={props.guidelinesList}
                onRoomsManage={onRoomsManage}
                iframeState = {iFrameState.iFrameState}
                openMultiBook = {multiBookPanelOpenHandler}
              />   
              <Dialog
                hidden={!dialogState.dlgDelete}
                onDismiss={() => dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE})}
                dialogContentProps={dialogContentProps}
                modalProps={modelProps}>
                <DialogFooter>
                    <PrimaryButton onClick={() => onDeleteRoomClickHandler(roomLoadedId)} text="Yes" />
                    <DefaultButton onClick={() => dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE})} text="No" />
                </DialogFooter>
              </Dialog>
              <IFrameDialog 
                url={iFrameState.iFrameUrl}
                width={iFrameState.iFrameState === "Add" ? '40%' : '70%'}
                height={'90%'}
                hidden={!iFrameState.iFrameShow}
                iframeOnLoad={(iframe) => onIFrameLoad(iframe)}
                onDismiss={(event) => onIFrameDismiss(event)}
                allowFullScreen = {true}
                dialogContentProps={{
                  type: DialogType.close,
                  showCloseButton: true
                }}
              />
            </React.Fragment>
          }

          {filteredRooms.length !==0 &&
            <IRooms 
              rooms={filteredRooms} 
              onCheckAvailClick={() => onCheckAvailClick} 
              onBookClick={()=> onBookClick}
              onViewDetailsClick={()=>onViewDetailsClick}
              onEditClick={() => onEditRoom}
              onDeleteClick={() => onDeleteRoomClick}
            />
          }
          
        </div>

        <div className={roomStyles.allCalendarCntnr}>
        {isFiltered &&
          <div className={roomStyles.filterWarning}>
            <MessageBar 
              messageBarType={MessageBarType.warning}
              isMultiline={false}
              actions={
              <div>
                  <MessageBarButton onClick={onResetRoomsClick}>View All</MessageBarButton>
              </div>
              }
            >
              Please note that you are not viewing all resources now.
            </MessageBar>
          </div>
        }
        <div className={`${styles.legendTop} ${styles.legendHz}`}>
          <ILegend 
            calSettings={calSettings} 
            onLegendChkChange={onLegendChkChange}
          />
        </div>        
        <ICalendar 
          // eventSources={filteredEventSources} 
          // eventSources={eventSources} 
          eventSources={eventSourcesState}
          showWeekends={showWeekends}
          openPanel={openPanel}
          handleDateClick={handleDateClick}
          context={props.context}
          listGUID = {listGUID}
          passCurrentDate = {passCurrentDate}/>

        <ILegendRooms 
          calSettings={calSettings} 
          rooms={filteredRooms}
        />

      {calMsgErrs.length > 0 &&
        <MessageBar className={styles.calErrsMsg} messageBarType={MessageBarType.warning}>
          Warning! Calendar Errors, please check
          <ul>
            {calMsgErrs.map((msg)=>{
              return <li>{msg}</li> ;
            })}
          </ul>
        </MessageBar>
      }
      </div>
      </div>
      <MessageBar className={roomStyles.helpMsgBar} isMultiline={false}>
        Need help? 
        <a href="https://pdsb1.sharepoint.com/ltss/classtech/rbs" target="_blank" data-interception="off">
          Visit our website.
        </a>
      </MessageBar>

      <IPanel
        dpdOptions={props.dpdOptions} 
        calSettings={calSettings}
        onChkChange={chkHandleChange}
        onDpdChange={dpdHandleChange}
        isOpen = {isOpen}
        dismissPanel = {dismissPanel}
        isDataLoading = {isDataLoading} 
        showWeekends= {showWeekends} 
        onChkViewChange= {chkViewHandleChange}
      />

      {eventDetails &&
        <IDialog 
          hideDialog={!dialogState.dlgDetails} 
          toggleHideDialog={() => dialogDispatch({type: ACTIONS.EVENT_DETAILS_TOGGLE})}
          eventDetails={eventDetails}
          handleAddtoCal = {handleAddtoCal}
        />
      }

      <Panel
        isOpen={isOpenDetails}
        onDismiss={dismissPanelDetails}
        headerText={roomInfo ? roomInfo.Title : 'Room Details'}
        className={roomStyles.roomBookPanel}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        isBlocking={false}
        // isLightDismiss={true}
      >
        
        <IRoomDetails roomInfo={roomInfo} />
        <IPreloader isDataLoading = {isDataLoading} text = "" />
        <Dialog
            hidden={!dialogState.dlgDelete}
            onDismiss={() => dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE})}
            dialogContentProps={dialogContentProps}
            modalProps={modelProps}
        >
            <DialogFooter>
                <PrimaryButton onClick={() => onDeleteRoomClickHandler(roomInfo.Id)} text="Yes" />
                <DefaultButton onClick={() => dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE})} text="No" />
            </DialogFooter>
        </Dialog>

        <div className={styles.panelBtns}>
          {isUserManage &&
            <React.Fragment>
              <PrimaryButton className={styles.marginL10} onClick={() => onEditRoom(roomInfo.Id)} text="Edit" />
              <PrimaryButton className={styles.marginL10} onClick={() => dialogDispatch({type: ACTIONS.ROOM_DELETE_TOGGLE})} text="Delete" />
            </React.Fragment>
          }
          <DefaultButton className={styles.marginL10} onClick={dismissPanelDetails} text="Cancel" />
        </div>
      </Panel>
      <Panel
        isOpen={isOpenBook}
        className={roomStyles.roomBookPanel}
        onDismiss={dismissPanelBook}
        headerText={roomInfo && bookFormMode === 'New' ? roomInfo.Title : (eventDetailsRoom ? eventDetailsRoom.Room : '')}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        isBlocking={false}>
        <IRoomBook 
          formField = {formField}
          errorMsgField={errorMsgField} 
          periodOptions = {periods}
          eventDetailsRoom = {eventDetailsRoom}
          onChangeFormField={onChangeFormField}
          dismissPanelBook={dismissPanelBook}
          bookFormMode = {bookFormMode}
          onNewBookingClick={onNewBookingClickHandler}
          onEditBookingClick={onEditBookingClickHandler}
          onDeleteBookingClick={onDeleteBookingClickHandler}
          onUpdateBookingClick={onUpdateBookingClickHandler}
          roomInfo={roomInfo}
          isCreator = {isCreator}
          isPeriods = {false}
        >
          <MessageBar 
            className={roomStyles.guidelinesMsg}
            messageBarType={MessageBarType.warning}
            isMultiline={false}
            truncated={true}
            overflowButtonAriaLabel="See more"> 
            <IRoomGuidelines guidelines = {guidelines} /> 
          </MessageBar>
        </IRoomBook>
        <IPreloader isDataLoading = {isDataLoading} text = "" />
      </Panel>

      <Panel
        isOpen={isOpenMultiBook}
        onDismiss={dismissPanelMultiBook}
        headerText={'Multiple Booking'}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        isBlocking={false}
        type={PanelType.medium}>
        
        <IMultiBook 
          context = {props.context}
          formField = {formFieldMultiBk}
          errorMsgField = {errorMsgFieldMultiBk}
          onChangeFormField = {onChangeFormFieldMultiBk}
          schoolCategory = {schoolCategory}
          schoolNum = {schoolNum}
          schoolCycleOptions = {schoolCycleOptions}
          schoolCycleDayOptions = {cycleDays}
          periodOptions = {allPeriods}
          roomOptions = {rooms.map(room => ({key: room.Id, text: room.Title}))}
          cancelMultiBook = {cancelBookingClickHandler}
          checkBookingClick = {checkBookingClickHandler}
          bookingsGridVisible = {mergedBookings.length > 0 ? true : false}
        />

        {isCheckBookingClicked && mergedBookings.length === 0 &&
          <MessageBar messageBarType={MessageBarType.warning} className={styles.marginT20}>
            Sorry, there are no days to book in the selected range. Please modify your search criteria and try again.
          </MessageBar>
        }

        <IPreloader 
          isDataLoading = {isMultiBookingDataLoading} 
              text = "Please wait, loading bookings..."
        />
        {mergedBookings.length > 0 &&
          <IMultiBookList 
            bookingList = {mergedBookings} 
            updateBookings = {updateBookings}
            bookEventsHandler = {isConflict ? toggleConfirmMultiDlg : mulitBookEventsHandler}
            cancelBookingClickHandler = {cancelBookingClickHandler}
            isConflict = {isConflict}
          />   
        }

        <Dialog
          hidden={hideConfirmMultiDlg}
          onDismiss={toggleConfirmMultiDlg}
          dialogContentProps={{type: DialogType.largeHeader, title: 'Overwrite/Skip Bookings', subText: 'Are you sure you want to overwrite existing bookings or skip yours? If not, please click "No" and review your bookings once again.'}}
          modalProps={{isBlocking: false, styles: { main: { maxWidth: 450 }}}}>
          <DialogFooter>
              <PrimaryButton onClick={mulitBookEventsHandler} text="Yes" />
              <DefaultButton onClick={toggleConfirmMultiDlg} text="No" />
          </DialogFooter>
        </Dialog>

        <IPreloader isDataLoading = {isDataLoading} text = "" />
        
      </Panel>



    </div>
  );
  
  
}


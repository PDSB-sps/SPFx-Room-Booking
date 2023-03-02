import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient, IHttpClientOptions, MSGraphClient, SPHttpClient} from "@microsoft/sp-http";

import {formatStartDate, formatEndDate, getDatesWindow, formateTime} from '../Services/EventFormat';
import {parseRecurrentEvent} from '../Services/RecurrentEventOps';

export const calsErrs : any = [];

const resolveCalUrl = (context: WebPartContext, calType:string, calUrl:string, calName:string, currentDate: string) : string => {
    
    let resolvedCalUrl:string;
    let restApiUrl :string = "/_api/web/lists/getByTitle('"+calName+"')/items";
    let restApiParams :string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Category&$top=1000&$orderby=EndDate desc`;
    let restApiParamsRoom: string = "?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Status,AddToMyCal,RoomName/ColorCalculated,RoomName/ID,RoomName/Title,Periods/ID,Periods/EndTime,Periods/Title,Periods/StartTime&$expand=RoomName,Periods&$orderby=EventDate desc&$top=1000";

    const {dateRangeStart, dateRangeEnd} = getDatesWindow(currentDate);

    let restApiParamsWRange :string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Category&$top=1000&$orderby=EndDate desc&$filter=fRecurrence eq 1 or EventDate ge '${dateRangeStart.toISOString()}' and EventDate le '${dateRangeEnd.toISOString()}'`;
    let restApiParamsRoomWRange: string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Status,AddToMyCal,GraphID,RoomName/ColorCalculated,RoomName/ID,RoomName/Title,Periods/ID,Periods/EndTime,Periods/Title,Periods/StartTime&$expand=RoomName,Periods&$orderby=EventDate desc&$top=1000&$filter=fRecurrence eq 1 or EventDate ge '${dateRangeStart.toISOString()}' and EventDate le '${dateRangeEnd.toISOString()}'`;
    
    restApiParams = restApiParamsWRange;
    restApiParamsRoom = restApiParamsRoomWRange;

    switch (calType){
        case "Internal":
        case "Rotary":
            resolvedCalUrl = calUrl + restApiUrl + restApiParams;
            break;
        case "Room":
            resolvedCalUrl = calUrl + restApiUrl + restApiParamsRoom;
            break;
        case "My School":
            resolvedCalUrl = context.pageContext.web.absoluteUrl + restApiUrl + restApiParams;
            break;
        case "External":
            resolvedCalUrl = calUrl;
            break;
    }
    return resolvedCalUrl;
};

const getGraphCals = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, BgColorHex: string}, currentDate: string) : Promise <{}[]> => {	
    	
    let graphUrl :string = calSettings.CalURL.substring(32, calSettings.CalURL.length),	
        calEvents : {}[] = [];	
    
    const {dateRangeStart, dateRangeEnd} = getDatesWindow(currentDate);

    return new Promise <{}[]> (async(resolve, reject)=>{	
        context.msGraphClientFactory	
            .getClient()	
            .then((client :MSGraphClient)=>{	
                client	
                    .api(`${graphUrl}?$filter=start/dateTime ge '${dateRangeStart.toISOString()}' and start/dateTime le '${dateRangeEnd.toISOString()}'&$top=100`)
                    .header('Prefer','outlook.timezone="Eastern Standard Time"')	
                    .get((error, response: any, rawResponse?: any)=>{	
                        if(error){	
                            calsErrs.push("MS Graph Error - " + calSettings.Title);
                        }	
                        if(response){	
                            response.value.map((result:any)=>{	
                                calEvents.push({	
                                    id: result.id,	
                                    title: result.subject,	
                                    // start: formatStartDate(result.start.dateTime),	
                                    // end: formatStartDate(result.end.dateTime),	
                                    start: result.start.dateTime,	
                                    end: result.end.dateTime,	
                                    _location: result.location.displayName,	
                                    _body: result.body.content,
                                    className: "eventHidden",
                                    allDay: result.isAllDay,
                                    calendar: calSettings.Title,
                                    calendarColor: calSettings.BgColorHex
                                });	
                            });	
                        }	
                        resolve(calEvents);	
                    });	
            }, (error)=>{	
                calsErrs.push(error);
            });	
    });	
};

export const addToMyGraphCal = async (context: WebPartContext) =>{
    
    const event = {
        "subject": "Let's add this to my calendar",
        "body": {
            "contentType": "HTML",
            "content": "Adding a dummy event to my graph calendar"
        },
        "start": {
            "dateTime": "2021-02-15T12:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "end": {
            "dateTime": "2021-02-15T14:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "location": {
            "displayName": "Peel CBO"
        },
        "attendees": [{
            "emailAddress": {
                "address": "mai.mostafa@peelsb.com",
                "name": "Mai Mostafa"
            },
            "type": "required"
        }]
    };

    context.msGraphClientFactory
        .getClient()
        .then((client :MSGraphClient)=>{
            client
                .api("/me/events")
                .post(event, (err, res) => {
                    console.log(res);
                });
        });

};

export const getDefaultCals = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, BgColorHex: string}, currentDate: string) : Promise <{}[]> => {
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName, currentDate),
        calEvents : {}[] = [] ;

    const myOptions: IHttpClientOptions = {
        headers : { 
            'Accept': 'application/json;odata=verbose'
        }
    };

    try{
        const _data = await context.httpClient.get(calUrl, HttpClient.configurations.v1, myOptions);
            
        if (_data.ok){
            const calResult = await _data.json();
            if(calResult){
                calResult.d.results.map((result:any)=>{
                    calEvents.push({
                        id: result.ID,
                        title: result.Title,
                        start: result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate,
                        end: result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate,
                        allDay: result.fAllDayEvent,
                        _location: result.Location,
                        _body: result.Description,
                        recurr: result.fRecurrence,
                        recurrData: result.RecurrenceData,
                        rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, formatStartDate(result.EventDate), formatEndDate(result.EndDate)) : null,
                        className: "eventHidden",
                        calendar: calSettings.Title,
                        calendarColor: calSettings.BgColorHex,
                        GraphId: result.GraphID
                    });
                });
            }
        }else{
            calsErrs.push(calSettings.Title + ' - ' + _data.statusText);
            return [];
        }
    } catch(error){
        calsErrs.push("External calendars invalid - " + error);
    }
        
    return calEvents;
};

export const getRoomsCal = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}, currentDate: string, roomId?: number) : Promise <{}[]> => {
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName, currentDate),
        calEvents : {}[] = [] ;

    const myOptions: IHttpClientOptions = {
        headers : { 
            'Accept': 'application/json;odata=verbose'
        }
    };

    try{
        const _data = await context.httpClient.get(calUrl, HttpClient.configurations.v1, myOptions);
            
        if (_data.ok){
            const calResult = await _data.json();
            // console.log("calResult", calResult);
            if(calResult){
                calResult.d.results.map((result:any)=>{
                    calEvents.push({
                        id: result.ID,
                        title: result.Title,
                        start: result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate,
                        end: result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate,
                        allDay: result.fAllDayEvent,
                        _location: result.Location,
                        _body: result.Description,
                        recurr: result.fRecurrence,
                        recurrData: result.RecurrenceData,
                        rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, formatStartDate(result.EventDate), formatEndDate(result.EndDate)) : null,
                        color: result.RoomName.ColorCalculated,
                        roomId: result.RoomName.ID,
                        roomTitle: result.RoomName.Title,
                        className: roomId ? (roomId == parseInt(result.RoomName.ID) ? 'roomEvent roomID-' + result.RoomName.ID : 'roomEventHidden roomEvent roomID-' + result.RoomName.ID) : 'roomEvent roomID-' + result.RoomName.ID,
                        status: result.Status,
                        period: result.Periods.Title,
                        periodId: result.Periods.ID,
                        addToCal: result.AddToMyCal,
                        GraphId: result.GraphID
                    });
                });
            }
        }else{
            alert("Calendar Error: " + calSettings.Title + ' - ' + _data.statusText);
            return [];
        }
    } catch(error){
        alert("Calendar Error for external calendars - " + error);
    }
        
    return calEvents;
};

export const getExtCals = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, BgColorHex: string}, currentDate: string, spCalPageSize?: number) : Promise <{}[]> => {
    
    const {dateRangeStart, dateRangeEnd} = getDatesWindow(currentDate);

    let calUrl :string = `${calSettings.CalURL}&startdate=${dateRangeStart.toISOString()}&enddate=${dateRangeEnd.toISOString()}`;
    let calEvents : {}[] = [] ;

    try{
        const _data = await context.httpClient.get(calUrl, HttpClient.configurations.v1);
        if (_data.ok){
            const calResult = await _data.json();
            if(calResult){
                console.log("new external cal results", calResult);
                calResult.map((result:any)=>{
                    calEvents.push({
                        id: result.id,
                        title: result.title,
                        start: new Date(result.settings.startdate).toISOString(),
                        end: new Date(result.settings.enddate).toISOString(),
                        _startTime: formateTime(result.settings.startdate),
                        _endTime: formateTime(result.settings.enddate),
                        _body: result.content,
                        calendar: calSettings.Title,
                        calendarColor: calSettings.BgColorHex,
                        allDay: false,
                        _location: null,
                        recurr: false,
                        className: "eventHidden"
                        
                    });
                });
                console.log("formatted new ext calEvents", calEvents);
            }
        }
    } catch(error){
        calsErrs.push("New External calendars invalid - " + error);
    }
    return calEvents;
};

export const getCalsData = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, BgColorHex: string}, currentDate: string, roomId?: number) : Promise <{}[]> => {
    if(calSettings.CalType == 'Graph'){
        return getGraphCals(context, calSettings, currentDate);
    }else if(calSettings.CalType == 'Room'){
        return getRoomsCal(context, calSettings, currentDate, roomId);
    }else if ( calSettings.CalType == 'External'){
        return getExtCals(context, calSettings, currentDate);
    }else{
        return getDefaultCals(context, calSettings, currentDate);
    }
};

export const reRenderCalendars = (calEventSources: any, calVisibility: {calId: string, calChk: boolean}) =>{
    const newCalEventSources = calEventSources.map((eventSource: any) => {
        if (eventSource.calId == calVisibility.calId) {
            const updatedEventSource = {...eventSource}; //shallow clone
            updatedEventSource.events = eventSource.events.map((event: any) => {
                event['className'] = !calVisibility.calChk ? 'eventHidden' : '';
                return event;
            });
            return updatedEventSource;
        } else {
            return {...eventSource}; //shallow clone
        }
    });
    return newCalEventSources;
};
export const getLegendChksState = (calsVisibilityState: any, calVisibility: any) => {
    const calsVisibilityArr = calsVisibilityState;
    if (calsVisibilityArr.filter(i => i.calId === calVisibility.calId).length === 0 ){
        calsVisibilityArr.push(calVisibility);
    }else{
        calsVisibilityArr.map(i=> i.calId == calVisibility.calId ? i.calChk = calVisibility.calChk : '' );
    }
    return calsVisibilityArr;
};

export const getMySchoolCalGUID = async (context: WebPartContext, calSettingsListName: string) =>{
    const calSettingsRestUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${calSettingsListName}')/items?$filter=CalType eq 'My School'&$select=CalName`;
    const calSettingsCall = await context.spHttpClient.get(calSettingsRestUrl, SPHttpClient.configurations.v1).then(response => response.json());
    const calName = calSettingsCall.value[0].CalName;

    const calRestUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${calName}')?$select=id`;
    const calCall = await context.spHttpClient.get(calRestUrl, SPHttpClient.configurations.v1).then(response => response.json());
    
    return calCall.Id;
};
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {HttpClientResponse, HttpClient, IHttpClientOptions, MSGraphClient, SPHttpClient} from "@microsoft/sp-http";

import {formatStartDate, formatEndDate} from '../Services/EventFormat';
import {parseRecurrentEvent} from '../Services/RecurrentEventOps';

export const calsErrs : any = [];

const resolveCalUrl = (context: WebPartContext, calType:string, calUrl:string, calName:string) : string => {
    let resolvedCalUrl:string,
        azurePeelSchoolsUrl :string = "https://pdsb1.azure-api.net/peelschools",
        restApiUrl :string = "/_api/web/lists/getByTitle('"+calName+"')/items",
        restApiParams :string = "?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData&$orderby=EventDate desc&$top=300",
        restApiParamsRoom: string = "?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Status,AddToMyCal,RoomName/ColorCalculated,RoomName/ID,RoomName/Title,Periods/ID,Periods/EndTime,Periods/Title,Periods/StartTime&$expand=RoomName,Periods&$orderby=EventDate desc&$top=1000";
        //restApiParams :string = "?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData&$filter=EventDate ge datetime'2019-08-01T00%3a00%3a00'";
        //$filter=EventDate ge datetime'2019-08-01T00%3a00%3a00'

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
            resolvedCalUrl = azurePeelSchoolsUrl + calUrl.substring(calUrl.indexOf('.org/') + 12, calUrl.length) + restApiUrl + restApiParams;
            break;
    }
    return resolvedCalUrl;
};

const getGraphCals = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}) : Promise <{}[]> => {	
    	
    let graphUrl :string = calSettings.CalURL.substring(32, calSettings.CalURL.length),	
        calEvents : {}[] = [];	
    return new Promise <{}[]> (async(resolve, reject)=>{	
        context.msGraphClientFactory	
            .getClient()	
            .then((client :MSGraphClient)=>{	
                client	
                    .api(graphUrl)	
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
                                    allDay: result.isAllDay
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

const getDefaultCals1 = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}) : Promise <{}[]> =>{
    
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName),
        calEvents : {}[] = [] ;

    const myOptions: IHttpClientOptions = {
        headers : { 
            'Accept': 'application/json;odata=verbose'
        }
    };

    console.log("calURL", calUrl);

    return new Promise <{}[]> (async(resolve, reject) =>{
        context.httpClient
            .get(calUrl, HttpClient.configurations.v1, myOptions)
            .then((response: HttpClientResponse) =>{
                response.json().then((results:any)=>{
                    results.d.results.map((result:any)=>{
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
                            rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, formatStartDate(result.EventDate), formatEndDate(result.EndDate)) : null
                        });
                    });
                    resolve(calEvents);
                });
            }).catch((error:any)=>{
                resolve([]);
                console.log("Calendar URL error!");
            });
    });
    
};

export const getDefaultCals = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}) : Promise <{}[]> => {
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName),
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
                        className: "eventHidden"
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

export const getRoomsCal = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}, roomId?: number) : Promise <{}[]> => {
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName),
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
                        addToCal: result.AddToMyCal
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

export const getCalsData = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}, roomId?: number) : Promise <{}[]> => {
    if(calSettings.CalType == 'Graph'){
        return getGraphCals(context, calSettings);
    }else if(calSettings.CalType == 'Room'){
        return getRoomsCal(context, calSettings, roomId);
    }else{
        return getDefaultCals(context, calSettings);
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
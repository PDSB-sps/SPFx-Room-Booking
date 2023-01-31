import { WebPartContext } from "@microsoft/sp-webpart-base";
import {HttpClient, IHttpClientOptions, MSGraphClient, SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";
import {addToMyGraphCal} from './RoomOperations';
import {formatStartDate, formatEndDate} from '../Services/EventFormat';
import * as moment from 'moment';

export const calsErrs : any = [];

export const getAllPeriods = async (context: WebPartContext, periodsList: string) =>{
    //console.log("Get All Periods Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${periodsList}')/items?$orderBy=SortOrder asc`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    return results.value.map(item => (
        {
            key: item.Id,
            text: item.Title + '  (' + moment(item.StartTime).format('hh:mm A') + ' - ' + moment(item.EndTime).format('hh:mm A') + ')',
            start: item.StartTime,
            end: item.EndTime,
        }
    ));
};

export const getSchoolCategory = (calUrl:string) => { // elementary or secondary
    calUrl = "https://pdsb1.sharepoint.com/sites/Rooms/1234/"; // for testing
    calUrl = calUrl.toLowerCase();
    const schoolLoc = calUrl.substring(calUrl.indexOf('/rooms/')+7).replace("/","");
    const schoolLocNum = Number(schoolLoc);
    if (schoolLocNum){
        if (schoolLocNum >= 1000 && schoolLocNum <= 2000) return {schoolNum: schoolLoc, schoolCategory: 'Elem'};
        if (schoolLocNum >= 2001 && schoolLocNum <= 3000) return {schoolNum: schoolLoc, schoolCategory: 'Sec'};
    }
    return {schoolNum: schoolLoc, schoolCategory: 'None'};
};

export const getSchoolCycles = async (context: WebPartContext, schoolLocNum: string) => {
    const restUrl = `https://pdsb1.sharepoint.com/sites/Rooms` + `/_api/web/lists/getByTitle('CalendarSettings')/items?$filter=CalName eq '${schoolLocNum}' and CalType eq 'Graph'`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());
    //console.log("school Cycles", results.value);
    return {cycleDays: results.value[0].CycleDays, calUrl: results.value[0].CalURL};
};

export const getGraphCalsMultiBook = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}, startDate: string, endDate: string, cycleDay: string) : Promise <{}[]> => {	
    let graphUrl :string = calSettings.CalURL.substring(32, calSettings.CalURL.length),	
        calEvents : {}[] = [];	
    
    startDate = startDate.substring(0, startDate.indexOf('T')) + "T00:00:00.0000000";
    endDate = endDate.substring(0, endDate.indexOf('T')) + "T00:00:00.0000000";
    
    return new Promise <{}[]> (async(resolve, reject)=>{	
        context.msGraphClientFactory	
            .getClient()	
            .then((client :MSGraphClient)=>{	
                client	
                    .api(graphUrl)	
                    .filter(`subject eq '${cycleDay}' and start/dateTime ge '${startDate}' and start/dateTime le '${endDate}'`)
                    .top(500)
                    .orderby('start/dateTime')
                    .header('Prefer','outlook.timezone="Eastern Standard Time"')	
                    .get((error, response: any, rawResponse?: any)=>{	
                        if(error){	
                            calsErrs.push("MS Graph Error - " + calSettings.Title);
                        }	
                        if(response){	
                            response.value.map((result:any)=>{	
                                calEvents.push({	
                                    title: result.subject,	
                                    start: moment(result.start.dateTime).format('YYYY-MM-DD'),	
                                    end: moment(result.end.dateTime).format('YYYY-MM-DD'),	
                                    // _body: result.body.content,
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

export const getBookedEvents = 
    async (
        context: WebPartContext, 
        calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}, 
        roomId: string,
        periodId: string,
        startDate: string,
        endDate: string
        ) : Promise <{}[]> => {
    
    startDate = startDate.substring(0, startDate.indexOf('T')) + "T00:00:00.0000000";
    endDate = endDate.substring(0, endDate.indexOf('T')) + "T00:00:00.0000000";

    const restApiUrl :string = "/_api/web/lists/getByTitle('"+calSettings.CalName+"')/items";
    //const restApiParamsRoom: string = `?$select=ID,Title,Author/EMail,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Status,AddToMyCal,RoomName/ColorCalculated,RoomName/ID,RoomName/Title,Periods/ID,Periods/EndTime,Periods/Title,Periods/StartTime&$expand=RoomName,Periods,Author&$filter=RoomName/ID eq '${roomId}' and Periods/ID eq '${periodId}' and EventDate ge '${startDate}' and EventDate le '${endDate}'&$orderby=EventDate desc&$top=1000`;
    const restApiParamsRoom: string = `?$select=ID,Title,Author/EMail,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Status,AddToMyCal,RoomName/ColorCalculated,RoomName/ID,RoomName/Title,Periods/ID,Periods/EndTime,Periods/Title,Periods/StartTime&$expand=RoomName,Periods,Author&$filter=RoomName/ID eq '${roomId}' and EventDate ge '${startDate}' and EventDate le '${endDate}'&$orderby=EventDate desc&$top=1000`;
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Events')/items` + restApiParamsRoom;
    
    let calEvents : {}[] = [] ;

    const myOptions: IHttpClientOptions = {
        headers : { 
            'Accept': 'application/json;odata=verbose'
        }
    };

    try{
        const _data = await context.httpClient.get(restUrl, HttpClient.configurations.v1, myOptions);
            
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
                        roomId: result.RoomName.ID,
                        roomTitle: result.RoomName.Title,
                        period: result.Periods.Title,
                        periodId: result.Periods.ID,
                        addToCal: result.AddToMyCal,
                        author: result.Author.EMail
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

export const mergeBookings = (existingBookings, multiBookings, multiBookingsFields) => {
    console.log("existingBookings", existingBookings);
    console.log("multiBookings", multiBookings);
    let isConflictBool = false;

    const mergedBookingsList = [];
    let isConflict :boolean;
    for (let i=0; i< multiBookings.length; i++){
        isConflict = false;
        for (let existingBooking of existingBookings){
            let bookingStartDate = multiBookings[i].start;
            let existingBookingStartDate = existingBooking.start.substring(0, existingBooking.start.indexOf('T'));
            if (bookingStartDate === existingBookingStartDate){
                // console.log("bookingStartDate === existingBookingStartDate", bookingStartDate === existingBookingStartDate);
                // console.log("isPeriodConflict fnc", isPeriodConflict(existingBooking, multiBookingsFields.periodField));
                if (isPeriodConflict(existingBooking, multiBookingsFields.periodField)){
                    isConflict = true;
                    mergedBookingsList.push({
                        title: multiBookingsFields.titleField,
                        description: multiBookingsFields.descpField,
                        room: multiBookingsFields.roomField,
                        period: multiBookingsFields.periodField,
                        start: multiBookings[i].start,
                        end: multiBookings[i].end,
                        index: i,
                        conflict : true,
                        overwrite  : false,
                        conflictTitle : existingBooking.title,
                        conflictAuthor : existingBooking.author,
                        conflictId : existingBooking.id,
                    });
                isConflictBool = true;
                }
            }
        } 
        if(!isConflict){
            mergedBookingsList.push({
                title: multiBookingsFields.titleField,
                description: multiBookingsFields.descpField,
                room: multiBookingsFields.roomField,
                period: multiBookingsFields.periodField,
                start: multiBookings[i].start,
                end: multiBookings[i].end,
                index: i,
                conflict: false,
                overwrite: true,
                conflictTitle : null,
                conflictAuthor : null,
                conflictId : null,
            });
        }
    }

    return {isConflictBool, mergedBookingsList};
};

const isPeriodConflict = (period1, period2) => {
    const period1Start = moment(period1.start).format('HHmm').toString();
    const period1End = moment(period1.end).format('HHmm').toString();
    const period2Start = moment(period2.start).format('HHmm').toString();
    const period2End = moment(period2.end).format('HHmm').toString();
    // console.log("period1Start", "period1End", "period2Start", "period2End");
    // console.log(period1Start, period1End, period2Start, period2End);
    if (
        period1Start >= period2Start && period1Start <= period2End ||
        period1End >= period2Start && period1End <= period2End ||
        period1Start <= period2Start && period1End >= period2End 
     ){
        return true;
    }
    return false;
}

export const addBooking = async (context: WebPartContext, roomsCalListName: string, eventDetails: any, roomInfo: any) => {
    console.log("roomInfo", roomInfo);
    console.log("eventDetails", eventDetails);
    const periodStartTime = eventDetails.periodField.start;
    const periodEndTime = eventDetails.periodField.end;
    
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${roomsCalListName}')/items`;
    const body: string = JSON.stringify({
        Title: eventDetails.titleField,
        Description: eventDetails.descpField,
        EventDate: eventDetails.dateField + periodStartTime.substring(periodStartTime.indexOf('T')),
        EndDate: eventDetails.dateField + periodEndTime.substring(periodEndTime.indexOf('T')),
        PeriodsId: eventDetails.periodField.key,
        RoomNameId: roomInfo.Id,
        Location: roomInfo.Title,
        AddToMyCal: eventDetails.addToCalField
    });
    const spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": ""
        },
        body: body
    };
    const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    if(_data.ok){
        console.log('New Event is added!');
    }

    if(eventDetails.addToCalField){
        addToMyGraphCal(context, eventDetails, roomInfo).then(()=>{
            console.log('Room added to My Calendar!');
        });
    }
};

import {WebPartContext} from "@microsoft/sp-webpart-base";
import { SPPermission } from "@microsoft/sp-page-context";
import {SPHttpClient, ISPHttpClientOptions, MSGraphClient} from "@microsoft/sp-http";
import {formatStartDate, formatEndDate} from '../Services/EventFormat';
import * as moment from 'moment';

export const getRooms = async (context: WebPartContext, roomsList: string) =>{
    console.log("Get Rooms Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${roomsList}')/items`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    return results.value;
};

const adjustLocation = (arr: []): {}[] =>{
    let arrAdj :{}[] = [];
    arrAdj.push({key: 'all', text:'All'});

    arr.map((item: string)=>{
        arrAdj.push({
            key: item.toLowerCase(),
            text: item
        });
    });

    return arrAdj;
};
export const getLocationGroup = async(context: WebPartContext, roomsList: string) =>{
    console.log("Get Rooms Location Group Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${roomsList}')/fields?$filter=EntityPropertyName eq 'LocationGroup'`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    return adjustLocation(results.value[0].Choices);
};


const getCanBookPeriods = (period: any, allPeriods: any) =>{  
    console.log("allPeriods", allPeriods);
    for (let i in allPeriods){
      if (
            // moment(period.start) >= moment(allPeriods[i].start) && moment(period.start) <= moment(allPeriods[i].end) ||
            // moment(period.end) >= moment(allPeriods[i].start) && moment(period.end) <= moment(allPeriods[i].end) ||
            // moment(period.start) <= moment(allPeriods[i].start) && moment(period.end) >= moment(allPeriods[i].end)
            moment(period.start).isSameOrAfter(moment(allPeriods[i].start)) && moment(period.start).isSameOrBefore(moment(allPeriods[i].end)) ||
            moment(period.end).isSameOrAfter(moment(allPeriods[i].start)) && moment(period.end).isSameOrBefore(moment(allPeriods[i].end)) ||
            moment(period.start).isSameOrBefore(moment(allPeriods[i].start)) && moment(period.end).isSameOrAfter(moment(allPeriods[i].end))
         ){
            console.log("period.text", period.text);
            console.log("moment(period.start) >= moment(allPeriods[i].start)", moment(period.start) >= moment(allPeriods[i].start));
            console.log("moment(period.start) <= moment(allPeriods[i].end)", moment(period.start) <= moment(allPeriods[i].end));
            allPeriods[i].disabled = true;
        }
    } 
    console.log("updatedPeriods", allPeriods);
    return allPeriods;
};
const updatePeriods = (allPeriods: any) => {
    let bookedPeriods: any = [], updatedPeriods : any = [];
    bookedPeriods = allPeriods.filter((period: any) => period.disabled); 
    console.log("bookedPeriods", bookedPeriods);
    if (bookedPeriods.length === 0){
        return allPeriods;
    }
    /*for (let j=0; j<bookedPeriods.length; j++){
        updatedPeriods = getCanBookPeriods(bookedPeriods[j], allPeriods);
    }*/
    updatedPeriods = bookedPeriods.map((bookedPeriod: any) => getCanBookPeriods(bookedPeriod, allPeriods));
    return updatedPeriods.flat();
};

const adjustPeriods = (arr: [], disabledPeriods: any): {}[] =>{
    let arrAdj :{}[] = [];

    arr.map((item: any)=>{
        arrAdj.push({
            key: item.Id,
            text: item.Title + '  (' + moment(item.StartTime).format('hh:mm A') + ' - ' + moment(item.EndTime).format('hh:mm A') + ')',
            start: item.StartTime,
            end: item.EndTime,
            //order: item.SortOrder,
            disabled: disabledPeriods.includes(item.Id) ? true : false
        });
    });

    return updatePeriods(arrAdj);
};
export const getPeriods = async (context: WebPartContext, periodsList: string, roomId: any, bookingDate: any) =>{
    console.log("Get Periods Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${periodsList}')/items?$orderBy=SortOrder asc`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    const restUrlEvents = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Events')/items?$filter=RoomNameId eq '${roomId}'`;
    const resultsEvents = await context.spHttpClient.get(restUrlEvents, SPHttpClient.configurations.v1).then(response => response.json());
    
    let bookedPeriods : any = [];
    let bookingDateDay = moment(bookingDate).format('MM-DD-YYYY');
    for (let resultEvent of resultsEvents.value){
        if(moment(resultEvent.EventDate).format('MM-DD-YYYY') === bookingDateDay){
            bookedPeriods.push(resultEvent.PeriodsId);
        }
    }

    return adjustPeriods(results.value, bookedPeriods);
};

// export const getFreePeriods  =  async (context: WebPartContext, periods: any, selectedDate: any, roomId: any) =>{
//     const restUrlEvents = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Events')/items?$filter=RoomNameId eq '${roomId}'`;
//     const resultsEvents = await context.spHttpClient.get(restUrlEvents, SPHttpClient.configurations.v1).then(response => response.json());
    
//     let bookedPeriods : any = [];
//     let bookingDateDay = moment(selectedDate).format('MM-DD-YYYY');
//     for (let resultEvent of resultsEvents.value){
//         if(moment(resultEvent.EventDate).format('MM-DD-YYYY') === bookingDateDay){
//             bookedPeriods.push(resultEvent.PeriodsId);
//         }
//     }

//     let updatedPeriods : any = [];
//     periods.map((period: any)=>{
//         updatedPeriods.push({
//             key: period.key,
//             text: period.text,
//             start: period.start,
//             end: period.end,
//             order: period.order,
//             disabled: bookedPeriods.includes(period.key) ? true : false
//         });
//     });

//     return updatedPeriods;
// };

export const getGuidelines = async (context: WebPartContext, guidelinesList: string) =>{
    console.log("Get Guidelines Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${guidelinesList}')/items`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    return results.value;
};

export const getRoomsCalendarName = (calendarSettingsList: any) : string =>{
    for (let calSetting of calendarSettingsList){
        if (calSetting.CalType === 'Room'){
            return calSetting.CalName;
        }
    }
    return 'Events';
};

export const getChosenDate = (startPeriodField: any, endPeriodField: any, formFieldParam: any) =>{
    const startPeriod = new Date(startPeriodField);
    const endPeriod = new Date(endPeriodField);
    const currDate = new Date(formFieldParam);

    const startPeriodHr = startPeriod.getHours();
    const startPeriodMin = startPeriod.getMinutes();
    const endPeriodHr = endPeriod.getHours();
    const endPeriodMin = endPeriod.getMinutes();

    const dateDay = currDate.getDate();
    const dateMonth = currDate.getMonth();
    const dateYear = currDate.getFullYear();

    let chosenStartDate = new Date();
    chosenStartDate.setDate(dateDay);
    chosenStartDate.setMonth(dateMonth);
    chosenStartDate.setFullYear(dateYear);
    chosenStartDate.setHours(startPeriodHr);
    chosenStartDate.setMinutes(startPeriodMin);

    let chosenEndDate = new Date();
    chosenEndDate.setDate(dateDay);
    chosenEndDate.setMonth(dateMonth);
    chosenEndDate.setFullYear(dateYear);
    chosenEndDate.setHours(endPeriodHr);
    chosenEndDate.setMinutes(endPeriodMin);

    return[chosenStartDate, chosenEndDate];
};

export const addToMyGraphCal = async (context: WebPartContext, eventDetails: any, roomInfo: any) =>{
    const event = {
        "subject": eventDetails.titleField,
        "body": {
            "contentType": "HTML",
            "content": eventDetails.descpField
        },
        "start": {
            "dateTime": getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[0],
            "timeZone": "Eastern Standard Time"
        },
        "end": {
            "dateTime": getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[1],
            "timeZone": "Eastern Standard Time"
        },
        "location": {
            "displayName": roomInfo.LocationGroup +' - '+ roomInfo.Title + ', ' + eventDetails.periodField.text
        }
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

export const addEvent = async (context: WebPartContext, roomsCalListName: string, eventDetails: any, roomInfo: any) => {
    //console.log("roomInfo", roomInfo);
    //console.log("eventDetails", eventDetails);
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${roomsCalListName}')/items`;
    const body: string = JSON.stringify({
        Title: eventDetails.titleField,
        Description: eventDetails.descpField,
        EventDate: getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[0],
        EndDate: getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[1],
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

export const deleteItem = async (context: WebPartContext, listName: string, itemId: any) => {
    const restUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})`;
    let spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"         
        },
    };

    const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('Item deleted!');
    }
};

export const updateEvent = async (context: WebPartContext, roomsCalListName: string, eventId: any, eventDetails: any, roomInfo: any) => {
    const restUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${roomsCalListName}')/items(${eventId})`,
    body: string = JSON.stringify({
        Title: eventDetails.titleField,
        Description: eventDetails.descpField,
        EventDate: getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[0],
        EndDate: getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[1],
        PeriodsId: eventDetails.periodField.key,
        RoomNameId: roomInfo.Id,
        Location: roomInfo.Title,
        AddToMyCal: eventDetails.addToCalField
    }),
    spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",    
        },
        body: body
    },
    _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    
    if (_data.ok){
        console.log('Event Booking is updated!');
    }
};

export const isEventCreator = async (context: WebPartContext, roomsCalListName: string, eventId: any) =>{
    const currUserId = context.pageContext.legacyPageContext["userId"];

    const restUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${roomsCalListName}')/items(${eventId})?$select=AuthorId`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    return currUserId == results.AuthorId;
};
export const isUserManage = (context: WebPartContext) : boolean =>{
    const userPermissions = context.pageContext.web.permissions,
        permission = new SPPermission (userPermissions.value);
    
    return permission.hasPermission(SPPermission.manageWeb);
};
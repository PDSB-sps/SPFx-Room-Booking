import {WebPartContext} from "@microsoft/sp-webpart-base";
import { SPPermission } from "@microsoft/sp-page-context";
import {SPHttpClient, ISPHttpClientOptions, MSGraphClient} from "@microsoft/sp-http";
import * as moment from 'moment';

export const getRooms = async (context: WebPartContext, roomsList: string) =>{
    console.log("Get Rooms Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${roomsList}')/items?$orderby=SortOrder asc`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());
    // console.log("rooms", results.value);
    return results.value;
};
export const getRoomInfo = async (context: WebPartContext, roomsList: string, roomId: string) => {
    console.log("Get Room Info Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${roomsList}')/items?$filter=Id eq ${roomId}`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    return results.value[0];
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
    const periodStart = moment(period.start).format('HHmm').toString();
    const periodEnd = moment(period.end).format('HHmm').toString();
    for (let i in allPeriods){
        const thisPeriodStart = moment(allPeriods[i].start).format('HHmm').toString();
        const thisPeriodEnd = moment(allPeriods[i].end).format('HHmm').toString();
      if (
            periodStart >= thisPeriodStart && periodStart <= thisPeriodEnd ||
            periodEnd >= thisPeriodStart && periodEnd <= thisPeriodEnd ||
            periodStart <= thisPeriodStart && periodEnd >= thisPeriodEnd 
         ){
            allPeriods[i].disabled = true;
        }
    } 
    return allPeriods;
};
const updatePeriods = (allPeriods: any) => {
    let bookedPeriods: any = [], updatedPeriods : any = [];
    bookedPeriods = allPeriods.filter((period: any) => period.disabled); 
    if (bookedPeriods.length === 0){
        return allPeriods;
    }
    for (let j=0; j<bookedPeriods.length; j++){
        updatedPeriods = getCanBookPeriods(bookedPeriods[j], allPeriods);
    }
    return updatedPeriods;
    
    //updatedPeriods = bookedPeriods.map((bookedPeriod: any) => getCanBookPeriods(bookedPeriod, allPeriods));
    //return updatedPeriods.flat();

    // const uniqueUpdatedPeriods = updatedPeriods.flat().filter((v,i,a)=>a.findIndex(t=>(t.key === v.key))===i);
    // return uniqueUpdatedPeriods;
};

const adjustPeriods = (arr: [], disabledPeriods: any): {}[] =>{
    let arrAdj :{}[] = [];

    // console.log("disabledPeriods", disabledPeriods);
    // console.log("selectedPeriod", selectedPeriod);

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
    //return arrAdj;
};
export const getPeriods = async (context: WebPartContext, periodsList: string, roomId: any, bookingDate: any, selectedPeriod?: any) =>{
    console.log("Get Periods Function");
    const restUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${periodsList}')/items?$orderBy=SortOrder asc`;
    const results = await context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1).then(response => response.json());

    const restUrlEvents = context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Events')/items?$filter=RoomNameId eq '${roomId}'&$top=600`;
    const resultsEvents = await context.spHttpClient.get(restUrlEvents, SPHttpClient.configurations.v1).then(response => response.json());
    
    let bookedPeriods : any = [];
    let bookingDateDay = moment(bookingDate).format('MM-DD-YYYY');
    for (let resultEvent of resultsEvents.value){
        if(moment(resultEvent.EventDate).format('MM-DD-YYYY') === bookingDateDay && resultEvent.PeriodsId !== selectedPeriod){
            bookedPeriods.push(resultEvent.PeriodsId);
        }
    }

    // console.log("resultsEvents.value", resultsEvents.value)
    // console.log("bookedPeriods", bookedPeriods);
    // console.log("selectedPeriod", selectedPeriod);
    console.log("adjustPeriods", adjustPeriods(results.value, bookedPeriods));

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
    
    // console.log("formFieldParam", formFieldParam)
    // console.log("currDate", currDate)

    const startPeriodHr = startPeriod.getHours();
    const startPeriodMin = startPeriod.getMinutes();
    const endPeriodHr = endPeriod.getHours();
    const endPeriodMin = endPeriod.getMinutes();

    const dateDay = currDate.getDate();
    const dateMonth = currDate.getMonth();
    const dateYear = currDate.getFullYear();

    // console.log("dateDay", dateDay);
    // console.log("dateMonth", dateMonth);
    // console.log("dateYear", dateYear);

    //let chosenStartDate = new Date();
    // chosenStartDate.setDate(dateDay);
    // chosenStartDate.setMonth(dateMonth);
    // chosenStartDate.setFullYear(dateYear);
    // chosenStartDate.setHours(startPeriodHr);
    // chosenStartDate.setMinutes(startPeriodMin);

    // let chosenEndDate = new Date();
    // chosenEndDate.setDate(dateDay);
    // chosenEndDate.setMonth(dateMonth);
    // chosenEndDate.setFullYear(dateYear);
    // chosenEndDate.setHours(endPeriodHr);
    // chosenEndDate.setMinutes(endPeriodMin);

    // A fix for the Feb-Mar issue - the date was setting the day first in Feb which only have 28 days. So wasn't working for 29,30,31 days of the month
    let chosenStartDate = new Date(dateYear, dateMonth, dateDay);
    chosenStartDate.setHours(startPeriodHr);
    chosenStartDate.setMinutes(startPeriodMin);

    let chosenEndDate = new Date(dateYear, dateMonth, dateDay);
    chosenEndDate.setHours(endPeriodHr);
    chosenEndDate.setMinutes(endPeriodMin);

    // console.log("[chosenStartDate, chosenEndDate]", [chosenStartDate, chosenEndDate]);

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
    // console.log("roomInfo", roomInfo);
    // console.log("eventDetails", eventDetails);
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
    const restUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/recycle`;
    let spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            // "X-HTTP-Method": "DELETE"         
        },
    };

    const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('Item is deleted! Please check Recycle Bin to restore it.');
    }
};

export const updateEvent = async (context: WebPartContext, roomsCalListName: string, eventId: any, eventDetails: any, eventDetailsRoom: any) => {
    const restUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${roomsCalListName}')/items(${eventId})`,
    body: string = JSON.stringify({
        Title: eventDetails.titleField,
        Description: eventDetails.descpField,
        EventDate: getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[0],
        EndDate: getChosenDate(eventDetails.periodField.start, eventDetails.periodField.end, eventDetails.dateField)[1],
        PeriodsId: eventDetails.periodField.key,
        RoomNameId: eventDetailsRoom.RoomId,
        Location: eventDetailsRoom.Room,
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
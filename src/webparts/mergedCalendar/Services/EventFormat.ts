// import * as moment from 'moment';
import * as moment from 'moment-timezone'; 

export const formateDate = (ipDate:any) :any => {
    //return moment(ipDate).format('YYYY-MM-DD hh:mm A'); 
    return moment.tz(ipDate, "America/Toronto").format('YYYY-MM-DD hh:mm A');
};

// only for user's view in the event details dialog
export const formateTime = (ipDate:any) :any => {
    return moment.tz(ipDate, "America/Toronto").format('YYYY-MM-DD hh:mm A');
};

export const formatStartDate = (ipDate:any) : any => {
    let startDateMod = new Date(ipDate);
    startDateMod.setTime(startDateMod.getTime());
    
    return moment.utc(startDateMod).format('YYYY-MM-DD') + "T" + moment.utc(startDateMod).format("hh:mm") + ":00Z";
};

export const formatEndDate = (ipDate:any) :any => {
    let endDateMod = new Date(ipDate);
    endDateMod.setTime(endDateMod.getTime());

    let nextDay = moment(endDateMod).add(1, 'days');
    return moment.utc(nextDay).format('YYYY-MM-DD') + "T" + moment.utc(nextDay).format("hh:mm") + ":00Z";
};

export const formatStrHtml = (str: string) : any => {
    let parser = new DOMParser();
    let htmlEl = parser.parseFromString(str, 'text/html');
    //console.log(htmlEl.body);
    return htmlEl.body;
};

export const formatEvDetails = (ev:any) : {} =>{
    let event = ev.event,
        evDetails : {} = {};

    evDetails = {
        EventId: event._def.publicId,
        Title: event.title,
        Start: event.startStr ? formateDate(event.startStr) : "",
        End: event.endStr ? formateDate(event.endStr) : "",
        // Start: event._def.extendedProps._startTime ? event._def.extendedProps._startTime : "",
        // End: event._def.extendedProps._endTime ? event._def.extendedProps._endTime : "",
        Location: event._def.extendedProps._location,
        Body: event._def.extendedProps._body ? event._def.extendedProps._body : null,
        AllDay: event.allDay,
        Recurr: event._def.extendedProps.recurr,
        RecurrData: event._def.extendedProps.recurrData,
        RecurringDef: event._def.extendedProps.recurringDef,
        Room: event._def.extendedProps.roomTitle,
        RoomId: event._def.extendedProps.roomId,
        Status: event._def.extendedProps.status,
        Period: event._def.extendedProps.period,
        PeriodId: event._def.extendedProps.periodId,
        RoomColor: event.backgroundColor,
        AddToMyCal: event._def.extendedProps.addToCal,
        Calendar: event._def.extendedProps.calendar,
        Color: event._def.extendedProps.calendarColor,
        GraphId: event._def.extendedProps.graphId
    };

    return evDetails;
};

export const getDatesWindow = (currentDate: string) => {
    const currentDateVal = new Date (currentDate);
    let dateRangeStart = new Date (currentDate), dateRangeEnd = new Date (currentDate);
    if (currentDateVal.getMonth() === 0){
        dateRangeStart.setMonth(11);
        dateRangeStart.setFullYear(currentDateVal.getFullYear()-1);
    }else{
        dateRangeStart.setMonth(currentDateVal.getMonth()-3);
    }
    if(currentDateVal.getMonth() === 11){
        dateRangeEnd.setMonth(0);
        dateRangeEnd.setFullYear(currentDateVal.getFullYear()+1);
    }else{
        dateRangeEnd.setMonth(currentDateVal.getMonth()+3);
    }

    // console.log("resolveCalUrl current currentDate", currentDate);
    // console.log("currentDate", new Date(currentDate));
    // console.log("dateRangeStart", dateRangeStart);
    // console.log("dateRangeEnd", dateRangeEnd);

    return {dateRangeStart, dateRangeEnd};
};


import * as React from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import rrulePlugin from '@fullcalendar/rrule';

import styles from '../MergedCalendar.module.scss';
import {ICalendarProps} from './ICalendarProps';

import {isUserManage} from '../../Services/WpProperties';

export default function ICalendar(props:ICalendarProps){

  const calendarRef = React.useRef<any>();
  
  const calendarNext = () => {
    let calendarApi = calendarRef.current.getApi();
    calendarApi.next();
  };
  const calendarPrev = () => {
    let calendarApi = calendarRef.current.getApi();
    calendarApi.prev();
  };
    
    return(
        <div className={styles.calendarCntnr}>
          <FullCalendar 
            ref={calendarRef}
            plugins = {
              [dayGridPlugin, timeGridPlugin, interactionPlugin, rrulePlugin]
            }
            headerToolbar = {{
              //left: 'prev,next today',
              left: 'customPrev,customNext today',
              center: 'title',
              right: isUserManage(props.context) ? 'dayGridMonth,timeGridWeek,timeGridDay settingsBtn' : 'dayGridMonth,timeGridWeek,timeGridDay' 
            }}
            customButtons = {{
              settingsBtn : {
                text : 'Settings',
                click : props.openPanel,
              },
              addEventBtn : {
                text: 'Add Event',
                click : ()=>{
                  window.open(
                    props.context.pageContext.web.absoluteUrl + '/_layouts/15/Event.aspx?ListGuid='+ props.listGUID +'&Mode=Edit',
                    '_blank' 
                  );
                }                
              },
              customPrev: {
                icon: 'chevron-left',
                click: function() {
                  props.passCurrentDate(calendarRef.current.getApi().getDate().toISOString());
                  calendarPrev();
                }
              },
              customNext: {
                icon:'chevron-right',
                click: function() {
                  props.passCurrentDate(calendarRef.current.getApi().getDate().toISOString());
                  calendarNext();
                }
              }
            }}          
            eventTimeFormat={{
              hour: 'numeric',
              minute: '2-digit',
              meridiem: 'short'
            }}
            initialView='dayGridMonth'   
            eventClassNames={styles.eventItem}           
            editable={false}
            selectable={true}
            selectMirror={true}
            dayMaxEvents={false}
            displayEventEnd={true}
            eventDisplay='block'
            weekends={props.showWeekends}
            eventClick={props.handleDateClick}
            eventSources = {props.eventSources}
            contentHeight='auto'
            eventContent = {(eventInfo)=>{
              if (eventInfo.event._def.extendedProps.roomTitle){
                return(
                  <div className="roomEvent">
                    <div>&nbsp;{eventInfo.event._def.extendedProps.roomTitle} - {eventInfo.event._def.extendedProps.period}</div>
                    <div><i>&nbsp;{eventInfo.event.title}</i></div>
                  </div>
                );
              }else{
                return(
                  <div>
                    <b>{eventInfo.timeText && eventInfo.timeText + ' '}</b>
                    <i>{eventInfo.event.title}</i>
                  </div>
                );
              }
            }}

          />
      </div> 
    );
}
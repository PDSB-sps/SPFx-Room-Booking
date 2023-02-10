# SPFx Room Booking system based on the SPFx-Merged-Calendar solution
A SPFx Merged Calendar React web-part. Aggregates different types of calendars; internal, external, graph, google using Full Calendar plugin.

# Features
- Merged Calendar features
- Adding a calendar of type Room
- Displaying rooms with title, image, color, and interaction options like: booking, view details, and show/hide room
- Displaying room details
- Show/Hide Rooms based on selection
- Booking a room with detecting conflicts and preventing them
- Add/Edit/Delete Booking
- Add to my calendar feature
- Popping notifications on add/edit/delete events using react hot toast library
- UI enhancements

# Dependencies
- Calendar Settings list
- Events list
- Rooms list
- Periods list 
- Guidelines list


# Screenshots
- Room Booking: Calendar, Rooms, Reservation form/panel, Booking management (Room, Period, Guideline) <br/>
![alt Calendar](https://github.com/Maya-Mostafa/SPFx-Room-Booking-/blob/main/RoomBooking.png) <br/>
- Booking Details (Add, Edit, Delete) <br/>
![alt Legend](https://github.com/Maya-Mostafa/SPFx-Room-Booking-/blob/main/RoomBookingDetails.png) <br/>
- Booking Management <br/>
![alt Settings](https://github.com/Maya-Mostafa/SPFx-Room-Booking-/blob/main/AddEditRoom.png) <br/>

# Libraries 
`npm install rrule`<br/>
`npm install --save @fullcalendar/react @fullcalendar/rrule @fullcalendar/daygrid @fullcalendar/timegrid @fullcalendar/interaction`<br/>
`npm install moment`<br/>
`npm install @fluentui/react`<br/>
`npm install @fluentui/react-hooks`<br/>
`npm install office-ui-fabric-core`<br/>
`npm install react-hot-toast`

# Testing
`gulp package-solution`<br/>
`gulp serve --nobrowser`

# Deployment
`gulp bundle --ship`<br/>
`gulp package-solution --ship`

# Room Booking Deployment version 
spfx-room-booking
84fd9f85-a309-4b1e-98fd-db8ae45e1323

# Room Booking Testing version
spfx-room-booking-testing
00f6c2d2-68b5-4e6e-ba23-03566cedad3d

# Room Bl=ooking Testing 2 version XXXX
spfx-room-booking-testing2
f4ad47d0-58ba-43b9-b87c-31723d6f1d03

# Update these files 
C:\myfiles\Github\SPFx-Room-Booking-\.yo-rc.json
C:\myfiles\Github\SPFx-Room-Booking-\package-lock.json
C:\myfiles\Github\SPFx-Room-Booking-\package.json
C:\myfiles\Github\SPFx-Room-Booking-\config\package-solution.json
C:\myfiles\Github\SPFx-Room-Booking-\src\webparts\mergedCalendar\MergedCalendarWebPart.manifest.json


# References
Used PnP-sfx controls for iFrameDialog
https://pnp.github.io/sp-dev-fx-controls-react/controls/IFrameDialog/

`npm i @pnp/sp` <br/>
`npm install @pnp/spfx-controls-react --save --save-exact` <br/>


# To avoid errors in deployment after changing the application ID
1- Delete first these folders:
    - dist
    - lib
    - temp
    - release
2- Run > gulp build
3- Run > gulp serve --nobrowser
4- Run > gulp package-solution 


- To Fix binding issues due to node environment changed
npm rebuild node-sass


#  Multiple Room Booking
-> Large Dialog/Panel Implementation -> panel

## Step 1 - Form (before & while entry) 
1- validation: mandatory fields, endDate > startDate
2- populate school cycle
	a- elementary : 5 & 10 days options
	b- secondary : current school rotart & read-only field
3- populate recurrence cycle day
	a- elementary : 5day ? 1|2|3|4|5 : 1|2|3|4|5|6|7|8|9|10 (if hard-coded)
	b- secondary: get from calendarSettings cycleDays field --> 1,2,3,4
4- load rooms in the room list (or the already loaded state)
5- load periods from the periods list (or the already loaded state)
6- add events booking to my calendar

## Step 2 - Bookings and/or conflicts overview
### Case 1: no conflicts
    1- list bookings overview (grid view)
    2- confirm & cancel btns
### Case 2: conflicts
    1- list bookings/conflicts overview (grid view)
    2- options for override/skip all/individual booking (grid view with checkboxes)
    3- confirm & cancel btns

## Steps 3 - Booking action (POST)
    1- showing preloader with user friendly message during: 
        a- bookings
        b- delete the prev booking if this an overridde(conflict)
        c- add to my calendar bookings
    2- thank you message



# Changes
1- Multi-room booking panel + dialog + features
2- current date window for internal, external & graph
3- external calendars
4- event details popup header styling
5- event slot display fix for regular & room events
6- recurrent events fixes
7- delete to recycle bin instead of perm


# Introducing start and end times instead of periods
1- In the IRoomBook screen, add start & end time fields


2- In the wp properties:
    1- Add  option for selecting periods or free start & end times
    2- Pick a range - NO




















    
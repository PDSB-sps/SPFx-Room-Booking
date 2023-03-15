import * as React from 'react';
import roomStyles from '../Room.module.scss';
import {CommandBarButton, IIconProps, Link, Icon, IContextualMenuProps, CommandButton} from '@fluentui/react';
import {IRoomsManageProps} from './IRoomsManageProps';
import {isUserManage} from '../../Services/WpProperties';

export default function IRoomsManage (props: IRoomsManageProps) {
    
    const addIcon: IIconProps = { iconName: 'Add' };
    const editIcon: IIconProps = { iconName: 'Edit' };

    const newRoomURL = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${props.roomsList}/Newform.aspx`;
    const newPeriodURL = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${props.periodsList}/Newform.aspx`;
    const newGuidelinesURL = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${props.guidelinesList}/Newform.aspx`;

    const editRoomURL = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${props.roomsList}/AllItems.aspx`;
    const editPeriodURL = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${props.periodsList}/AllItems.aspx`;
    const editGuidelinesURL = `${props.context.pageContext.web.serverRelativeUrl}/Lists/${props.guidelinesList}/AllItems.aspx`;

    const addMenuProps: IContextualMenuProps = {
        items: [
          {
            key: 'addRoom',
            text: 'Room',
            onClick : () => props.onRoomsManage(newRoomURL, "Add")
          },
          {
            key: 'addPeriod',
            text: 'Period',
            onClick : () => props.onRoomsManage(newPeriodURL, "Add")
          },
          {
            key: 'addGuideline',
            text: 'Guideline',
            onClick : () => props.onRoomsManage(newGuidelinesURL, "Add")
          },
        ],
        // By default, the menu will be focused when it opens. Uncomment the next line to prevent this.
        // shouldFocusOnMount: false
      };

      const editMenuProps: IContextualMenuProps = {
        items: [
          {
            key: 'addRoom',
            text: 'Rooms',
            onClick : () => props.onRoomsManage(editRoomURL, "All")
          },
          {
            key: 'addPeriod',
            text: 'Periods',
            onClick : () => props.onRoomsManage(editPeriodURL, "All")
          },
          {
            key: 'addGuidelines',
            text: 'Guidelines',
            onClick : () => props.onRoomsManage(editGuidelinesURL, "All")
          },
        ],
        // By default, the menu will be focused when it opens. Uncomment the next line to prevent this.
        // shouldFocusOnMount: false
      };


    return(
        <div className={roomStyles.roomsManage}>
            <hr/>
            <CommandButton iconProps={addIcon} text="Add" menuProps={addMenuProps} />
            <CommandButton iconProps={editIcon} text="Edit" menuProps={editMenuProps} />
            {/* <CommandBarButton iconProps={addIcon} text="Add Room" onClick={() => props.onRoomsManage(newRoomURL)} />
            <CommandBarButton iconProps={addIcon} text="Add Period" onClick={() => props.onRoomsManage(newPeriodURL)}/>
            <CommandBarButton iconProps={addIcon} text="Add Guidelines" onClick={() => props.onRoomsManage(newGuidelinesURL)}/> */}
            {isUserManage(props.context) && props.isPeriods ?
              <CommandBarButton className={roomStyles.siteManage} iconProps={{iconName: 'CalendarYear'}} text="Multiple Booking" onClick={props.openMultiBook} />
              :
              <Link className={roomStyles.siteManage} href={`${props.context.pageContext.web.serverRelativeUrl}/SitePages/Site-Management.aspx`}>
                <Icon iconName='Settings' />Site Management
              </Link>
            }
        </div>
    );
}
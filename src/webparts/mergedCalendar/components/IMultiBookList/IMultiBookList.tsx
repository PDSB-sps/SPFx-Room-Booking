import * as React from 'react';
import roomStyles from '../Room.module.scss';
import styles from '../MergedCalendar.module.scss';
import { IMultiBookListProps } from './IMultiBookListProps';
import {DetailsList, DetailsListLayoutMode, DefaultButton, TooltipHost, Icon, initializeIcons, Toggle, SelectionMode, PrimaryButton, MessageBar, MessageBarType} from '@fluentui/react';
// import { ListView, IViewField, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";

export default function IMultiBookList(props: IMultiBookListProps){
    
    initializeIcons();

    // const onToggleHandler = (itemIndex) => {
    //     return (ev: React.MouseEvent<HTMLElement>, checked?: boolean) => {
    //         props.updateBookings(itemIndex, checked);
    //     };
    // };

    const onToggleHandler = (ev: React.MouseEvent<HTMLElement>, checked?: boolean) => {
        props.updateBookings(checked);
    };

    const columns = [
        {
            key: "column0",
            name: "Conflict",
            fieldName: "conflict",
            isIconOnly: true,
            onRender: (item) => (
                <>
                    {item.conflict &&
                        <TooltipHost content={`Conflict with another booking`}>
                            <Icon style={{color: 'red'}} iconName="AlertSolid" />
                        </TooltipHost>
                    }
                </>
            ),
            minWidth: 16,
			maxWidth: 16,
        },
        {
			key: "column1",
			name: "Booking Date",
			fieldName: "start",
			minWidth: 100,
			isResizable: true,
		},
		{
			key: "column2",
			name: "Title",
			fieldName: "conflictTitle",
			minWidth: 100,
            maxWidth: 400,
			isResizable: true,
		},
        {
			key: "column3",
			name: "Created By",
			fieldName: "conflictAuthor",
			minWidth: 100,
            maxWidth: 400,
			isResizable: true,
		},
        // {
		// 	key: "column4",
		// 	name: "Action",
		// 	minWidth: 180,
		// 	maxWidth: 300,
		// 	isResizable: true,
        //     onRender: (item) => (
        //         <>
        //             {item.conflict &&
        //                 <Toggle
        //                     defaultChecked = {false}
        //                     onText="Overwrite"
        //                     offText="Skip"
        //                     onChange={onToggleHandler(item.index)}
        //               />
        //             }
        //         </>
        //     ),
		// },
	];

/*
    const viewFields:IViewField [] = [
        {
            displayName: "Conflict",
            name: "conflict",
            render: (item) => (
                <>
                    {item.conflict &&
                        <TooltipHost content={`Conflict with another booking`}>
                            <Icon style={{color: 'red'}} iconName="AlertSolid" />
                        </TooltipHost>
                    }
                </>
            ),
            minWidth: 16,
			maxWidth: 16,
        },
        {
			displayName: "Booking Date",
			name: "start",
			minWidth: 100,
			isResizable: true,
		},
		{
			displayName: "Title",
			name: "conflictTitle",
			minWidth: 100,
            maxWidth: 400,
			isResizable: true,
		},
        {
			displayName: "Created By",
			name: "conflictAuthor",
			minWidth: 100,
            maxWidth: 400,
			isResizable: true,
		},
        {
			displayName: "Action",
            name: "action",
			minWidth: 180,
			maxWidth: 300,
			isResizable: true,
            render: (item) => (
                <>
                    {item.conflict &&
                        <Toggle
                            defaultChecked
                            onText="Overwrite"
                            offText="Skip"
                            onChange={onToggleHandler(item.index)}
                      />
                    }
                </>
            ),
		},
	];

    const groupByFields: IGrouping[] = [
        {
            name: "start", 
            order: GroupOrder.ascending 
        },
      ];
*/

    return (
        <>
            <hr/>
            <h3 className={styles.marginT20}>Bookings</h3>
            {props.isConflict &&
                <>
                    <MessageBar messageBarType={MessageBarType.warning}>
                        Please carefully review your bookings and the existing conflicts. Your bookings will automatically be <i>skipped</i> and will not <i>overwrite</i> the existing ones. 
                        If you want to change that, please toggle the booking action from <i>skip</i> to <i>overwrite</i>.
                    </MessageBar>
                    <br/>
                    <Toggle
                        defaultChecked = {false}
                        onText="Overwrite All"
                        offText="Skip All"
                        onChange={onToggleHandler}
                    />
                </>
            }
            <DetailsList
                items={props.bookingList}
                columns={columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
            />
            {/* <ListView
                items={props.bookingList}
                viewFields={viewFields}
                groupByFields={groupByFields}
            /> */}
            <div className={styles.marginT20}>
                <PrimaryButton text="Confirm Bookings" onClick={props.bookEventsHandler} className={styles.marginR10} />
                <DefaultButton text="Cancel" onClick={props.cancelBookingClickHandler}/>
            </div>
                
        </>
    );
}
const moment = library.load('moment-timezone');
const uuid = library.load('uuid');
const pageSize = 100;

//full synchronization code
async function fullSync({ client, dataStore }) {
    let userId = [];
    let URL = `/v1.0/users?$top=${pageSize}`;
    let nextPage = '';
    do {
        const userRequest = await client.fetch(URL);
        if (!userRequest.ok) {
            throw new Error(`Users sync failed ${userRequest.status}:${userRequest.statusText}.`);
        }
        const userResponse = await userRequest.json();
        for (const value of userResponse.value) {
            userId.push(value.id);
            dataStore.save('users', {
                id: value?.id ?? null,
                mail: value?.mail ?? null,
                user_principal_name: value?.userPrincipalName ?? null,
                display_name: value?.displayName ?? null
            });
        }
        nextPage = userResponse['@odata.nextLink'] ?? null;
        if (nextPage != null) URL = `/v1.0/users?${nextPage.split(`?`)[1]}`;
    } while (nextPage);
    const startDate = moment.utc().subtract(1, 'd').format();
    await Promise.all([
        calendarView(client, dataStore, userId, startDate),
        myEvents(client, dataStore, userId)
    ])
}

// calendar View dataloading code
async function calendarView(client, dataStore, userId, startDate) {
    const endDate = moment.utc().add(30, 'd').format();
    for (const id of userId) {
        const calendarViewRequest = await client.fetch(
            `/v1.0/users/${id}/calendarView?startdatetime=${startDate}&enddatetime=${endDate}&$top=${pageSize}`
        );
        if (!calendarViewRequest.ok && calendarViewRequest.status != 404) {
            throw new Error(`Calendar sync failed ${calendarViewRequest.status}:${calendarViewRequest.statusText}.`);
        }
        const calendarViewResponse = await calendarViewRequest.json();
        if (calendarViewRequest.ok) {
            for (const calendarValue of calendarViewResponse.value) {
                dataStore.save('calender_view', {
                    body_content: calendarValue?.body?.content ?? null,
                    body_preview: calendarValue?.bodyPreview ?? null,
                    end_date_time: calendarValue?.end?.dateTime ?? null,
                    i_cal_u_id: calendarValue?.iCalUId ?? null,
                    id: calendarValue?.id ?? null,
                    is_cancelled: calendarValue?.isCancelled ?? null,
                    is_online_meeting: calendarValue?.isOnlineMeeting ?? null,
                    location_display_name: calendarValue?.location?.displayName ?? null,
                    online_meeting_join_url: calendarValue.onlineMeeting?.joinUrl ?? null,
                    online_meeting_provider: calendarValue?.onlineMeetingProvider ?? null,
                    organizer_email_address_a: calendarValue?.organizer?.emailAddress?.address ?? null,
                    organizer_email_address_n: calendarValue?.organizer?.emailAddress?.name ?? null,
                    original_start_time_zone: calendarValue?.originalStartTimeZone ?? null,
                    series_master_id: calendarValue?.seriesMasterId ?? null,
                    start_date_time: calendarValue?.start?.dateTime ?? null,
                    subject: calendarValue?.subject ?? null
                });
                for (const attendees of calendarValue.attendees) {
                    dataStore.save('calender_view_attendees', {
                        unique_id: uuid.v4(),
                        parent_i_cal_u_id: calendarValue?.iCalUId ?? null,
                        root_i_cal_u_id: calendarValue?.iCalUId ?? null,
                        email_address_address: attendees?.emailAddress?.address ?? null,
                        email_address_name: attendees?.emailAddress?.name ?? null,
                        type: attendees?.type ?? null
                    });
                }
            }
        }
    }
}

// myEvents dataloading code
async function myEvents(client, dataStore, userId) {
    for (const id of userId) {
        const myEventRequest = await client.fetch(`/v1.0/users/${id}/calendar/events?$top=${pageSize}`);
        if (!myEventRequest.ok && myEventRequest.status != 404) {
            throw new Error(`Events sync failed ${myEventRequest.status}:${myEventRequest.statusText}`);
        }
        const myEventResponse = await myEventRequest.json();
        if (myEventRequest.ok) {
            for (const myEventsValue of myEventResponse.value) {
                dataStore.save('my_events', {
                    i_cal_u_id: myEventsValue?.iCalUId ?? null,
                    id: myEventsValue?.id ?? null,
                    recurrence_pattern_day_of: myEventsValue?.recurrence?.pattern?.dayOfMonth ?? null,
                    recurrence_pattern_type: myEventsValue?.recurrence?.pattern?.type ?? null,
                    recurrence_range_end_date: myEventsValue?.recurrence?.range?.endDate ?? null
                });
            }
        }
    }
}

//service action for Create One Time Event with Current Timezone
async function createOutlookOneTimeEventCurrentTimezone({
    dataStore,
    client,
    actionParameters,
}) {
    console.log('hello')
    const startDate = moment.utc().subtract(30, 'm').format();
    const responseOfcreateOutlookOneTimeEventCurrentTimezone = await client.fetch(
        `v1.0/me/events`,
        {
            method: "POST",
            body: JSON.stringify({

                subject: actionParameters.subject,
                body: {
                    contentType: "HTML",
                    content: actionParameters.content
                },
                start: {
                    dateTime: actionParameters.startDateTime,
                    timeZone: ""
                },
                end: {
                    dateTime: actionParameters.endDateTime,
                    timeZone: ""
                },
                location: {
                    displayName: actionParameters.location
                },
                attendees: [{
                    emailAddress: {
                        address: actionParameters.email1
                    },
                    type: actionParameters.type1
                },
                {
                    emailAddress: {
                        address: actionParameters.email2
                    },
                    type: actionParameters.type2
                }, {
                    emailAddress: {
                        address: actionParameters.email3
                    },
                    type: actionParameters.type3
                },
                {
                    emailAddress: {
                        address: actionParameters.email4
                    },
                    type: actionParameters.type4
                },
                {
                    emailAddress: {
                        address: actionParameters.email5
                    },
                    type: actionParameters.type5
                },
                {
                    emailAddress: {
                        address: actionParameters.email6
                    },
                    type: actionParameters.type6
                }
                ],
                allowNewTimeProposals: true,
                isOnlineMeeting: actionParameters.isOnlineMeeting,
                onlineMeetingProvider: actionParameters.onlineMeetingProvider

            }),
        }
    );
    console.log(JSON.stringify(responseOfcreateOutlookOneTimeEventCurrentTimezone));
    if (!responseOfcreateOutlookOneTimeEventCurrentTimezone.ok) {
        throw new Error(`createOutlookOneTimeEventCurrentTimezone (${responseOfcreateOutlookOneTimeEventCurrentTimezone.status}: ${responseOfcreateOutlookOneTimeEventCurrentTimezone.statusText})`);
    }
    const user_id = [actionParameters.userId];
    console.log(...user_id);
    await calendarView(client, dataStore, user_id, startDate);
    await myEvents(client, dataStore, user_id)
}

//service action for Create One Time Event with Custom Timezone
async function createOutlookOneTimeEventCustomTimezone({
    dataStore,
    client,
    actionParameters,
}) {
    const startDate = moment.utc().subtract(30, 'm').format();
    const responseOfcreateOutlookOneTimeEventCustomTimezone = await client.fetch(
        `/v1.0/me/events`,
        {
            method: "POST",
            body: JSON.stringify({

                subject: actionParameters.subject,
                body: {
                    contentType: "HTML",
                    content: actionParameters.content
                },
                start: {
                    dateTime: `${actionParameters.startDate}T${actionParameters.startTime}`,
                    timeZone: actionParameters.timezone
                },
                end: {
                    dateTime: `${actionParameters.endDate}T${actionParameters.endTime}`,
                    timeZone: actionParameters.timezone
                },
                location: {
                    displayName: actionParameters.location
                },
                attendees: [
                    {
                        emailAddress: {
                            address: actionParameters.email1
                        },
                        type: actionParameters.type1
                    },
                    {
                        emailAddress: {
                            address: actionParameters.email2
                        },
                        type: actionParameters.type2
                    }, {
                        emailAddress: {
                            address: actionParameters.email3
                        },
                        type: actionParameters.type3
                    }, {
                        emailAddress: {
                            address: actionParameters.email4
                        },
                        type: actionParameters.type4
                    },
                    {
                        emailAddress: {
                            address: actionParameters.email5
                        },
                        type: actionParameters.type5
                    }, {
                        emailAddress: {
                            address: actionParameters.email6
                        },
                        type: actionParameters.type6
                    }
                ],
                allowNewTimeProposals: actionParameters.allowNewTimeProposals,
                isOnlineMeeting: actionParameters.isOnlineMeeting,
                onlineMeetingProvider: actionParameters.onlineMeetingProvider

            }),
        }
    );
    //console.log(JSON.stringify(responseOfcreateOutlookOneTimeEventCustomTimezone));
    if (!responseOfcreateOutlookOneTimeEventCustomTimezone.ok) {
        throw new Error(`createOutlookOneTimeEventCustomTimezone (${responseOfcreateOutlookOneTimeEventCustomTimezone.status}: ${responseOfcreateOutlookOneTimeEventCustomTimezone.statusText})`);
    }
    const user_id = [actionParameters.userId];
    console.log(user_id);
    await calendarView(client, dataStore, user_id, startDate);
    await myEvents(client, dataStore, user_id)
}

//service action for Create Recurring Event with Current Timezone
async function createRecurringEventCurrentTimeZone({
    dataStore,
    client,
    actionParameters,
}) {
    const startDate = moment.utc().subtract(30, 'm').format();
    const responseOfcreateRecurringEventCurrentTimeZone = await client.fetch(
        `/v1.0/me/events`,
        {
            method: "POST",
            body: JSON.stringify({

                subject: actionParameters.subject,
                body: {
                    contentType: "HTML",
                    content: actionParameters.content
                },
                start: {
                    dateTime: actionParameters.startDateTime,
                    timeZone: "UTC"
                },
                end: {
                    dateTime: actionParameters.endDateTime,
                    timeZone: "UTC"
                },
                location: {
                    displayName: actionParameters.location
                },
                attendees: [
                    {
                        emailAddress: {
                            address: actionParameters.email1
                        },
                        type: actionParameters.type1
                    },
                    {
                        emailAddress: {
                            address: actionParameters.email2
                        },
                        type: actionParameters.type2
                    }, {
                        emailAddress: {
                            address: actionParameters.email3
                        },
                        type: actionParameters.type3
                    }, {
                        emailAddress: {
                            address: actionParameters.email4
                        },
                        type: actionParameters.type4
                    },
                    {
                        emailAddress: {
                            address: actionParameters.email5
                        },
                        type: actionParameters.type5
                    }, {
                        emailAddress: {
                            address: actionParameters.email6
                        },
                        type: actionParameters.type6
                    }
                ],
                allowNewTimeProposals: true,
                isOnlineMeeting: actionParameters.isOnlineMeeting,
                onlineMeetingProvider: actionParameters.onlineMeetingProvider,
                recurrence: {
                    pattern: {
                        type: actionParameters.recurrencetype,
                        interval: 1,
                        daysOfWeek: [actionParameters.days],
                        dayOfMonth: actionParameters.dayOfMonth
                    },
                    range: {
                        type: "endDate",
                        startDate: moment(actionParameters.startDate).format('YYYY-MM-DD'),
                        endDate: actionParameters.enddate
                    }
                }
            }),
        }
    );
    //console.log(JSON.stringify(responseOfcreateRecurringEventCurrentTimeZone));
    if (!responseOfcreateRecurringEventCurrentTimeZone.ok) {
        throw new Error(`createRecurringEventCurrentTimeZone (${responseOfcreateRecurringEventCurrentTimeZone.status}: ${responseOfcreateRecurringEventCurrentTimeZone.statusText})`);
    }
    const user_id = [actionParameters.userId];
    console.log(user_id);
    await calendarView(client, dataStore, user_id, startDate);
    await myEvents(client, dataStore, user_id)
}

//service action for Create Recurring Event with Custom Timezone 
async function createRecurringEventCustomTimeZone({
    dataStore,
    client,
    actionParameters,
}) {
    const startDate = moment.utc().subtract(30, 'm').format();
    const responseOfcreateRecurringEventCustomTimeZone = await client.fetch(
        `/v1.0/me/events`,
        {
            method: "POST",
            body: JSON.stringify({

                subject: actionParameters.subject,
                body: {
                    contentType: "HTML",
                    content: actionParameters.content
                },
                start: {
                    dateTime: `${actionParameters.startDate}T${actionParameters.startTime}`,
                    timeZone: actionParameters.timezone
                },
                end: {
                    dateTime: `${actionParameters.endDate}T${actionParameters.endTime}`,
                    timeZone: actionParameters.timezone
                },
                location: {
                    displayName: actionParameters.location
                },
                attendees: [{
                    emailAddress: {
                        address: actionParameters.email1
                    },
                    type: actionParameters.type1
                },
                {
                    emailAddress: {
                        address: actionParameters.email2
                    },
                    type: actionParameters.type2
                },
                {
                    emailAddress: {
                        address: actionParameters.email3
                    },
                    type: actionParameters.type3
                },
                {
                    emailAddress: {
                        address: actionParameters.email4
                    },
                    type: actionParameters.type4
                },
                {
                    emailAddress: {
                        address: actionParameters.email5
                    },
                    type: actionParameters.type5
                },
                {
                    emailAddress: {
                        address: actionParameters.email6
                    },
                    type: actionParameters.type6
                }
                ],
                allowNewTimeProposals: actionParameters.allowNewTimeProposals,
                isOnlineMeeting: actionParameters.isOnlineMeeting,
                onlineMeetingProvider: actionParameters.onlineMeetingProvider,
                recurrence: {
                    pattern: {
                        type: actionParameters.recurrencetype,
                        interval: 1,
                        daysOfWeek: [actionParameters.days],
                        dayOfMonth: actionParameters.dayOfMonth
                    },
                    range: {
                        type: "endDate",
                        startDate: actionParameters.startDate,
                        endDate: actionParameters.recurEndDate
                    }
                }
            })
        })
    //console.log(JSON.stringify(responseOfcreateRecurringEventCustomTimeZone));
    if (!responseOfcreateRecurringEventCustomTimeZone.ok) {
        throw new Error(`createRecurringEventCustomTimeZone (${responseOfcreateRecurringEventCustomTimeZone.status}: ${responseOfcreateRecurringEventCustomTimeZone.statusText})`);
    }
    const user_id = [actionParameters.userId];
    console.log(user_id);
    await calendarView(client, dataStore, user_id, startDate);
    await myEvents(client, dataStore, user_id)
};

//service action for Create Recurring Office Hours with Current Timezone
async function createRecurringOfficeHoursWithCurrentTimezone({
    dataStore,
    client,
    actionParameters,
}) {
    const startDate = moment.utc().subtract(30, 'm').format();
    const responseOfcreateRecurringOfficeHoursWithCurrentTimezone = await client.fetch(
        `/v1.0/me/events`,
        {
            method: "POST",
            body: JSON.stringify({

                subject: actionParameters.subject,
                body: {
                    contentType: "HTML",
                    content: actionParameters.content
                },
                start: {
                    dateTime: actionParameters.startDateTime,
                    timeZone: "UTC"
                },
                end: {
                    dateTime: actionParameters.endDateTime,
                    timeZone: "UTC"
                },
                allowNewTimeProposals: actionParameters.allowNewTimeProposals,
                isOnlineMeeting: actionParameters.isOnlineMeeting,
                onlineMeetingProvider: "teamsForBusiness",
                recurrence: {
                    pattern: {
                        type: actionParameters.recurrencetype,
                        interval: 1,
                        daysOfWeek: [actionParameters.days],
                        dayOfMonth: actionParameters.dayOfMonth
                    },
                    range: {
                        type: "endDate",
                        startDate: moment(actionParameters.startDate).format('YYYY-MM-DD'),
                        endDate: actionParameters.enddate
                    }
                }
            }
            )
        })
    //console.log(JSON.stringify(responseOfcreateRecurringOfficeHoursWithCurrentTimezone));
    if (!responseOfcreateRecurringOfficeHoursWithCurrentTimezone.ok) {
        throw new Error(`createRecurringOfficeHoursWithCurrentTimezone (${responseOfcreateRecurringOfficeHoursWithCurrentTimezone.status}: ${responseOfcreateRecurringOfficeHoursWithCurrentTimezone.statusText})`);
    }
    const user_id = [actionParameters.userId];
    console.log(user_id);
    await calendarView(client, dataStore, user_id, startDate);
    await myEvents(client, dataStore, user_id)
};


integration.define({
    synchronizations: [
        {
            name: 'Outlook',
            fullSyncFunction: fullSync
        }
    ],
    model: {
        tables: [
            {
                name: 'users',
                columns: [
                    {
                        name: 'id',
                        type: 'STRING',
                        length: 255,
                        primaryKey: true
                    },
                    {
                        name: 'mail',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'user_principal_name',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'display_name',
                        type: 'STRING',
                        length: 255
                    }
                ]
            },
            {
                name: 'calender_view',
                columns: [
                    {
                        name: 'i_cal_u_id',
                        type: 'STRING',
                        length: 255,
                        primaryKey: true
                    },
                    {
                        name: 'id',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'body_content',
                        type: 'STRING',
                        length: 2000
                    },
                    {
                        name: 'body_preview',
                        type: 'STRING',
                        length: 2000
                    },
                    {
                        name: 'end_date_time',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'is_cancelled',
                        type: 'BOOLEAN'
                    },
                    {
                        name: 'is_online_meeting',
                        type: 'BOOLEAN'
                    },
                    {
                        name: 'location_display_name',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'online_meeting_join_url',
                        type: 'STRING',
                        length: 500
                    },
                    {
                        name: 'online_meeting_provider',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'organizer_email_address_a',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'organizer_email_address_n',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'original_start_time_zone',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'series_master_id',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'start_date_time',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'subject',
                        type: 'STRING',
                        length: 255
                    }
                ]
            },
            {
                name: 'calender_view_attendees',
                columns: [
                    {
                        name: 'unique_id',
                        type: 'STRING',
                        length: 255,
                        primaryKey: true
                    },
                    {
                        name: 'parent_i_cal_u_id',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'root_i_cal_u_id',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'email_address_address',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'email_address_name',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'type',
                        type: 'STRING',
                        length: 255
                    }
                ]
            },
            {
                name: 'my_events',
                columns: [
                    {
                        name: 'i_cal_u_id',
                        type: 'STRING',
                        length: 255,
                        primaryKey: true
                    },
                    {
                        name: 'id',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'recurrence_pattern_day_of',
                        type: 'INTEGER'
                    },
                    {
                        name: 'recurrence_pattern_type',
                        type: 'STRING',
                        length: 255
                    },
                    {
                        name: 'recurrence_range_end_date',
                        type: 'DATE'
                    }
                ]
            }
        ],
        relationships: [
            {
                name: 'nested_table_1',
                primaryTable: 'calender_view',
                foreignTable: 'calender_view_attendees',
                columnPairs: [
                    {
                        primaryKey: 'i_cal_u_id',
                        foreignKey: 'parent_i_cal_u_id'
                    }
                ]
            },
            {
                name: 'fk_my_events_id',
                primaryTable: 'calender_view',
                foreignTable: 'my_events',
                columnPairs: [
                    {
                        primaryKey: 'series_master_id',
                        foreignKey: 'id'
                    }
                ]
            },
            {
                name: 'fk_calender_v_organizer_e',
                primaryTable: 'users',
                foreignTable: 'calender_view',
                columnPairs: [
                    {
                        primaryKey: 'user_principal_name',
                        foreignKey: 'organizer_email_address_a'
                    }
                ]
            }
        ],
    },
    actions: [
        //createOutlookOneTimeEventCurrentTimezone
        {
            name: "createOutlookOneTimeEventCurrentTimezone",
            parameters: [
                {
                    name: "userId",
                    type: "STRING"
                },
                {
                    name: "content",
                    type: "STRING"
                },
                {
                    name: "email1",
                    type: "STRING"
                },
                {
                    name: "email2",
                    type: "STRING"
                },
                {
                    name: "email3",
                    type: "STRING"
                },
                {
                    name: "email4",
                    type: "STRING"
                },
                {
                    name: "email5",
                    type: "STRING"
                },
                {
                    name: "email6",
                    type: "STRING"
                },
                {
                    name: "type1",
                    type: "STRING"
                },
                {
                    name: "type2",
                    type: "STRING"
                },
                {
                    name: "type3",
                    type: "STRING"
                },
                {
                    name: "type4",
                    type: "STRING"
                },
                {
                    name: "type5",
                    type: "STRING"
                },
                {
                    name: "type6",
                    type: "STRING"
                },
                {
                    name: "startDateTime",
                    type: "STRING"
                },
                {
                    name: "endDateTime",
                    type: "STRING"
                },
                {
                    name: "isOnlineMeeting",
                    type: "BOOLEAN"
                },
                {
                    name: "location",
                    type: "STRING"
                },
                {
                    name: "onlineMeetingProvider",
                    type: "STRING"
                },
                {
                    name: "subject",
                    type: "STRING"
                },
            ],
            function: createOutlookOneTimeEventCurrentTimezone,
        },
        //createOutlookOneTimeEventCustomTimezone
        {
            name: "createOutlookOneTimeEventCustomTimezone",
            parameters: [
                {
                    name: "isOnlineMeeting",
                    type: "BOOLEAN"
                },
                {
                    name: "endDate",
                    type: "STRING"
                },
                {
                    name: "subject",
                    type: "STRING"
                },
                {
                    name: "timezone",
                    type: "STRING"
                },
                {
                    name: "onlineMeetingProvider",
                    type: "STRING"
                },
                {
                    name: "content",
                    type: "STRING"
                },
                {
                    name: "email3",
                    type: "STRING"
                },
                {
                    name: "email2",
                    type: "STRING"
                },
                {
                    name: "email1",
                    type: "STRING"
                },
                {
                    name: "startTime",
                    type: "STRING"
                },
                {
                    name: "email6",
                    type: "STRING"
                },
                {
                    name: "email5",
                    type: "STRING"
                },
                {
                    name: "email4",
                    type: "STRING"
                },
                {
                    name: "type5",
                    type: "STRING"
                },
                {
                    name: "type4",
                    type: "STRING"
                },
                {
                    name: "type3",
                    type: "STRING"
                },
                {
                    name: "type2",
                    type: "STRING"
                },
                {
                    name: "type6",
                    type: "STRING"
                },
                {
                    name: "type1",
                    type: "STRING"
                },
                {
                    name: "userId",
                    type: "STRING"
                },
                {
                    name: "location",
                    type: "STRING"
                },
                {
                    name: "endTime",
                    type: "STRING"
                },
                {
                    name: "startDate",
                    type: "STRING"
                },

            ],
            function: createOutlookOneTimeEventCustomTimezone,
        },
        //createRecurringEventCurrentTimeZone
        {
            name: "createRecurringEventCurrentTimeZone",
            parameters: [
                {
                    name: "isOnlineMeeting",
                    type: "BOOLEAN"
                },
                {
                    name: "subject",
                    type: "STRING"
                },
                {
                    name: "onlineMeetingProvider",
                    type: "STRING"
                },
                {
                    name: "content",
                    type: "STRING"
                },
                {
                    name: "email6",
                    type: "STRING"
                },
                {
                    name: "email5",
                    type: "STRING"
                },
                {
                    name: "email4",
                    type: "STRING"
                },
                {
                    name: "email3",
                    type: "STRING"
                },
                {
                    name: "emai2",
                    type: "STRING"
                },
                {
                    name: "email1",
                    type: "STRING"
                },
                {
                    name: "type1",
                    type: "STRING"
                },
                {
                    name: "type2",
                    type: "STRING"
                },
                {
                    name: "type3",
                    type: "STRING"
                },
                {
                    name: "type4",
                    type: "STRING"
                },
                {
                    name: "type5",
                    type: "STRING"
                },
                {
                    name: "type6",
                    type: "STRING"
                },
                {
                    name: "endDateTime",
                    type: "STRING"
                },
                {
                    name: "userId",
                    type: "STRING"
                },
                {
                    name: "startDateTime",
                    type: "STRING"
                },
                {
                    name: "enddate",
                    type: "STRING"
                },
                {
                    name: "dayOfMonth",
                    type: "STRING"
                },
                {
                    name: "days",
                    type: "STRING"
                },
                {
                    name: "location",
                    type: "STRING"
                },
                {
                    name: "startDate",
                    type: "STRING"
                },
                {
                    name: "recurrencetype",
                    type: "STRING"
                },
            ],
            function: createRecurringEventCurrentTimeZone,
        },
        //createRecurringEventCustomTimeZone
        {
            name: "createRecurringEventCustomTimeZone",
            parameters: [
                {
                    name: "isOnlineMeeting",
                    type: "BOOLEAN"
                },
                {
                    name: "endDate",
                    type: "STRING"
                },
                {
                    name: "subject",
                    type: "STRING"
                },
                {
                    name: "timezone",
                    type: "STRING"
                },
                {
                    name: "onlineMeetingProvider",
                    type: "STRING"
                },
                {
                    name: "content",
                    type: "STRING"
                },
                {
                    name: "email6",
                    type: "STRING"
                },
                {
                    name: "email5",
                    type: "STRING"
                },
                {
                    name: "email4",
                    type: "STRING"
                },
                {
                    name: "email3",
                    type: "STRING"
                },
                {
                    name: "emai2",
                    type: "STRING"
                },
                {
                    name: "email1",
                    type: "STRING"
                },
                {
                    name: "type1",
                    type: "STRING"
                },
                {
                    name: "type2",
                    type: "STRING"
                },
                {
                    name: "type3",
                    type: "STRING"
                },
                {
                    name: "type4",
                    type: "STRING"
                },
                {
                    name: "type5",
                    type: "STRING"
                },
                {
                    name: "type6",
                    type: "STRING"
                },
                {
                    name: "recurEndDate",
                    type: "STRING"
                },
                {
                    name: "startTime",
                    type: "STRING"
                },
                {
                    name: "userId",
                    type: "STRING"
                },
                {
                    name: "dayOfMonth",
                    type: "STRING"
                },
                {
                    name: "days",
                    type: "STRING"
                },
                {
                    name: "location",
                    type: "STRING"
                },
                {
                    name: "endTime",
                    type: "STRING"
                },
                {
                    name: "startDate",
                    type: "STRING"
                },
                {
                    name: "recurrencetype",
                    type: "STRING"
                },
            ],
            function: createRecurringEventCustomTimeZone,
        },
        //createRecurringOfficeHoursWithCurrentTimezone
        {
            name: "createRecurringOfficeHoursWithCurrentTimezone",
            parameters: [
                {
                    name: "isOnlineMeeting",
                    type: "BOOLEAN"
                },
                {
                    name: "subject",
                    type: "STRING"
                },
                {
                    name: "endDateTime",
                    type: "STRING"
                },
                {
                    name: "userId",
                    type: "STRING"
                },
                {
                    name: "content",
                    type: "STRING"
                },
                {
                    name: "startDateTime",
                    type: "STRING"
                },
                {
                    name: "enddate",
                    type: "STRING"
                },
                {
                    name: "dayOfMonth",
                    type: "STRING"
                },
                {
                    name: "days",
                    type: "STRING"
                },
                {
                    name: "startDate",
                    type: "STRING"
                },
                {
                    name: "recurrencetype",
                    type: "STRING"
                },
            ],
            function: createRecurringOfficeHoursWithCurrentTimezone,
        },

    ]
});

async function fullSync({ client, dataStore }) {
	let userId = [];
	let URL = `/v1.0/users?$top=${pageSize}`;
	let nextPage = '';
	do {
		const userRequest = await client.fetch(URL);
		if (!userRequest.ok) {
			throw new Error(`Users sync failed ${userRequest.status}:${userRequest.statusText}.`);
		}
		const userResponse = await userRequest.json();
		for (const value of userResponse.value) {
			userId.push(value.id);
			dataStore.save('user', {
				id: value?.id ?? null,
				mail: value?.mail ?? null,
				user_principal_name: value?.userPrincipalName ?? null,
				display_name: value?.displayName ?? null
			});
		}
		nextPage = userResponse['@odata.nextLink'] ?? null;
		if (nextPage != null) URL = `/v1.0/users?${nextPage.split(`?`)[1]}`;
	} while (nextPage);
	await Promise.all([
		calendarView(client, dataStore, userId),
		myEvents(client, dataStore, userId)
	])
}

async function calendarView(client, dataStore, userId) {
const startDate = moment.utc().subtract(1, 'd').format();
const endDate = moment.utc().add(30, 'd').format();
	for (const id of userId) {
		const calendarViewRequest = await client.fetch(
			`/v1.0/users/${id}/calendarView?startdatetime=${startDate}&enddatetime=${endDate}&$top=${pageSize}`
		);
		if (!calendarViewRequest.ok && calendarViewRequest.status != 404) {
			throw new Error(`Calendar sync failed ${calendarViewRequest.status}:${calendarViewRequest.statusText}.`);
		}
		const calendarViewResponse = await calendarViewRequest.json();
		if (calendarViewRequest.ok) {
			for (const calendarValue of calendarViewResponse.value) {
				dataStore.save('calendar_view', {
					body_content: calendarValue?.body?.content ?? null,
					body_preview: calendarValue?.bodyPreview ?? null,
					end_date_time: calendarValue?.end?.dateTime ?? null,
					i_cal_u_id: calendarValue?.iCalUId ?? null,
					id: calendarValue?.id ?? null,
					is_cancelled: calendarValue?.isCancelled ?? null,
					is_online_meeting: calendarValue?.isOnlineMeeting ?? null,
					location_display_name: calendarValue?.location?.displayName ?? null,
					online_meeting_join_url: calendarValue.onlineMeeting?.joinUrl ?? null,
					online_meeting_provider: calendarValue?.onlineMeetingProvider ?? null,
					organizer_email_address_a: calendarValue?.organizer?.emailAddress?.address ?? null,
					organizer_email_address_n: calendarValue?.organizer?.emailAddress?.name ?? null,
					original_start_time_zone: calendarValue?.originalStartTimeZone ?? null,
					series_master_id: calendarValue?.seriesMasterId ?? null,
					start_date_time: calendarValue?.start?.dateTime ?? null,
					subject: calendarValue?.subject ?? null
				});
				for (const attendees of calendarValue.attendees) {
					dataStore.save('calender_view_attendees', {
						unique_id: uuid.v4(),
						parent_i_cal_u_id: calendarValue?.iCalUId ?? null,
						root_i_cal_u_id: calendarValue?.iCalUId ?? null,
						email_address_address: attendees?.emailAddress?.address ?? null,
						email_address_name: attendees?.emailAddress?.name ?? null,
						type: attendees?.type ?? null
					});
				}
			}
		}
	}
}

async function myEvents(client, dataStore, userId) {
	for (const id of userId) {
		const myEventRequest = await client.fetch(`/v1.0/users/${id}/calendar/events?$top=${pageSize}`);
		if (!myEventRequest.ok && myEventRequest.status != 404) {
			throw new Error(`Events sync failed ${myEventRequest.status}:${myEventRequest.statusText}`);
		}
		const myEventResponse = await myEventRequest.json();
		if (myEventRequest.ok) {
			for (const myEventsValue of myEventResponse.value) {
				dataStore.save('my_events', {
					i_cal_u_id: myEventsValue?.iCalUId ?? null,
					id: myEventsValue?.id ?? null,
					recurrence_pattern_day_of: myEventsValue?.recurrence?.pattern?.dayOfMonth ?? null,
					recurrence_pattern_type: myEventsValue?.recurrence?.pattern?.type ?? null,
					recurrence_range_end_date: myEventsValue?.recurrence?.range?.endDate ?? null
				});
			}
		}
	}
}

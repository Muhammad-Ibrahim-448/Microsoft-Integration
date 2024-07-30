frappe.ui.form.on(cur_frm.doc.doctype, {
    refresh: function(frm) {
        frm.add_custom_button(__('Create Meeting'), function() {
            frappe.call({
                method: "custom_app.custom_app.doctype.microsoft_integration.microsoft_integration.get_active_users",
                callback: function(active_users_response) {
                    if (active_users_response.message) {
                        var active_users = active_users_response.message;
                        frappe.prompt([
                            {'fieldname': 'subject', 'fieldtype': 'Data', 'label': 'Subject', 'reqd': true},
                            {'fieldname': 'body', 'fieldtype': 'Text', 'label': 'Body', 'reqd': false},
                            {'fieldname': 'start', 'fieldtype': 'Datetime', 'label': 'Start Date and Time', 'reqd': true},
                            {'fieldname': 'end', 'fieldtype': 'Datetime', 'label': 'End Date and Time', 'reqd': true},
                            {'fieldname': 'location', 'fieldtype': 'Data', 'label': 'Location', 'reqd': true},
                            {
                                'fieldname': 'attendees', 'fieldtype': 'MultiSelect', 'label': 'Attendees',
                                'options': active_users.map(user => ({
                                    label: `${user.full_name} <${user.email}>`,
                                    value: user.email // Store only the email address
                                }))
                            }
                        ], function(values) {
                            console.log('Raw attendees data:', values.attendees);
                            // Ensure attendees is always treated as an array
                            let attendees = Array.isArray(values.attendees) ? values.attendees : [values.attendees];
                            console.log('Attendees data before sending:', attendees);
                            show_meeting_popup({
                                subject: values.subject,
                                body: values.body,
                                start: values.start,
                                end: values.end,
                                location: values.location,
                                attendees: attendees // Send only email addresses
                            });
                        }, __('Enter Meeting Details'), 'Schedule');
                    } else {
                        frappe.msgprint('Failed to fetch active users.');
                    }
                }
            });
        });
    }
});



function show_meeting_popup(values) {
    try {
        let attendeesJson = JSON.stringify(values.attendees); // Convert to JSON string

        frappe.call({
            method: "custom_app.custom_app.doctype.microsoft_integration.microsoft_integration.schedule_meeting",
            args: {
                subject: values.subject,
                body: values.body,
                start: frappe.datetime.str_to_user(values.start),
                end: frappe.datetime.str_to_user(values.end),
                location: values.location,
                attendees: attendeesJson // Send as JSON string
            },
            callback: function(response) {
                console.log("Meeting scheduled response: ", response);
                if (response.message.status === "success") {
                    frappe.msgprint(__('Meeting has been scheduled successfully.'));
                } else {
                    frappe.msgprint(__('Failed to schedule meeting: ' + response.message.message));
                }
            },
            error: function(error) {
                console.error("Error scheduling meeting: ", error);
                frappe.msgprint(__('An error occurred while scheduling the meeting.'));
            }
        });
    } catch (e) {
        console.error("Error in show_meeting_popup: ", e);
        frappe.msgprint(__('An error occurred while preparing the request.'));
    }
}

frappe.ui.form.on('Sales Order', {
    refresh: function(frm) {
        frm.add_custom_button(__('Create Teams Meeting'), function() {
            frappe.call({
                method: "custom_app.custom_app.doctype.microsoft_integration.microsoft_integration.create_teams_meeting",
                args: {
                    docname: frm.docname
                },
                callback: function(r) {
                    if (r.message) {
                        frappe.msgprint(__('Teams meeting created successfully!'));
                    } else {
                        frappe.msgprint(__('Failed to create Teams meeting.'));
                    }
                }
            });
        });
    }
});

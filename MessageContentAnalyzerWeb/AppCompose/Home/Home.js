/// <reference path="../App.js" />
var item;
var table;

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            table = document.getElementById("details");
            item = Office.context.mailbox.item;
            $('#set-subject').click(setSubject);
            $('#get-subject').click(getSubject);
            $('#add-to-recipients').click(addToRecipients);
            $('#add-attachment').click(addAttachments);
        });
    };

    function setSubject() {
        Office.cast.item.toItemCompose(item).subject.setAsync("Hello world!");
    }

    function getSubject() {
        Office.cast.item.toItemCompose(item).subject.getAsync(function (result) {
            app.showNotification('The current subject is', result.value)
        });
    }

    function addToRecipients() {
        var addressToAdd = {
            displayName: Office.context.mailbox.userProfile.displayName,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
        };

        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
        }
    }

    function addAttachments() {
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).addFileAttachmentAsync("https://i.imgur.com/ucI9vyz.png", "image file", { asyncContext: null },
                function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        app.showNotification(asyncResult.error.message);
                    }
                    else {
                        app.showNotification('ID of added attachment: ' + asyncResult.value);
                    }
                });
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).addFileAttachmentAsync("https://i.imgur.com/ucI9vyz.png", "image file");
        }
    }

})();
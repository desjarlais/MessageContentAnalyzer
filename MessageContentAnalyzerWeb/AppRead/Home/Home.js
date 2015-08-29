/// <reference path="../App.js" />

var item;
var table;
var diag;
var userprofile;
var settings;
var lastAccess;
var customProp;
var customPropError;
var roamingPropError;

(function () {
    "use strict";
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            table = document.getElementById("details");
            item = Office.context.mailbox.item;
            diag = Office.context.mailbox.diagnostics;
            userprofile = Office.context.mailbox.userProfile;
            settings = Office.context.roamingSettings;

            // load custom props for the current item
            item.loadCustomPropertiesAsync(customPropCallback);
            $("#footer").hide();
            
            // build up html and populate data
            buildHtmlTable(item.itemType);
            displayMailboxInfo();
            displayMessageDetails(item.itemType);

            // get the previous roamed setting
            lastAccess = settings.get("LastAccess");

            // initialize button clicks
            $('#sendRequest').click(sendRequest);
            $('#getCustomProps').click(getCustomProps);
            $('.header').click(function () {
                $(this).nextUntil('tr.header').slideToggle(10);
            })
        });
    };
    
    // app level setting for the mailbox
    function setRoamingSetting() {
        settings.set("LastAccess", Date());
        settings.saveAsync(roamingCallback);
    }

    function roamingCallback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            roamingPropError = asyncResult.error.message;
        }
    }
    
    // item level custom prop
    function customPropCallback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            customPropError = asyncResult.error.message;
        }
        else {
            customProp = asyncResult.value;
            customProp.set("myProp", "myValue");
            customProp.saveAsync(saveCallback);
        }
    }

    function saveCallback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            customPropError = asyncResult.error.message;
        }
    }

    // build html table
    function buildHtmlTable(type) {
        // start with common info that doesn't depend on the type of item (message/appointment)
        var mailboxInfo = '<tr class="header"><td colspan="2" span style="cursor:default">User Info [+/-]</td></tr>' +
                '<tr id="rowdisplayname"><th>Display Name:</th><td id="displayname"></td></tr>' +
                '<tr id="rowemailaddress"><th>Email Address:</th><td id="emailaddress"></td></tr>' +
                '<tr id="rowtimezone"><th>TimeZone:</th><td id="timezone"></td></tr>' +
                '<tr class="header"><td colspan="2" span style="cursor:default">Diagnostics [+/-]</td></tr>' +
                '<tr id="rowhostname"><th>Host Name:</th><td id="hostname"></td></tr>' +
                '<tr id="rowhostversion"><th>Host Version:</th><td id="hostversion"></td></tr>' +
                '<tr id="rowowaview"><th>OWA View:</th><td id="owaview"></td></tr>' +
                '<tr id="rowlastaccess"><th>Last Accessed:</th><td id="lastaccess"></td></tr>' +
                '<tr class="header"><td colspan="2" span style="cursor:default">Item Info [+/-]</td></tr>' +
                '<tr id="rowdatetimecreated"><th>DateTimeCreated:</th><td id="dtcreate"></td></tr>' +
                '<tr id="rowdatetimemodified"><th>DateTimeModified:</th><td id="dtmodify"></td></tr>' +
                '<tr id="rowitemclass"><th>Item Class:</th><td id="itemcclass"></td></tr>' +
                '<tr id="rowitemtype"><th>Item Type:</th><td id="itemtype"></td></tr>' +
                '<tr id="rowewsid"><th>EWS Item ID:</th><td id="ewsid"></td></tr>' +
                '<tr id="rowcustomprop"><th>Custom Property 1:</th><td id="customprop"></td></tr>' +
                '<tr class="header"><td colspan="2" span style="cursor:default">Entities [+/-]</td></tr>' +
                '<tr id="rowaddresses"><th>Addresses:</th><td id="addresses"></td></tr>' +
                '<tr id="rowcontacts"><th>Contacts:</th><td id="contacts"></td></tr>' +
                '<tr id="rowphone"><th>Phone Numbers:</th><td id="phone"></td></tr>' +
                '<tr id="rowurl"><th>URLs:</th><td id="url"></td></tr>' +
                '<tr id="rowemails"><th>Emails:</th><td id="emails"></td></tr>' +
                '<tr id="rowmeetingsuggestions"><th>Meeting Suggestions:</th><td id="meetings"></td></tr>' +
                '<tr id="rowtasksuggestions"><th>Task Suggestions:</th><td id="tasks"></td></tr>';

        if (type == Office.MailboxEnums.ItemType.Message) {
            // now populate the message specific data
            var msg = '<tr class="header"><td colspan="2" span style="cursor:default">Message Info [+/-]</td></tr>' +
                '<tr id="rowattachmentid"><th>Attachment ID:</th><td id="attachmentid"></td></tr>' +
                '<tr id="rowcc"><th>Cc:</th><td id="cc"></td></tr>' +
                '<tr id="rowconversationid"><th>Conversation ID:</th><td id="conversationid"></td></tr>' +
                '<tr id="rowfrom"><th>From:</th><td id="from"></td></tr>' +
                '<tr id="rowmsgid"><th>Internet Message ID:</th><td id="messageid"></td></tr>' +
                '<tr id="rownormalizedsubject"><th>Normalized Subject:</th><td id="normalizedsubject"></td></tr>' +
                '<tr id="rowsender"><th>Sender:</th><td id="sender"</td></tr>' +
                '<tr id="rowsubject"><th>Subject:</th><td id="subject"></td></tr>' +
                '<tr id="rowto"><th>To:</th><td id="to"></td></tr>' +
                '<tr id="rowresources"><th>Resources:</th><td id="resources"></td></tr>'


            var el = document.getElementById('details');
            el.innerHTML = mailboxInfo + msg;
        }
        else {
            // or populate the appointment specific data
            var appt = '<tr class="header"><td colspan="2" span style="cursor:default">Appointment Info [+/-]</td></tr>' +
                '<tr id="rowsubject"><th>Subject:</th><td id="subject"></td></tr>' +
                '<tr id="rownormalizedsubject"><th>Normalized Subject:</th><td id="normalizedsubject"></td></tr>' +
                '<tr id="rowstart"><th>Start:</th><td id="start"></td></tr>' +
                '<tr id="rowend"><th>End:</th><td id="end"></td></tr>' +
                '<tr id="rowlocation"><th>Location:</th><td id="location"></td></tr>' +
                '<tr id="rowrequiredattendees"><th>Required Attendees:</th><td id="requiredattendees"></td></tr>' +
                '<tr id="rowoptionalattendees"><th>Optional Attendees:</th><td id="optionalattendees"></td></tr>' +
                '<tr id="rowresources"><th>Resources:</th><td id="resources"></td></tr>' +
                '<tr id="rowattachmentid"><th>Attachment ID:</th><td id="attachmentid"></td></tr>' +
                '<tr id="roworganizer"><th>Organizer:</th><td id="organizer"></td></tr>';

            var el = document.getElementById('details');
            el.innerHTML = mailboxInfo + appt;
        }
    }

    // display mailbox info
    function displayMailboxInfo() {
        $('#displayname').text(userprofile.displayName);
        $('#emailaddress').text(userprofile.emailAddress);
        $('#timezone').text(userprofile.timeZone);
        $('#hostname').text(diag.hostName);
        $('#hostversion').text(diag.hostVersion);
        $('#hostowaview').text(diag.OWAView);
        $('#ewsid').text(item.itemId);
        $('#dtcreate').text(item.dateTimeCreated);
        $('#dtmodify').text(item.dateTimeModified);
        $('#itemclass').text(item.itemClass);
        $('#itemtype').text(item.itemType);
        $('#lastaccess').text(settings.get("LastAccess"));
    }

    // display the fields and entities for the current item
    function displayMessageDetails(type) {
        // populate common fields
        $('#subject').text(item.subject);
        $('#normalizedsubject').text(item.normalizedSubject);
        addFieldToTable(item.attachments, '#attachmentid', "rowattachmentid");

        // populate unique fields
        if (type == Office.MailboxEnums.ItemType.Message) {
            $('#conversationid').text(item.conversationId);
            $('#messageid').text(item.internetMessageId);
            $('#sender').text(item.sender.emailAddress);
            addFieldToTable(item.cc, '#cc', "rowcc");
            addFieldToTable(item.from, '#from', "rowfrom");
            addFieldToTable(item.to, '#to', "rowto");
        }
        else {
            $('#start').text(item.start);
            $('#end').text(item.end);
            $('#location').text(item.location);
            addFieldToTable(item.organizer, '#organizer', "roworganizer");
            addFieldToTable(item.resources, '#resources', "rowresources");
            addFieldToTable(item.requiredAttendees, '#requiredattendees', "rowrequiredattendees");
            addFieldToTable(item.optionalAttendees, '#optionalattendees', "rowoptionalattendees");
        }

        // populate entities
        addEntitiesToTable(item.getEntities().addresses, '#addresses', "rowaddresses", "Address #");
        addEntitiesToTable(item.getEntities().contacts, '#contacts', "rowcontacts", "Contacts #");
        addEntitiesToTable(item.getEntities().phoneNumbers, '#phone', "rowphone", "Phone #");
        addEntitiesToTable(item.getEntities().urls, '#url', "rowurl", "URL #");
        addEntitiesToTable(item.getEntities().emailAddresses, '#emails', "rowemails", "Email #");
        addEntitiesToTable(item.getEntities().taskSuggestions, '#tasks', "rowtasksuggestions", "Task #");
        addEntitiesToTable(item.getEntities().meetingSuggestions, '#meetings', "rowmeetingsuggestions", "Meeting Suggestion #");
    }

    // check field length and populate table
    function addFieldToTable(field, cellTag, rowTag) {
        if (field.length === 0) {
            $(cellTag).text("Field values unavailable.");
        }
        else if (field.length == undefined)
        {
            if (rowTag === "roworganizer") {
                if (_isOrganizer()) {
                    $(cellTag).text("you are the organizer.");
                }
                else {
                    $(cellTag).text(item.organizer.emailAddress + " is the organizer.");
                }
            }
        }
        else {
            var fieldStartRow = document.getElementById(rowTag).rowIndex + 1;
            var attendees = item.requiredAttendees;
            
            for (var i = 0; i < field.length; i++) {
                var row = table.insertRow(fieldStartRow);
                var cell1 = row.insertCell(0);
                var cell2 = row.insertCell(1);
                if (rowTag === "rowattachmentid") {
                    cell1.innerHTML = field[i].name;
                    cell2.innerHTML = field[i].id;
                }
                else if (rowTag === "rowrequiredattendees") {
                    if (_isOrganizer()) {
                        cell2.innerHTML = field[i].emailAddress + " || Attendee Response = " + attendees[i].appointmentResponse;
                    }
                    else {
                        cell2.innerHTML = field[i].emailAddress;
                    }
                }
                else if (rowTag === "rowoptionalattendees") {
                    if (_isOrganizer()) {
                        cell2.innerHTML = field[i].emailAddress + " || Attendee Response = " + attendees[i].appointmentResponse;
                    }
                    else {
                        cell2.innerHTML = field[i].emailAddress;
                    }
                }
                else if (rowTag === "rowcc" || rowTag === "rowbcc" || rowTag === "rowfrom" || rowTag === "rowto") {
                    cell2.innerHTML = field[i].emailAddress;
                }
                else {
                    cell2.innerHTML = field[i];
                }
            }
        }
    }

    // check entity length and populate table
    function addEntitiesToTable(entity, cellTag, rowTag, colText) {
        if (entity.length === 0) {
            $(cellTag).text("Entity values unavailable.");
        }
        else {
            var entityStartRow = document.getElementById(rowTag).rowIndex + 1;

            for (var i = 0; i < entity.length; i++) {
                var row = table.insertRow(entityStartRow);
                var cell1 = row.insertCell(0);
                var cell2 = row.insertCell(1);
                cell1.innerHTML = colText + (i + 1) + ":";
                switch (colText) {
                    case "Contacts #":
                        cell2.innerHTML = entity[i].personName;
                        break;
                    case "Phone #":
                        cell2.innerHTML = entity[i].phoneString;
                        break;
                    case "Attendee #":
                        cell2.innerHTML = entity[i].emailAddress;
                        break;
                    case "Task #":
                        cell2.innerHTML = entity[i].taskString;
                        break;
                    case "Meeting Suggestion #":
                        cell2.innerHTML = entity[i].meetingString;
                        break;
                    default:
                        cell2.innerHTML = entity[i];
                        break;
                }
            }
        }
    }

    function getSoapEnvelope(request) {
        // Wrap an Exchange Web Services request in a SOAP envelope. 
        var result =
        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013_SP1" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' + request +
        ' </soap:Body>' +
        '</soap:Envelope>';

        return result;
    };

    function getSubjectRequest(id) {
        // Return a GetItem EWS operation request for the subject of the specified item.  
        var result =
        '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '        <t:AdditionalProperties>' +
        '            <t:FieldURI FieldURI="item:Subject"/>' +
        '        </t:AdditionalProperties>' +
        '      </ItemShape>' +
        '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
        '    </GetItem>';

        return result;
    };

    // get custom properties
    function getCustomProps() {
        var customVal = document.getElementById("itemCustomProps");
        customVal.innerText = customProp.get("myProp");
    }

    // Send an EWS request for the message's subject. 
    function sendRequest() {
        // Create a local variable that contains the mailbox. 
        var mailbox = Office.context.mailbox;
        var request = getSubjectRequest(mailbox.item.itemId);
        var envelope = getSoapEnvelope(request);

        mailbox.makeEwsRequestAsync(envelope, callback);
    };

    // Function called when the EWS request is complete. 
    function callback(asyncResult) {
        var response = asyncResult.value;
        var context = asyncResult.context;

        // Process the returned response here. 
        var responseSpan = document.getElementById("response");
        responseSpan.innerText = response;
    };

    // check if an item is an appointment or meeting request
    var _isCalendarItem = function()
    {
        if ((item.itemType == Office.MailboxEnums.ItemType.Appointment) ||
            (item.itemClass.indexOf("IPM.Schedule") != -1))
        {
            return true;
        }

        return false;
    }

    // check if the current user is the organizer of a meeting
    var _isOrganizer = function()
    {
        if ((item.itemType == Office.MailboxEnums.ItemType.Appointment) &&
            (userprofile.emailAddress == item.organizer.emailAddress))
        {
            return true;
        }

        return false;
    }
})();
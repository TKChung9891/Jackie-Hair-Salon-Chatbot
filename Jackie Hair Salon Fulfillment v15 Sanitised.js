// TK: ver15 Jackie Hair Salon English Version Sanitised
// Intents served by Fulfillment Code:
//      Intent 'MakeAppt' => makeAppointment        }
//      Intent 'ChangeStylist'=> makeAppointment    }  
//      Intent 'ChangeDateTime' => makeAppointment  }
//      Intent 'ChangeTime' => makeAppointment      } All four map to same function
//      Intent 'OrderShampoo' => ShampooHandler     ] 
//      Intent 'AddShampoo' => ShampooHandler       ] Both map to same function
//      Intent 'DeleteAppt' => DeleteApptHandler
//      Intent 'GiveFeedBack' => GiveFeedBackHandler

'use strict';

const functions = require('firebase-functions');
const { google } = require('googleapis');
const { WebhookClient } = require('dialogflow-fulfillment');

// Firstore DB Configuration
// Require "firebase-admin": "^8.2.0" dependency at package json
const admin = require('firebase-admin');
admin.initializeApp();
const db = admin.firestore();

// Email & Nodemailer Configuration
// Require "nodemailer": "^4.6.7" dependency at package json
// ChatbotEmail needs less secure setting at Google Account
const ChatbotEmail = 'xxx@gmail.com'; //For chatbot send out confirmation emails
const FulfillmentEmail = 'xxx@gmail.com'; //For Shampoo Fulfillment Team receive emails
const nodemailer = require('nodemailer');
const mailTransport = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: '465',
    secure: 'true',
    service: 'Gmail',
    auth: {
        user: ChatbotEmail,
        pass: 'xxx'     // Insert password to Chatbot Email here
    }
});

// Google Calendar Configuration
// Calendar ID from shared Google Calendar  
const CalendarIdJackie = "xxx@group.calendar.google.com";
const CalendarIdSamantha = "xxx@group.calendar.google.com";
const CalendarIdBrian = "xxx@group.calendar.google.com";

// JSON File downloaded from Google Calendar Service Acct create credentials
// Starting with "type": "service_account"... 
const serviceAccount = {
    "type": "service_account", "...": "..."
};
const serviceAccountAuth = new google.auth.JWT({
    email: serviceAccount.client_email,
    key: serviceAccount.private_key,
    scopes: 'https://www.googleapis.com/auth/calendar'
});
const calendar = google.calendar('v3');
const timeZoneOffset = '+08:00';

process.env.DEBUG = 'dialogflow:*'; // enables lib debugging statements

exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
    const agent = new WebhookClient({ request, response });

    function makeAppointment(agent) {
        // Calculate appointment start and end datetimes (end = +1hr from start)
        const AppointmentDate = agent.parameters.Date.split('T')[0];
        const AppointmentTime = agent.parameters.Time.split('T')[1].substr(0, 8);
        const dateTimeStart = new Date(AppointmentDate + 'T' + AppointmentTime + timeZoneOffset);
        const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
        const CustName = agent.parameters.CustName;
        const appointment_type = CustName + ' ' + agent.parameters.CustMobile + ' ' + agent.parameters.Service;
        var CalendarId;
        var Stylist;
        Stylist = agent.parameters.Stylist;
        if (Stylist == "Jackie") { CalendarId = CalendarIdJackie; }
        if (Stylist == "Samantha") { CalendarId = CalendarIdSamantha; }
        if (Stylist == "Brian") { CalendarId = CalendarIdBrian; }

        // Check the availibility of the time, and make an appointment if there is time on the calendar
        return createCalendarEvent(dateTimeStart, dateTimeEnd, CustName, appointment_type, CalendarId).then(() => {
            agent.add(`Let me see if we can fit you in for ${Stylist} on ${AppointmentDate} at ${AppointmentTime}! Yes It is fine!.`);
            agent.setFollowupEvent('SuccessAppt');
        }).catch(() => {
            agent.add(`I'm sorry, there are no slots available for ${Stylist} on ${AppointmentDate} at ${AppointmentTime}.`);
            agent.setFollowupEvent('FailureAppt');
        });
    }

    function createCalendarEvent(dateTimeStart, dateTimeEnd, CustName, appointment_type, CalendarId) {
        return new Promise((resolve, reject) => {
            calendar.events.list({
                auth: serviceAccountAuth, // List events for time period
                calendarId: CalendarId,
                timeMin: dateTimeStart.toISOString(),
                timeMax: dateTimeEnd.toISOString()
            }, (err, calendarResponse) => {
                // Check if there is a event already on the Calendar
                if (err || calendarResponse.data.items.length > 0) {
                    reject(err || new Error('Requested time conflicts with another appointment'));
                } else {
                    // Create event for the requested time period
                    calendar.events.insert({
                        auth: serviceAccountAuth,
                        calendarId: CalendarId,
                        resource: {
                            summary: CustName,
                            description: appointment_type,
                            start: { dateTime: dateTimeStart },
                            end: { dateTime: dateTimeEnd }
                        }
                    }, (err, event) => {
                        err ? reject(err) : resolve(event);
                    }
                    );
                }
            });
        });
    }

    function DeleteApptHandler(agent) {
        // Calculate appointment start and end datetimes (end = +1hr from start)
        const CustName = agent.parameters.CustName;
        const AppointmentDate = agent.parameters.Date.split('T')[0];
        const AppointmentTime = agent.parameters.Time.split('T')[1].substr(0, 8);
        const dateTimeStart = new Date(AppointmentDate + 'T' + AppointmentTime + timeZoneOffset);
        const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));

        var CalendarId;
        var Stylist;
        Stylist = agent.parameters.Stylist;
        if (Stylist == "Jackie") { CalendarId = CalendarIdJackie; }
        if (Stylist == "Samantha") { CalendarId = CalendarIdSamantha; }
        if (Stylist == "Brian") { CalendarId = CalendarIdBrian; }

        // Check if customer indeed made the appointment and delete calendar event accordingly
        return deleteCalendarEvent(CustName, dateTimeStart, dateTimeEnd, CalendarId).then(() => {
            agent.add(`Appointment Deleted`);
            agent.setFollowupEvent('SuccessDeleteAppt');
        }).catch(() => {
            agent.add(`No appointment found in requested time`);
            agent.setFollowupEvent('FailureDeleteAppt');
        });
    }

    function deleteCalendarEvent(CustName, dateTimeStart, dateTimeEnd, CalendarId) {
        return new Promise((resolve, reject) => {
            calendar.events.list({
                auth: serviceAccountAuth, // List events for time period
                calendarId: CalendarId,
                timeMin: dateTimeStart.toISOString(),
                timeMax: dateTimeEnd.toISOString()
            }, (err, calendarResponse) => {
                // Check if there is a event on the Calendar
                if (calendarResponse.data.items.length != 0) {
                    var ApptCustName = calendarResponse.data.items[0].summary;
                }
                if (err || (calendarResponse.data.items.length == 0) || (ApptCustName != CustName)) {
                    reject(err || new Error('No appointment by this Customer at requested time'));
                } else {
                    // Delete event for the requested time period
                    calendar.events.delete({
                        auth: serviceAccountAuth,
                        calendarId: CalendarId,
                        eventId: calendarResponse.data.items[0].id
                    }, (err, event) => {
                        err ? reject(err) : resolve(event);
                    }
                    );
                }
            });
        });
    }

    function GiveFeedBackHandler(agent) {
        const DateService = agent.parameters.DateService.split('T')[0].split('-').reverse().join('-');
        const DateCurrent = new Date().toISOString().split('T')[0].split('-').reverse().join('-');
        const Stylist = agent.parameters.Stylist;
        const Rating = agent.parameters.Rating;
        const Comment = agent.parameters.Comment;
        const CustName = agent.parameters.CustName;

        db.collection("FeedbackDB").add({ DateCurrent: DateCurrent, DateService: DateService, Stylist: Stylist, Rating: Rating, Comment: Comment, CustName: CustName });
        agent.add(`Thank you, ${CustName} for your feedback. Your feedback has been submitted. You can start again, or end call. Have a nice day!`);
    }

    function ShampooHandler(agent) {
        const DateCurrent = new Date().toISOString().split('T')[0].split('-').reverse().join('-');
        const CollectDate = agent.parameters.CollectDate.split('T')[0].split('-').reverse().join('-');
        const ShampooQty = agent.parameters.ShampooQty;
        const CustName = agent.parameters.CustName;
        const CustMobile = agent.parameters.CustMobile;

        // Code for sending email to order fulfillment team
        sendEmail(FulfillmentEmail, CustName, CustMobile, ShampooQty, CollectDate);

        // Write order document to Firestore DB
        db.collection("ShampooOrderDB").add({ DateCurrent: DateCurrent, CollectDate: CollectDate, ShampooQty: ShampooQty, CustName: CustName, CustMobile: CustMobile });
        agent.add(`Thank you, ${CustName} for your order. We will get ready the ${ShampooQty} bottles of shampoo for your collection on ${CollectDate}.You can start again or end call. Have a nice day!`);
    }

    // Send email to the Shampoo Order Fulfillment Team
    function sendEmail(RecipientEmail, CustName, CustMobile, ShampooQty, CollectDate) {
        const mailOptions = {
            from: ChatbotEmail,
            to: RecipientEmail
        };
        mailOptions.subject = 'New Shampoo Order from ' + CustName;
        mailOptions.html = 'Dear Shampoo Order Fulfillment Team' +
            '<p> This is an automatically generated email for Shampoo Order.' +
            '<p> Please be informed that there is a new order as follows.' +
            '<p><strong> Customer Name: </strong>' + CustName +
            '<br><strong> Customer Mobile: </strong>' + CustMobile +
            '<br><strong> Shampoo Quantity: </strong>' + ShampooQty +
            '<br><strong> Date of Collection: </strong>' + CollectDate +
            '<p> Please get ready the shampoo for collection. <p> Thank you.' +
            '<p> From: Jackie Hair Salon Chatbot';
        return mailTransport.sendMail(mailOptions).then(() => {
            return console.log('Email sent to:', RecipientEmail);
        });
    }

    let intentMap = new Map();
    intentMap.set('MakeAppt', makeAppointment);
    intentMap.set('GiveFeedBack', GiveFeedBackHandler);
    intentMap.set('ChangeStylist', makeAppointment);
    intentMap.set('ChangeDateTime', makeAppointment);
    intentMap.set('ChangeTime', makeAppointment);
    intentMap.set('OrderShampoo', ShampooHandler);
    intentMap.set('AddShampoo', ShampooHandler);
    intentMap.set('DeleteAppt', DeleteApptHandler);
    agent.handleRequest(intentMap);
});
package com.itechart.mwssample.service;


import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.availability.AvailabilityData;
import microsoft.exchange.webservices.data.core.enumeration.misc.error.ServiceError;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.response.AttendeeAvailability;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.misc.availability.AttendeeInfo;
import microsoft.exchange.webservices.data.misc.availability.GetUserAvailabilityResults;
import microsoft.exchange.webservices.data.misc.availability.TimeWindow;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.EmailAddressCollection;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.complex.availability.CalendarEvent;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.List;

public class AppointmentService {

    private final static String EMAIL_FOR_ROOM_LIST = "building.tolstogo@itechart-group.com";

    public void createAppointment(ExchangeService service, String subject, String body, Date startDate, Date endDate, String location, String roomEmail) throws Exception{
        Appointment appointment = new  Appointment(service);
        appointment.setSubject(subject);
        appointment.setBody(MessageBody.getMessageBodyFromText(body));

        appointment.setStart(startDate);
        appointment.setEnd(endDate);

        if (location!=null && roomEmail!=null){
            appointment.setLocation(location);
            appointment.getRequiredAttendees().add(roomEmail);

        }
        appointment.save();

        appointment.load();
        Thread.sleep(100000);
        System.out.println(appointment.getMyResponseType());
    }

    public void printRoomEvents(ExchangeService service, String roomEmail, Date startDate, Date endDate) throws Exception {
        List<AttendeeInfo> attendees = new ArrayList<AttendeeInfo>();
        attendees.add(new AttendeeInfo(roomEmail));

        GetUserAvailabilityResults results = service.getUserAvailability(
                attendees,
                new TimeWindow(startDate, endDate),
                AvailabilityData.FreeBusyAndSuggestions);

        for (AttendeeAvailability attendeeAvailability : results.getAttendeesAvailability()) {
            System.out.println("Availability for " + attendees.get(0).getSmtpAddress());
            if (attendeeAvailability.getErrorCode() == ServiceError.NoError) {
            for (CalendarEvent calendarEvent : attendeeAvailability.getCalendarEvents()) {
                System.out.println("Calendar event");
                System.out.println("  Start time: " + calendarEvent.getStartTime().toString());
                System.out.println("  End time: " + calendarEvent.getEndTime().toString());

                if (calendarEvent.getDetails() != null)
                {
                    System.out.println("  Subject: " + calendarEvent.getDetails().getSubject());
                }
            }
            }
        }
    }

    public void createRecurringAppointment(ExchangeService service, String subject, String body,
                                           Date startDate, Date endDate, Date recurrenceEndDate) throws Exception{
        Appointment appointment = new Appointment(service);
        appointment.setSubject(subject);
        appointment.setBody(MessageBody.getMessageBodyFromText(body));

        appointment.setStart(startDate);
        appointment.setEnd(endDate);

        //From the date of the first meeting, 3 days between each occurrence.
        appointment.setRecurrence(new Recurrence.DailyPattern(appointment.getStart(), 3));

        appointment.getRecurrence().setStartDate(appointment.getStart());
        appointment.getRecurrence().setEndDate(recurrenceEndDate);
        appointment.save();
    }


    public List<Appointment> findAppointments(ExchangeService service, Date startDate, Date endDate) throws Exception{
        CalendarFolder cf=CalendarFolder.bind(service, WellKnownFolderName.Calendar);
        FindItemsResults<Appointment> appointments = cf.findAppointments(new CalendarView(startDate, endDate));
        return appointments.getItems();
    }

    public void printAppointments(List<Appointment> appointments) throws ServiceLocalException {
        for (Appointment appointment : appointments) {
            System.out.println("\nAPPOINTMENT:");
            System.out.println("Id: " + appointment.getId().toString());
            System.out.println("Subject: " + appointment.getSubject());
            System.out.println("Start: " + appointment.getStart());
            System.out.println("End: " + appointment.getEnd());
            System.out.println("Recurring: " + appointment.getIsRecurring());
        }
    }

    public Collection<EmailAddress> getOrganizationRooms(ExchangeService service)throws Exception{
        return service.getRooms(new EmailAddress(EMAIL_FOR_ROOM_LIST));
    }

    public void printRoomList(ExchangeService service) throws Exception {
        for (EmailAddress email : getOrganizationRooms(service)) {
            System.out.println("\nROOM:");
            System.out.println("Name: " + email.getName());
            System.out.println("Address: " + email.getAddress());
            System.out.println("Routing type: " + email.getRoutingType());
        }
    }

}

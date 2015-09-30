package com.itechart.mwssample.service;


import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;

import java.util.Date;
import java.util.List;

public class AppointmentService {

    public void createAppointment(ExchangeService service, String subject, String body, Date startDate, Date endDate) throws Exception{
        Appointment appointment = new  Appointment(service);
        appointment.setSubject(subject);
        appointment.setBody(MessageBody.getMessageBodyFromText(body));

        appointment.setStart(startDate);
        appointment.setEnd(endDate);

        appointment.save();
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



}

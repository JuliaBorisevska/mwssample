package com.itechart.mwssample.main;


import com.itechart.mwssample.service.AppointmentService;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;

import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

//Getting Started Guide: https://github.com/OfficeDev/ews-java-api/wiki/Getting-Started-Guide

public class App
{
    private final static String EXCHANGE_SERVICE_URL = "https://webmail.itechart-group.com/ews/Exchange.asmx";
    private final static String USERNAME = "yulia.baryseuskaya";
    private final static String PASSWORD = "qwerty";

    public static void main( String[] args )
    {
        try (ExchangeService service = new ExchangeService()) {

            ExchangeCredentials credentials = new WebCredentials(USERNAME, PASSWORD);
            service.setCredentials(credentials);
            URI uri = new URI(EXCHANGE_SERVICE_URL);
            service.setUrl(uri);

            AppointmentService appointmentService = new AppointmentService();
            SimpleDateFormat formatter = new  SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Date startDate = formatter.parse("2015-10-06 16:00:00");
            Date endDate = formatter.parse("2015-10-06 16:30:00");

            //appointmentService.createAppointment(service, "JAVA TEST Appointment","Test Body Msg",startDate,endDate);
            for (EmailAddress address : appointmentService.getOrganizationRooms(service)){
                if ("Room 1002-2".equals(address.getName())){
                    appointmentService.createAppointment(service, "TEST","",startDate,endDate,address.getName(),address.getAddress());
                }
            }

            formatter = new SimpleDateFormat("yyyy-MM-dd");
            Date recurrenceEndDate = formatter.parse("2015-12-01");
//            appointmentService.createRecurringAppointment(service, "JAVA TEST Recurring Appointment","Test Body Msg",
//                    startDate,endDate,recurrenceEndDate);

            List<Appointment> appointments = appointmentService.findAppointments(service,
                    startDate, formatter.parse("2015-11-01 10:00:00"));
            //appointmentService.printAppointments(appointments);

            //appointmentService.printRoomList(service);

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}

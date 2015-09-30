package com.itechart.mwssample.main;


import com.itechart.mwssample.service.AppointmentService;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;

import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

//Getting Started Guide: https://github.com/OfficeDev/ews-java-api/wiki/Getting-Started-Guide

public class App
{
    private final static String EXCHANGE_SERVICE_URL = "https://webmail.itechart-group.com/ews/Exchange.asmx";
    private final static String USERNAME = "yulia.baryseuskaya";
    private final static String PASSWORD = "123qwerty";

    public static void main( String[] args )
    {
        try (ExchangeService service = new ExchangeService()) {

            ExchangeCredentials credentials = new WebCredentials(USERNAME, PASSWORD);
            service.setCredentials(credentials);
            URI uri = new URI(EXCHANGE_SERVICE_URL);
            service.setUrl(uri);

            AppointmentService appointmentService = new AppointmentService();
            SimpleDateFormat formatter = new  SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Date startDate = formatter.parse("2015-10-01 12:00:00");
            Date endDate = formatter.parse("2015-10-01 13:00:00");

            appointmentService.createAppointment(service, "JAVA TEST Appointment","Test Body Msg",startDate,endDate);

            formatter = new SimpleDateFormat("yyyy-MM-dd");
            Date recurrenceEndDate = formatter.parse("2015-12-01");
            appointmentService.createRecurringAppointment(service, "JAVA TEST Recurring Appointment","Test Body Msg",
                    startDate,endDate,recurrenceEndDate);

            List<Appointment> appointments = appointmentService.findAppointments(service,
                    startDate, formatter.parse("2015-11-01 10:00:00"));
            appointmentService.printAppointments(appointments);

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}

<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/807435054/23.2.6%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1235328)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Synchronize User Appointments with Microsoft EWS using Exchange Online

This demo application synchronizes user appointments with the Microsoft Exchange Web Service (EWS) using Exchange Online (bi-directionally). The application itself exports appointments from the DevExpress WinForms Scheduler control to EWS and imports EWS events to our award-winning WinForms Scheduler control.

## Get Started (Prerequisites)

Install the following NuGet packages:

  * Microsoft.Exchange.WebServices
  * Microsoft.Identity.Client

Register your application in Azure (see [Register a client application in Microsoft Entra ID](https://learn.microsoft.com/en-us/azure/healthcare-apis/register-application) if you have yet to register apps on Azure). Once registered, specify `tenantId` and `clientId` in *App.config*:

```
<appSettings>
  <add key="clientId" value="YOUR_CLIENT_ID"/>
  <add key="tenantId" value="YOUT_TENANT_ID"/>
</appSettings>
```

## Implementation Details

The `GetExchangeWebService` method (*Authorization.cs*) implements authorization and obtains an [ExchangeService](https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.exchangeservice?view=exchange-ews-api) object.

> **NOTE**
>
> If EWS is installed on a server within a domain, review the following topic to authorize and obtain the ExchangeService object: [Get started with EWS Managed API client applications](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications).

The `Synchronizer` class (*Synchronizer.cs*) implements sync APIs:

*	`GetExchangeAppointments(startSearchDate, endSearchDate)` — Returns standard and recurring EWS appointments for the specified time interval.
*	`ExchangeToScheduler(appointmentStorage, exchangeAppointment)` — Imports the specified EWS appointment to the DevExpress WinForms Scheduler.
*	`SchedulerToExchange(userAppointment)` — Exports the specified user appointment from the DevExpress Scheduler to EWS.

## Limitations

Synchronization does not support [changed occurrences](https://docs.devexpress.com/WindowsForms/1753/controls-and-libraries/scheduler/appointments#appointment-types) if associated duration does not fall within the specified time interval.

## Files to Review

* [Authorization.cs](./CS/EWSSyncExample/Authorization.cs)
* [Synchronizer.cs](./CS/EWSSyncExample/Synchronizer.cs)
* [ExchangeExampleForm.cs](./CS/EWSSyncExample/ExchangeExampleForm.cs)

## More Examples

* [Synchronize User Appointments with Microsoft 365 Calendars](https://github.com/DevExpress-Examples/winforms-scheduler-synchronize-appointments-with-outlook-365)

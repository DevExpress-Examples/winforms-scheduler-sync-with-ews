using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

using ExchangeAppointment = Microsoft.Exchange.WebServices.Data.Appointment;
using SchedulerAppointment = DevExpress.XtraScheduler.Appointment;
using ExchangeWeekOfMonth = Microsoft.Exchange.WebServices.Data.DayOfTheWeekIndex;
using SchedulerWeekOfMonth = DevExpress.XtraScheduler.WeekOfMonth;
using ExchangeWeekDay = Microsoft.Exchange.WebServices.Data.DayOfTheWeek;
using SchedulerWeekDay = DevExpress.XtraScheduler.WeekDays;
using Recurrence = Microsoft.Exchange.WebServices.Data.Recurrence;

namespace EWSSyncExample {
    public class Synchronizer {
        protected ExchangeService service;

        public Synchronizer(ExchangeService service) {
            this.service = service;
        }

        PropertySet recurrentPropertySet;
        PropertySet normalPropertySet;

        protected PropertySet GetPropertySet(bool recurrent) {
            if(recurrent) {
                if(recurrentPropertySet == null)
                    recurrentPropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    AppointmentSchema.AppointmentType,
                    AppointmentSchema.Subject,
                    AppointmentSchema.TextBody,
                    AppointmentSchema.FirstOccurrence,
                    AppointmentSchema.LastOccurrence,
                    AppointmentSchema.ModifiedOccurrences,
                    AppointmentSchema.DeletedOccurrences,
                    AppointmentSchema.Recurrence,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.ReminderMinutesBeforeStart,
                    AppointmentSchema.IsAllDayEvent,
                    AppointmentSchema.IsReminderSet,
                    AppointmentSchema.IsCancelled);
                return recurrentPropertySet;
            }
            else {
                if(normalPropertySet == null)
                    normalPropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    AppointmentSchema.AppointmentType,
                    AppointmentSchema.Subject,
                    AppointmentSchema.TextBody,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.ReminderMinutesBeforeStart,
                    AppointmentSchema.IsAllDayEvent,
                    AppointmentSchema.IsReminderSet,
                    AppointmentSchema.IsCancelled);
                return normalPropertySet;
            }
        }

        public List<ExchangeAppointment> GetExchangeAppointments(DateTime startSearchDate, DateTime endSearchDate) {
            List<ExchangeAppointment> res = new List<ExchangeAppointment>();
            HashSet<string> recurrentIds = new HashSet<string>(); 
            CalendarView calView = new CalendarView(startSearchDate, endSearchDate);
            
            calView.PropertySet = new PropertySet(BasePropertySet.IdOnly,
                                                    AppointmentSchema.AppointmentType,
                                                    AppointmentSchema.IsCancelled,
                                                    AppointmentSchema.ICalUid);

            FindItemsResults<ExchangeAppointment> findResults = service.FindAppointments(WellKnownFolderName.Calendar, calView);
            var singleAptProps = GetPropertySet(false);

            foreach(ExchangeAppointment appt in findResults.Items) {
                if(appt.AppointmentType == AppointmentType.Occurrence || appt.AppointmentType == AppointmentType.Exception) {
                    if(recurrentIds.Add(appt.ICalUid)) {                       
                        res.Add(GetRecurrentTemplate(appt.Id));
                        continue;
                    }
                }

                if(appt.AppointmentType == AppointmentType.Single) {
                    res.Add(ExchangeAppointment.Bind(service, appt.Id, singleAptProps));
                }
            }

            return res.Where(n => !n.IsCancelled).ToList();
        }

        protected ExchangeAppointment GetRecurrentTemplate(ItemId itemId) {
            ExchangeAppointment calendarItem = Appointment.Bind(service, itemId, new PropertySet(AppointmentSchema.AppointmentType));
            PropertySet props = GetPropertySet(true);
            switch(calendarItem.AppointmentType) {                
                case AppointmentType.RecurringMaster:
                    return ExchangeAppointment.Bind(service, itemId, props);                    
                case AppointmentType.Occurrence:
                case AppointmentType.Exception:
                    return ExchangeAppointment.BindToRecurringMaster(service, itemId, props);                    
                default:
                    return null;
            }            
        }

        void ConvertRecurrenceMSToDX(ExchangeAppointment exchangeAppointment, SchedulerAppointment schedulerAppointment) {

            if(exchangeAppointment.Recurrence is Recurrence.YearlyPattern yearlyPattern) {                
                schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Yearly;
                schedulerAppointment.RecurrenceInfo.WeekOfMonth = 0;
                schedulerAppointment.RecurrenceInfo.DayNumber = yearlyPattern.DayOfMonth;
                schedulerAppointment.RecurrenceInfo.Month = (int)yearlyPattern.Month;
            }

            if(exchangeAppointment.Recurrence is Recurrence.RelativeYearlyPattern relativeYearlyPattern) {
                schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Yearly;
                schedulerAppointment.RecurrenceInfo.WeekOfMonth = relativeYearlyPattern.DayOfTheWeekIndex.ToSchedulerWeekOfMonth();
                schedulerAppointment.RecurrenceInfo.WeekDays = relativeYearlyPattern.DayOfTheWeek.ToSchedulerWeekDay();
                schedulerAppointment.RecurrenceInfo.Month = (int)relativeYearlyPattern.Month;
            }

            if(exchangeAppointment.Recurrence is Recurrence.MonthlyPattern monthlyPattern) {
                schedulerAppointment.RecurrenceInfo.Periodicity = monthlyPattern.Interval;
                schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Monthly;
                schedulerAppointment.RecurrenceInfo.WeekOfMonth = 0;
                schedulerAppointment.RecurrenceInfo.DayNumber = monthlyPattern.DayOfMonth;
            }

            if(exchangeAppointment.Recurrence is Recurrence.RelativeMonthlyPattern relativeMonthlyPattern) {            
                schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Monthly;
                schedulerAppointment.RecurrenceInfo.WeekOfMonth = relativeMonthlyPattern.DayOfTheWeekIndex.ToSchedulerWeekOfMonth();
                schedulerAppointment.RecurrenceInfo.WeekDays = relativeMonthlyPattern.DayOfTheWeek.ToSchedulerWeekDay();
            }

            if(exchangeAppointment.Recurrence is Recurrence.WeeklyPattern weeklyPattern) {
                schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Weekly;
                schedulerAppointment.RecurrenceInfo.WeekDays = weeklyPattern.DaysOfTheWeek.ToSchedulerWeekDays();
            }

            if(exchangeAppointment.Recurrence is Recurrence.DailyPattern)
                schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Daily;

            if(exchangeAppointment.Recurrence is Recurrence.IntervalPattern)
                schedulerAppointment.RecurrenceInfo.Periodicity = (exchangeAppointment.Recurrence as Recurrence.IntervalPattern).Interval;

            schedulerAppointment.RecurrenceInfo.Start = exchangeAppointment.Recurrence.StartDate.Add(exchangeAppointment.Start.TimeOfDay);

            if(exchangeAppointment.Recurrence.HasEnd) {
                if(exchangeAppointment.Recurrence.NumberOfOccurrences.HasValue) {
                    schedulerAppointment.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.OccurrenceCount;
                    schedulerAppointment.RecurrenceInfo.OccurrenceCount = exchangeAppointment.Recurrence.NumberOfOccurrences.Value;
                }
                if(exchangeAppointment.Recurrence.EndDate.HasValue) {
                    schedulerAppointment.RecurrenceInfo.End = exchangeAppointment.Recurrence.EndDate.Value;
                    schedulerAppointment.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.EndByDate;
                }
            }
            else
                schedulerAppointment.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.NoEndDate;

            if((exchangeAppointment.DeletedOccurrences?.Count > 0) || (exchangeAppointment.ModifiedOccurrences?.Count > 0)) {
                var calculator = DevExpress.XtraScheduler.OccurrenceCalculator.CreateInstance(schedulerAppointment.RecurrenceInfo);

                if(exchangeAppointment.DeletedOccurrences?.Count > 0)
                    foreach(DeletedOccurrenceInfo deleted in exchangeAppointment.DeletedOccurrences) {
                        int occurrenceIndex = calculator.FindOccurrenceIndex(deleted.OriginalStart, schedulerAppointment);
                        schedulerAppointment.CreateException(DevExpress.XtraScheduler.AppointmentType.DeletedOccurrence, occurrenceIndex);
                    }

                if(exchangeAppointment.ModifiedOccurrences?.Count > 0)
                    foreach(OccurrenceInfo modified in exchangeAppointment.ModifiedOccurrences) {
                        int occurrenceIndex = calculator.FindOccurrenceIndex(modified.OriginalStart, schedulerAppointment);
                        SchedulerAppointment modifiedSchedulerAppointment = schedulerAppointment.CreateException(DevExpress.XtraScheduler.AppointmentType.ChangedOccurrence, occurrenceIndex);
                        modifiedSchedulerAppointment.Start = modified.Start;
                        modifiedSchedulerAppointment.End = modified.End;
                    }
            }
        }

        void ConvertRecurrenceDXtoMS(SchedulerAppointment schedulerAppointment, ExchangeAppointment exchangeAppointment) {

            switch(schedulerAppointment.RecurrenceInfo.Type) {
                case DevExpress.XtraScheduler.RecurrenceType.Yearly:
                    if(schedulerAppointment.RecurrenceInfo.WeekOfMonth == 0) {
                        Recurrence.YearlyPattern pattern = new Recurrence.YearlyPattern();
                        pattern.DayOfMonth = schedulerAppointment.RecurrenceInfo.DayNumber;
                        pattern.Month = (Month)schedulerAppointment.RecurrenceInfo.Month;
                        exchangeAppointment.Recurrence = pattern;
                    }
                    else {
                        Recurrence.RelativeYearlyPattern pattern = new Recurrence.RelativeYearlyPattern();
                        pattern.DayOfTheWeekIndex = schedulerAppointment.RecurrenceInfo.WeekOfMonth.ToExchangeWeekOfMonth();
                        pattern.DayOfTheWeek = schedulerAppointment.RecurrenceInfo.WeekDays.ToExchangeWeekDays().First();
                        pattern.Month = (Month)schedulerAppointment.RecurrenceInfo.Month;
                        exchangeAppointment.Recurrence = pattern;
                    }
                    break;
                case DevExpress.XtraScheduler.RecurrenceType.Monthly:
                    if(schedulerAppointment.RecurrenceInfo.WeekOfMonth == 0) {
                        Recurrence.MonthlyPattern pattern = new Recurrence.MonthlyPattern();
                        pattern.DayOfMonth = schedulerAppointment.RecurrenceInfo.DayNumber;
                        exchangeAppointment.Recurrence = pattern;
                    }
                    else {
                        Recurrence.RelativeMonthlyPattern pattern = new Recurrence.RelativeMonthlyPattern();
                        pattern.DayOfTheWeekIndex = schedulerAppointment.RecurrenceInfo.WeekOfMonth.ToExchangeWeekOfMonth();
                        pattern.DayOfTheWeek = schedulerAppointment.RecurrenceInfo.WeekDays.ToExchangeWeekDays().First();
                        exchangeAppointment.Recurrence = pattern;
                    }
                    break;
                case DevExpress.XtraScheduler.RecurrenceType.Weekly:
                    exchangeAppointment.Recurrence = new Recurrence.WeeklyPattern(schedulerAppointment.RecurrenceInfo.Start,
                    schedulerAppointment.RecurrenceInfo.Periodicity,
                    schedulerAppointment.RecurrenceInfo.WeekDays.ToExchangeWeekDays());
                    break;
                case DevExpress.XtraScheduler.RecurrenceType.Daily:
                    exchangeAppointment.Recurrence = new Recurrence.DailyPattern(schedulerAppointment.RecurrenceInfo.Start,
                    schedulerAppointment.RecurrenceInfo.Periodicity);
                    break;
                default:
                    throw new NotImplementedException("Recurrence type " + schedulerAppointment.RecurrenceInfo.Type + " is not supported.");
            }

            if(exchangeAppointment.Recurrence is Recurrence.IntervalPattern)
                (exchangeAppointment.Recurrence as Recurrence.IntervalPattern).Interval = schedulerAppointment.RecurrenceInfo.Periodicity;

            exchangeAppointment.Recurrence.StartDate = schedulerAppointment.RecurrenceInfo.Start.Date;

            switch(schedulerAppointment.RecurrenceInfo.Range) {
                case DevExpress.XtraScheduler.RecurrenceRange.OccurrenceCount:
                    exchangeAppointment.Recurrence.NumberOfOccurrences = schedulerAppointment.RecurrenceInfo.OccurrenceCount;
                    break;
                case DevExpress.XtraScheduler.RecurrenceRange.EndByDate:
                    exchangeAppointment.Recurrence.EndDate = schedulerAppointment.RecurrenceInfo.End;
                    break;
                case DevExpress.XtraScheduler.RecurrenceRange.NoEndDate:
                    break;
                default:
                    throw new NotImplementedException("Recurrence range type " + schedulerAppointment.RecurrenceInfo.Range.ToString() + " is not supported.");
            }
        }

        public void SchedulerToExchange(SchedulerAppointment schedulerAppointment, bool addReminders = false) {
            if((schedulerAppointment.Type != DevExpress.XtraScheduler.AppointmentType.Pattern)
                && (schedulerAppointment.Type != DevExpress.XtraScheduler.AppointmentType.Normal))
                throw new Exception("Appointment type " + schedulerAppointment.Type + " is not support");

            ExchangeAppointment exchangeAppointment = new ExchangeAppointment(service);
            exchangeAppointment.Subject = schedulerAppointment.Subject;
            exchangeAppointment.Body = schedulerAppointment.Description;
            exchangeAppointment.Start = schedulerAppointment.Start;
            exchangeAppointment.End = schedulerAppointment.End;

            if((addReminders) && (schedulerAppointment.HasReminder))
                exchangeAppointment.ReminderMinutesBeforeStart = schedulerAppointment.Reminder.TimeBeforeStart.Minutes;

            if(schedulerAppointment.Type == DevExpress.XtraScheduler.AppointmentType.Pattern)
                ConvertRecurrenceDXtoMS(schedulerAppointment, exchangeAppointment);

            exchangeAppointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);

            if(schedulerAppointment.HasExceptions) {
                foreach(SchedulerAppointment exception in schedulerAppointment.GetExceptions()) {
                    if(exception.Type == DevExpress.XtraScheduler.AppointmentType.DeletedOccurrence) {
                        ExchangeAppointment appointment = ExchangeAppointment.BindToOccurrence(service, exchangeAppointment.Id, exception.RecurrenceIndex + 1, new PropertySet());
                        appointment.Delete(DeleteMode.MoveToDeletedItems);
                        continue;
                    }

                    if(exception.Type == DevExpress.XtraScheduler.AppointmentType.ChangedOccurrence) {
                        ExchangeAppointment appointment = ExchangeAppointment.BindToOccurrence(service, exchangeAppointment.Id, exception.RecurrenceIndex + 1,
                            new PropertySet(AppointmentSchema.Start,
                                            AppointmentSchema.End,
                                            AppointmentSchema.Subject,
                                            AppointmentSchema.OriginalStart));
                        appointment.Start = exception.Start;
                        appointment.End = exception.End;
                        appointment.Subject = exception.Subject;
                        appointment.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy);
                        continue;
                    }
                }
            }
        }

        public void ExchangeToScheduler(DevExpress.XtraScheduler.IAppointmentStorageBase appointmentStorage, ExchangeAppointment exchangeAppointment, bool addReminders = false) {

            SchedulerAppointment schedulerAppointment = null;
            if(exchangeAppointment.AppointmentType == AppointmentType.Single)
                schedulerAppointment = appointmentStorage.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Normal);
            else
                schedulerAppointment = appointmentStorage.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Pattern);

            schedulerAppointment.AllDay = exchangeAppointment.IsAllDayEvent;
            schedulerAppointment.Start = exchangeAppointment.Start;
            schedulerAppointment.End = exchangeAppointment.End;
            schedulerAppointment.Subject = exchangeAppointment.Subject;
            schedulerAppointment.Description = exchangeAppointment.TextBody;

            if(exchangeAppointment.AppointmentType != AppointmentType.Single)
                ConvertRecurrenceMSToDX(exchangeAppointment, schedulerAppointment);

            if(addReminders) {
                DevExpress.XtraScheduler.Reminder reminder = schedulerAppointment.CreateNewReminder();
                reminder.TimeBeforeStart = TimeSpan.FromMinutes(exchangeAppointment.ReminderMinutesBeforeStart);
                schedulerAppointment.Reminders.Add(reminder);
            }

            appointmentStorage.Add(schedulerAppointment);
        }
    }
}

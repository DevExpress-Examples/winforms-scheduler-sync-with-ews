using System;
using System.Collections.Generic;

using ExchangeWeekOfMonth = Microsoft.Exchange.WebServices.Data.DayOfTheWeekIndex;
using SchedulerWeekOfMonth = DevExpress.XtraScheduler.WeekOfMonth;
using ExchangeWeekDay = Microsoft.Exchange.WebServices.Data.DayOfTheWeek;
using SchedulerWeekDay = DevExpress.XtraScheduler.WeekDays;

using Microsoft.Exchange.WebServices.Data;

namespace EWSSyncExample {
    public static class ExchangeSchedulerExtConverters {
        public static SchedulerWeekOfMonth ToSchedulerWeekOfMonth(this ExchangeWeekOfMonth weekIndex) {
            switch(weekIndex) {
                case ExchangeWeekOfMonth.First:
                    return SchedulerWeekOfMonth.First;
                case ExchangeWeekOfMonth.Second:
                    return SchedulerWeekOfMonth.Second;
                case ExchangeWeekOfMonth.Third:
                    return SchedulerWeekOfMonth.Third;
                case ExchangeWeekOfMonth.Fourth:
                    return SchedulerWeekOfMonth.Fourth;
                case ExchangeWeekOfMonth.Last:
                    return SchedulerWeekOfMonth.Last;
            }
            throw new Exception("Incorrect WeekIndex: " + weekIndex.ToString());
        }
        
        public static ExchangeWeekOfMonth ToExchangeWeekOfMonth(this SchedulerWeekOfMonth weekIndex) {
            switch(weekIndex) {
                case SchedulerWeekOfMonth.First:
                    return ExchangeWeekOfMonth.First;
                case SchedulerWeekOfMonth.Second:
                    return ExchangeWeekOfMonth.Second;
                case SchedulerWeekOfMonth.Third:
                    return ExchangeWeekOfMonth.Third;
                case SchedulerWeekOfMonth.Fourth:
                    return ExchangeWeekOfMonth.Fourth;
                case SchedulerWeekOfMonth.Last:
                    return ExchangeWeekOfMonth.Last;
            }
            throw new Exception("Incorrect WeekIndex: " + weekIndex.ToString());
        }

        public static ExchangeWeekDay[] ToExchangeWeekDays(this SchedulerWeekDay dw) {
            List<ExchangeWeekDay> result = new List<ExchangeWeekDay>();
            if((dw & SchedulerWeekDay.Sunday) != 0) result.Add(ExchangeWeekDay.Sunday);
            if((dw & SchedulerWeekDay.Monday) != 0) result.Add(ExchangeWeekDay.Monday);
            if((dw & SchedulerWeekDay.Tuesday) != 0) result.Add(ExchangeWeekDay.Tuesday);
            if((dw & SchedulerWeekDay.Wednesday) != 0) result.Add(ExchangeWeekDay.Wednesday);
            if((dw & SchedulerWeekDay.Thursday) != 0) result.Add(ExchangeWeekDay.Thursday);
            if((dw & SchedulerWeekDay.Friday) != 0) result.Add(ExchangeWeekDay.Friday);
            if((dw & SchedulerWeekDay.Saturday) != 0) result.Add(ExchangeWeekDay.Saturday);
            if((dw & SchedulerWeekDay.EveryDay) != 0) result.Add(ExchangeWeekDay.Day);
            if((dw & SchedulerWeekDay.WeekendDays) != 0) result.Add(ExchangeWeekDay.WeekendDay);
            if((dw & SchedulerWeekDay.WorkDays) != 0) result.Add(ExchangeWeekDay.Weekday);
            return result.ToArray();
        }

        public static SchedulerWeekDay ToSchedulerWeekDay(this ExchangeWeekDay dw) {
            switch(dw) {
                case ExchangeWeekDay.Sunday:
                    return SchedulerWeekDay.Sunday;
                case ExchangeWeekDay.Monday:
                    return SchedulerWeekDay.Monday;
                case ExchangeWeekDay.Tuesday:
                    return SchedulerWeekDay.Tuesday;
                case ExchangeWeekDay.Wednesday:
                    return SchedulerWeekDay.Wednesday;
                case ExchangeWeekDay.Thursday:
                    return SchedulerWeekDay.Thursday;
                case ExchangeWeekDay.Friday:
                    return SchedulerWeekDay.Friday;
                case ExchangeWeekDay.Saturday:
                    return SchedulerWeekDay.Saturday;
                case ExchangeWeekDay.Day:
                    return SchedulerWeekDay.EveryDay;
                case ExchangeWeekDay.WeekendDay:
                    return SchedulerWeekDay.WeekendDays;
                case ExchangeWeekDay.Weekday:
                    return SchedulerWeekDay.WorkDays;
            }
            throw new Exception("Incorrect day of the week: " + dw.ToString());
        }
        public static SchedulerWeekDay ToSchedulerWeekDays(this DayOfTheWeekCollection dws) {
            SchedulerWeekDay result = 0;
            foreach(ExchangeWeekDay dw in dws)
                result = result | dw.ToSchedulerWeekDay();
            return result;
        }
    }
}

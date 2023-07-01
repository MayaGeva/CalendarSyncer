using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class RecurrenceConverter
    {
        public static AppointmentRecurrence ConvertRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            return pattern.RecurrenceType switch
            {
                OlRecurrenceType.olRecursDaily => GetDailyRecurrence(item),
                OlRecurrenceType.olRecursWeekly => GetWeeklyRecurrence(item),
                OlRecurrenceType.olRecursMonthly => GetMonthlyOnDayRecurrence(item),
                OlRecurrenceType.olRecursMonthNth => GetMonthlyRecurrence(item),
                OlRecurrenceType.olRecursYearly => GetYearlyOnDayRecurrence(item),
                OlRecurrenceType.olRecursYearNth => GetYearlyRecurrence(item),
                _ => new AppointmentRecurrence(),
            };
        }
        static AppointmentRecurrence GetDailyRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            AppointmentRecurrence result = new AppointmentRecurrence()
            {
                Unit = AppointmentRecurrenceUnit.Daily,
                Interval = (uint)pattern.Interval,
            };
            if (!pattern.NoEndDate)
            {
                result.Until = pattern.EndTime;
                result.Occurrences = (uint)pattern.Occurrences;
            }
            return result;
        }
        static AppointmentRecurrence GetWeeklyRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            AppointmentRecurrence result = new AppointmentRecurrence()
            {
                DaysOfWeek = (AppointmentDaysOfWeek)pattern.DayOfWeekMask,
                Unit = AppointmentRecurrenceUnit.Weekly,
                Interval = (uint)pattern.Interval,
            };
            if (!pattern.NoEndDate)
            {
                result.Until = pattern.EndTime;
                result.Occurrences = (uint)pattern.Occurrences;
            }
            return result;
        }
        static AppointmentRecurrence GetMonthlyOnDayRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            AppointmentRecurrence result = new AppointmentRecurrence()
            {
                Day = (uint)pattern.DayOfMonth,
                Unit = AppointmentRecurrenceUnit.MonthlyOnDay,
                Interval = (uint)pattern.Interval,
            };
            if (!pattern.NoEndDate)
            {
                result.Until = pattern.EndTime;
                result.Occurrences = (uint)pattern.Occurrences;
            }
            return result;
        }
        static AppointmentRecurrence GetMonthlyRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            AppointmentRecurrence result = new AppointmentRecurrence()
            {
                DaysOfWeek = (AppointmentDaysOfWeek)pattern.DayOfWeekMask,
                Unit = AppointmentRecurrenceUnit.Monthly,
                Interval = (uint)pattern.Interval,
            };
            if (!pattern.NoEndDate)
            {
                result.Until = pattern.EndTime;
                result.Occurrences = (uint)pattern.Occurrences;
            }
            return result;
        }
        static AppointmentRecurrence GetYearlyOnDayRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            AppointmentRecurrence result = new AppointmentRecurrence()
            {
                Day = (uint)pattern.DayOfMonth,
                Month = (uint)pattern.MonthOfYear,
                Unit = AppointmentRecurrenceUnit.YearlyOnDay,
                Interval = (uint)pattern.Interval,
            };
            if (!pattern.NoEndDate)
            {
                result.Until = pattern.EndTime;
                result.Occurrences = (uint)pattern.Occurrences;
            }
            return result;
        }
        static AppointmentRecurrence GetYearlyRecurrence(AppointmentItem item)
        {
            RecurrencePattern pattern = item.GetRecurrencePattern();
            AppointmentRecurrence result = new AppointmentRecurrence()
            {
                DaysOfWeek = (AppointmentDaysOfWeek)pattern.DayOfWeekMask,
                Month = (uint)pattern.MonthOfYear,
                Unit = AppointmentRecurrenceUnit.YearlyOnDay,
                Interval = (uint)pattern.Interval,
            };
            if (!pattern.NoEndDate)
            {
                result.Until = pattern.EndTime;
                result.Occurrences = (uint)pattern.Occurrences;
            }
            return result;
        }
    }
}

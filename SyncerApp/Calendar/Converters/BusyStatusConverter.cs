using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class BusyStatusConverter
    {
        public static AppointmentBusyStatus ConvertBusyStatus(OlBusyStatus status)
        {
            return status switch
            {
                OlBusyStatus.olBusy => AppointmentBusyStatus.Busy,
                OlBusyStatus.olTentative => AppointmentBusyStatus.Tentative,
                OlBusyStatus.olWorkingElsewhere => AppointmentBusyStatus.WorkingElsewhere,
                OlBusyStatus.olFree => AppointmentBusyStatus.Free,
                OlBusyStatus.olOutOfOffice => AppointmentBusyStatus.OutOfOffice,
                _ => AppointmentBusyStatus.Free,
            };
        }
        public OlBusyStatus ConvertBusyStatus(AppointmentBusyStatus status)
        {
            return status switch
            {
                AppointmentBusyStatus.Busy => OlBusyStatus.olBusy,
                AppointmentBusyStatus.Tentative => OlBusyStatus.olTentative,
                AppointmentBusyStatus.WorkingElsewhere => OlBusyStatus.olWorkingElsewhere,
                AppointmentBusyStatus.Free => OlBusyStatus.olFree,
                AppointmentBusyStatus.OutOfOffice => OlBusyStatus.olOutOfOffice,
                _ => OlBusyStatus.olFree,
            };
        }
    }
}

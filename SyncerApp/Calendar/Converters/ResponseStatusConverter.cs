using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class ResponseStatusConverter
    {
        public static AppointmentParticipantResponse ConvertResponseStatus(OlResponseStatus status)
        {
            return status switch
            {
                OlResponseStatus.olResponseAccepted => AppointmentParticipantResponse.Accepted,
                OlResponseStatus.olResponseTentative => AppointmentParticipantResponse.Tentative,
                OlResponseStatus.olResponseDeclined => AppointmentParticipantResponse.Declined,
                OlResponseStatus.olResponseNotResponded => AppointmentParticipantResponse.None,
                _ => AppointmentParticipantResponse.Unknown,
            };
        }
    }
}

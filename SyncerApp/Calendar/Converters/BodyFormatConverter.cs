using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class BodyFormatConverter
    {
        public static AppointmentDetailsKind ConvertBodyFormat(OlBodyFormat bodyFormat)
        {
            return bodyFormat switch
            {
                OlBodyFormat.olFormatHTML => AppointmentDetailsKind.Html,
                OlBodyFormat.olFormatPlain => AppointmentDetailsKind.PlainText,
                _ => AppointmentDetailsKind.PlainText,
            };
        }
        public static OlBodyFormat ConvertBodyFormat(AppointmentDetailsKind bodyFormat)
        {
            return bodyFormat switch
            {
                AppointmentDetailsKind.Html=> OlBodyFormat.olFormatHTML,
                AppointmentDetailsKind.PlainText => OlBodyFormat.olFormatPlain,
                _ => OlBodyFormat.olFormatUnspecified,
            };
        }
    }
}

using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class SensitivityConverter
    {
        public static AppointmentSensitivity ConvertSensitivity(OlSensitivity sensitivity)
        {
            return sensitivity switch
            {
                OlSensitivity.olNormal => AppointmentSensitivity.Public,
                OlSensitivity.olPrivate => AppointmentSensitivity.Private,
                _ => AppointmentSensitivity.Private,
            };
        }

        public static OlSensitivity ConvertSensitivity(AppointmentSensitivity sensitivity)
        {
            return sensitivity switch
            {
                AppointmentSensitivity.Public => OlSensitivity.olNormal,
                AppointmentSensitivity.Private => OlSensitivity.olPrivate,
                _ => OlSensitivity.olNormal,
            };
        }
    }
}

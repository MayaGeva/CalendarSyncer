using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class OrganizerConverter
    {
        public static AppointmentOrganizer ConvertOrganizer(AddressEntry entry)
        {
            return new AppointmentOrganizer()
            {
                DisplayName = entry.Name,
                Address = entry.Address
            };
        }
    }
}

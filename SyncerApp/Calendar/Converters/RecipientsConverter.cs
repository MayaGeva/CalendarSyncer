using Microsoft.Office.Interop.Outlook;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar.Converters
{
    internal class RecipientsConverter
    {
        public static List<AppointmentInvitee> ConvertRecipients(Recipients recipients)
        {
            List<AppointmentInvitee> invitees = new List<AppointmentInvitee>();
            foreach (Recipient recipient in recipients)
            {
                invitees.Add(ConvertRecipient(recipient));
            }
            return invitees;
        }
        public static AppointmentInvitee ConvertRecipient(Recipient recipient)
        {
            return new AppointmentInvitee()
            {
                Address = recipient.Address,
                DisplayName = recipient.Name,
                Response = ResponseStatusConverter.ConvertResponseStatus(recipient.MeetingResponseStatus)
            };
        }
    }
}

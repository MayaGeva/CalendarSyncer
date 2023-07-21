// See https://aka.ms/new-console-template for more information
using OutlookLauncher;
using System.Web;
if (args != null && args.Length > 0) {
    Uri uri = new(args[0]);
    try
    {
        string? entryId = HttpUtility.ParseQueryString(uri.Query).Get("local_id");
        var cal = new OutlookCalendar();
        if (entryId != null)
        {
            var appointment = cal.GetAppointmentFromGlobalId(entryId);
            appointment?.Display();
        }
    }
    catch (ArgumentNullException)
    {

    }
}
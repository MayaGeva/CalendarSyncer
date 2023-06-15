using OutlookListener.Calendar;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.Serialization.Json;
using System.Text;
using Windows.Storage;

namespace OutlookListener
{
    internal class AppointmentSender
    {
        string pipeName;
        public AppointmentSender()
        {
            this.pipeName = $"Sessions\\{Process.GetCurrentProcess().SessionId}\\AppContainerNamedObjects\\{ApplicationData.Current.LocalSettings.Values["PackageSid"]}\\calendar-pipe";
        }
        public void SendAppointment(CalendarAppointment appointment)
        {
            try
            {
                var set = ApplicationData.Current.LocalSettings.Values;
                var pipeClient = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                pipeClient.Connect();
                string appointmentXml = GetAppointmentJson(appointment);

                Console.WriteLine(appointmentXml);
                var appointmentBytes = Encoding.UTF8.GetBytes(appointmentXml);
                pipeClient.Write(appointmentBytes, 0, appointmentBytes.Length);
                pipeClient.Close();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        string GetAppointmentJson(CalendarAppointment appointment)
        {
            using MemoryStream memoryStream = new MemoryStream();
            using StreamReader reader = new StreamReader(memoryStream);
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(CalendarAppointment));
            serializer.WriteObject(memoryStream, appointment);
            memoryStream.Position = 0;
            return reader.ReadToEnd();
        }
    }
}

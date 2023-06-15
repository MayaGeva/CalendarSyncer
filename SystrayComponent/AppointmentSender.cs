using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.Serialization.Json;
using System.Text;
using SystrayComponent.Calendar;
using Windows.Storage;

namespace SystrayComponent
{
    internal class AppointmentSender
    {
        public void SendAppointment(CalendarAppointment appointment)
        {
            try
            {
                string pipeName = $"Sessions\\{Process.GetCurrentProcess().SessionId}\\AppContainerNamedObjects\\{ApplicationData.Current.LocalSettings.Values["PackageSid"]}\\calendar-pipe";
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

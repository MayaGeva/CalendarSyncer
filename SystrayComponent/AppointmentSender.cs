using System.Collections.Concurrent;
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
        BlockingCollection<CalendarAppointment> calendarAppointments;
        public AppointmentSender(BlockingCollection<CalendarAppointment> appointments) 
        {
            calendarAppointments = appointments;
        }

        public async Task HandleAppointments()
        {
            bool keep_running = true;
            while (keep_running)
            {
                CalendarAppointment appointment = calendarAppointments.Take();
                if (appointment.Action == AppointmentDetails.AppointmentAction.NoAction) 
                { 
                    keep_running = false;
                } 
                else
                {
                    await SendAppointment(appointment);
                }
            }
        }

        public async Task SendAppointment(CalendarAppointment appointment)
        {
            try
            {
                string pipeName = $"Sessions\\{Process.GetCurrentProcess().SessionId}\\AppContainerNamedObjects\\{ApplicationData.Current.LocalSettings.Values["PackageSid"]}\\calendar-pipe";
                var pipeClient = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                await pipeClient.ConnectAsync();
                string appointmentXml = await GetAppointmentJson(appointment);

                var appointmentBytes = Encoding.UTF8.GetBytes(appointmentXml);
                await pipeClient.WriteAsync(appointmentBytes);
                pipeClient.Close();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        async Task<string> GetAppointmentJson(CalendarAppointment appointment)
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

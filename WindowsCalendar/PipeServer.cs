using WindowsCalendar.AppointmentDetails;
using System.IO.Pipes;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;
using System.Collections.Concurrent;

namespace WindowsCalendar
{
    internal class PipeServer
    {
        BlockingCollection<CalendarAppointment> appointments;
        public PipeServer(BlockingCollection<CalendarAppointment> calendarAppointments) 
        {
            appointments = calendarAppointments;
        }

        public async Task RunServer()
        {
            while (true)
            {
                NamedPipeServerStream pipeServer = new NamedPipeServerStream(@"LOCAL\calendar-pipe", PipeDirection.InOut, 10, PipeTransmissionMode.Message, PipeOptions.Asynchronous);
                pipeServer.WaitForConnection();
                StreamReader reader = new StreamReader(pipeServer);
                string jsonAppointment = reader.ReadToEnd();
                CalendarAppointment calendarAppointment = await GetCalendarAppointmentFromJson(jsonAppointment);
                if (calendarAppointment != null)
                {
                    appointments.Add(calendarAppointment);
                }
                pipeServer.Disconnect();
            }
        }

        async Task<CalendarAppointment> GetCalendarAppointmentFromJson(string appointmentJson)
        {
            CalendarAppointment calendarAppointment;
            using (Stream stream = new MemoryStream())
            {
                byte[] data = System.Text.Encoding.UTF8.GetBytes(appointmentJson);
                stream.Write(data, 0, data.Length);
                stream.Position = 0;
                DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(CalendarAppointment));
                calendarAppointment = (CalendarAppointment)deserializer.ReadObject(stream);
            }
            return calendarAppointment;
        }
    }
}

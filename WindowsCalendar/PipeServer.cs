using WindowsCalendar.AppointmentDetails;
using System;
using System.IO.Pipes;
using System.IO;
using System.Threading.Tasks;
using Windows.ApplicationModel.Appointments;
using System.Runtime.Serialization.Json;
using WindowsCalendar.Calendar;

namespace WindowsCalendar
{
    internal class PipeServer
    {
        WindowsAppCalendar calendar;
        public PipeServer() 
        {
            calendar = new WindowsAppCalendar();
        }

        public async Task RunServer()
        {
            while (true)
            {
                NamedPipeServerStream pipeServer = new NamedPipeServerStream(@"LOCAL\calendar-pipe", PipeDirection.InOut, 1, PipeTransmissionMode.Message, PipeOptions.Asynchronous);
                pipeServer.WaitForConnection();
                StreamReader reader = new StreamReader(pipeServer);
                string jsonAppointment = reader.ReadToEnd();
                CalendarAppointment calendarAppointment = GetCalendarAppointmentFromJson(jsonAppointment);
                if (calendarAppointment != null)
                {
                    await HandleAppointment(calendarAppointment);
                }
                pipeServer.Disconnect();
            }
        }

        CalendarAppointment GetCalendarAppointmentFromJson(string appointmentJson)
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

        async Task HandleAppointment(CalendarAppointment calendarAppointment)
        {
            Appointment appointment = calendarAppointment.ToAppointment();
            switch (calendarAppointment.Action)
            {
                case AppointmentAction.AddItem:
                    await calendar.AddAppointment(appointment);
                    break;
                case AppointmentAction.RemoveItem:
                    await calendar.RemoveAppointment(appointment);
                    break;
                case AppointmentAction.ChangeItem:
                    await calendar.ModifyAppointment(appointment);
                    break;
                default:
                    throw new NotImplementedException();
            }
        }
    }
}

using Microsoft.Data.Sqlite;
using System;
using System.IO;
using System.Threading.Tasks;
using Windows.Storage;

namespace CalendarPoc
{
    internal class CalendarStorageSettings
    {
        public static string CALENDAR_APPOINTMENTS_FILE_NAME { get; } = "appointmentDetails.db";
        public static readonly string CALENDAR_APPOINTMENTS_TABLE = "AppointmentIds";
        public static readonly string APPOINTMENT_LOCAL_ID = "LocalId";
        public static readonly string APPOINTMENT_ROAMING_ID = "RoamingId";

        string appointments_path;
        public CalendarStorageSettings() 
        {
            appointments_path = Path.Combine(ApplicationData.Current.LocalFolder.Path, CALENDAR_APPOINTMENTS_FILE_NAME);
        }

        public async Task InitializeDatabase()
        {
            await ApplicationData.Current.LocalFolder.CreateFileAsync(CALENDAR_APPOINTMENTS_FILE_NAME, CreationCollisionOption.OpenIfExists);
            using (SqliteConnection db = new SqliteConnection($"Filename={appointments_path}"))
            {
                db.Open();
                string tableCommand = string.Format("CREATE TABLE IF NOT EXISTS " +
                    "{0} ({1} NVARCHAR PRIMARY KEY, " +
                    "{2} NVARCHAR)", CALENDAR_APPOINTMENTS_TABLE, APPOINTMENT_ROAMING_ID, APPOINTMENT_LOCAL_ID);

                SqliteCommand createTable = new SqliteCommand(tableCommand, db);

                createTable.ExecuteReader();
            }
        }

        public void AddLocalIdMapping(string roamingId, string localId)
        {
            using (SqliteConnection db = new SqliteConnection($"Filename={appointments_path}"))
            {
                db.Open();
                string tableCommand = string.Format("INSERT INTO {0} VALUES ({1}, {2});", CALENDAR_APPOINTMENTS_TABLE, roamingId, localId);

                SqliteCommand insertCommand = new SqliteCommand(tableCommand, db);
                SqliteDataReader query = insertCommand.ExecuteReader();
            }
        }

        public string GetLocalIdFromRoamingId(string roamingId)
        {
            using (SqliteConnection db = new SqliteConnection($"Filename={appointments_path}"))
            {
                db.Open();
                string tableCommand = string.Format("SELECT {0} FROM {1} WHERE {2} = '{3}'", APPOINTMENT_LOCAL_ID, CALENDAR_APPOINTMENTS_TABLE, APPOINTMENT_ROAMING_ID, roamingId);

                SqliteCommand selectCommand = new SqliteCommand(tableCommand, db);
                SqliteDataReader query = selectCommand.ExecuteReader();

                return query.GetString(0);
            }

        }

        public void RemoveLocalIdMapping(string roamingId)
        {
            using (SqliteConnection db = new SqliteConnection($"Filename={appointments_path}"))
            {
                db.Open();
                string removeCommand = string.Format("DELETE FROM {0} WHERE {1} = '{2}'", CALENDAR_APPOINTMENTS_TABLE, APPOINTMENT_ROAMING_ID, roamingId);

                SqliteCommand createTable = new SqliteCommand(removeCommand, db);

                createTable.ExecuteReader();
            }
        }
    }
}

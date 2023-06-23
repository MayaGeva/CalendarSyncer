namespace SystrayComponent
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            Mutex mutex = null;
            if (!Mutex.TryOpenExisting("MySystrayExtensionMutex", out mutex))
            {
                mutex = new Mutex(false, "MySystrayExtensionMutex");
                ApplicationConfiguration.Initialize();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
                mutex.Close();
            }
        }
    }
}
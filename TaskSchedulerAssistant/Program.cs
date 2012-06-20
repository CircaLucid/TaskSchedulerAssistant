using System.ServiceProcess;

namespace TaskSchedulerAssistant
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
			{ 
				new TaskSchedulerAssistantService() 
			};
            //((TaskSchedulerAssistantService)ServicesToRun[0]).log(((TaskSchedulerAssistantService)ServicesToRun[0])._xmlPath);
            ServiceBase.Run(ServicesToRun);
        }
    }
}

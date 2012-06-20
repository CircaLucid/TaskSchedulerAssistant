using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;
//using System.Xml;


namespace TaskSchedulerAssistant
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
            ServiceController sc = new ServiceController(TaskSchedulerAssistant.TaskSchedulerAssistantService.ServName);
            sc.Start();
            sc.Dispose();
        }
    }
}

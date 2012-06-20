using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Xml;

namespace TaskSchedulerAssistant
{
    //[System.ComponentModel.DesignerCategory("")]
    public partial class TaskSchedulerAssistantService : ServiceBase
    {
        public const String ServName = "TaskSchedulerAssistant";

        private Timer _xmlReload = new Timer(6000);

        internal bool debug = false;
        internal Dictionary<FileSystemWatcher, String> fsws = new Dictionary<FileSystemWatcher, String>();
        internal FileSystemWatcher fswConfig;
        internal KeyValuePair<String, DateTime> lastRun = new KeyValuePair<String, DateTime>();
        internal String xmlPath;

        public TaskSchedulerAssistantService()
        {
            if (debug) log("TaskSchedulerAssistantService");
            //start();
        }

        protected override void OnStart(String[] args)
        {
            if (debug) log("OnStart");
            start(args);
        }

        protected override void OnStop()
        {
            if (debug) log("OnStop");
            stop();
        }

        /// <summary>
        /// Generates the config xml
        /// </summary>
        internal bool createConfigXml()
        {
            if (debug) log("createConfigXml");
            log("createConfigXml: Generating new config file " + xmlPath);
            try
            {
                String str_path = xmlPath.Substring(0, xmlPath.LastIndexOf("\\"));
                if (!Directory.Exists(str_path)) Directory.CreateDirectory(str_path);
            }
            catch (Exception ex)
            {
                log("createConfigXml: Error, probably don't have permission to create the xml file in it's location." + xmlPath, ref ex);
                return false;
            }
            XmlAttribute xAttr;
            XmlDocument xDoc = new XmlDocument();
            XmlElement xERoot = xDoc.CreateElement(ServName);
            XmlNode xNTask, xNTrig;
            xDoc.AppendChild(xDoc.CreateXmlDeclaration("1.0", null, null));
            xDoc.AppendChild(xERoot);
            TaskService ts = new TaskService();
            foreach (Task t in ts.RootFolder.Tasks)
            {
                try
                {
                    xNTask = xDoc.CreateElement("task");
                    xAttr = xDoc.CreateAttribute("id");
                    xAttr.Value = t.Name;
                    xNTask.Attributes.Append(xAttr);
                    xNTrig = xDoc.CreateElement("trigger");
                    xAttr = xDoc.CreateAttribute("path");
                    xAttr.Value = " ";
                    xNTrig.Attributes.Append(xAttr);
                    xNTask.AppendChild(xNTrig);
                    xERoot.AppendChild(xNTask);
                    if (debug) log("createConfigXml: Adding " + t.Name);
                }
                catch (Exception ex)
                {
                    log("createConfigXml: Error creating the XML. " + xmlPath, ref ex);
                    return false;
                }
            }
            try
            {
                xDoc.Save(xmlPath);
            }
            catch (Exception ex)
            {
                log("createConfigXml: Error, could not save xml to path: " + xmlPath, ref ex);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Initializes the config XML fsw and timer
        /// </summary>
        internal bool initConfigXml()
        {
            if (debug) log("initConfigXml");
            //watch settings file for changes
            fswConfig = new FileSystemWatcher();
            fswConfig.InternalBufferSize = 32768;
            fswConfig.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            fswConfig.Changed += new FileSystemEventHandler(loadXmlFileSystemEventHandler);
            fswConfig.Created += new FileSystemEventHandler(loadXmlFileSystemEventHandler);
            fswConfig.Renamed += new RenamedEventHandler(loadXmlFileSystemEventHandler);
            try
            {
                fswConfig.Path = xmlPath.Substring(0, xmlPath.LastIndexOf("\\"));
                fswConfig.Filter = xmlPath.Substring(xmlPath.LastIndexOf("\\") + 1);
            }
            catch (Exception ex)
            {
                log("initConfigXml: Error setting fswConfig Path or Filter: " + fswConfig.Path + "/" + fswConfig.Filter, ref ex);
                return false;
            } try
            {
                fswConfig.EnableRaisingEvents = true;
                log("initConfigXml: Spawned watcher: " + fswConfig.Path + "\\" + fswConfig.Filter);
            }
            catch (Exception ex)
            {
                log("initConfigXml: Error spawning watcher: " + fswConfig.Path + "/" + fswConfig.Filter, ref ex);
                return false;
            }
            // Load the XML for the first time
            loadXml();
            return true;
        }

        /// <summary>
        /// Kills all the threads identified in a hashtable
        /// </summary>
        internal bool killFileSystemWatchers(Dictionary<FileSystemWatcher, String> threads)
        {
            if (debug) log("killFileSystemWatchers");
            if (threads.Count == 0) return true;
            foreach (KeyValuePair<FileSystemWatcher, String> fsw in threads)
            {
                log("killFileSystemWatchers: " + fsw.Key.Path);
                if (String.IsNullOrEmpty(fsw.Key.Path) || String.IsNullOrEmpty(fsw.Key.Filter)) return false;
                try
                {
                    fsw.Key.EnableRaisingEvents = false;
                    fsw.Key.Dispose();
                    log("killFileSystemWatchers: Killed watcher: " + fsw.Key.Path + "\\" + fsw.Key.Filter + " --> " + fsw.Value);
                }
                catch (Exception ex)
                {
                    log("killFileSystemWatchers: Error killing watcher: " + fsw.Key.Path + "\\" + fsw.Key.Filter, ref ex, EventLogEntryType.Error);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Reads in the config XML, sets the retry timer if it fails
        /// </summary>
        internal bool loadXml()
        {
            if (debug) log("loadXml");
            // Prevent doubling up
            if (!triggerAllow("loadXml")) return false;
            _xmlReload.Enabled = false;
            // Lock all the threads and load the config
            lock (_xmlReload)
            {
                XmlDocument xmlSettings = new XmlDocument();
                try
                {
                    FileStream xmlFile = new FileStream(xmlPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    xmlSettings.Load(xmlFile);
                    xmlFile.Close();
                }
                catch (Exception ex)
                {
                    log("loadXml: Error accessing " + xmlPath + ", re-attempting in " + (_xmlReload.Interval / 1000).ToString() + " seconds.", ref ex, EventLogEntryType.Error);
                    // Start the timer to automatically re-attempt
                    _xmlReload.Enabled = true;
                    return false;
                }
                // Separate out and initiate the triggers
                if (!startFileSystemWatchers(xmlSettings.SelectNodes("//" + ServName + "/task")))
                {
                    return false;
                };
            }
            if (debug) log("loadXml: Loaded " + xmlPath);
            return true;
        }

        /// <summary>
        /// The Timer trigger to retry loading the config xml
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The System.IO.ElapsedEventArgs that contains the event data.</param>
        private void loadXmlElapsedEventHandler(Object sender, ElapsedEventArgs e)
        {
            if (debug) log("loadXmlElapsedEventHandler");
            if (!loadXml())
            {
                log("loadXmlElapsedEventHandler: Error loading the config xml after 2 attempts, quitting.", EventLogEntryType.Error);
                stop();
            }
        }

        /// <summary>
        /// fsw trigger for the config XML
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The System.IO.FileSystemEventArgs that contains the event data.</param>
        private void loadXmlFileSystemEventHandler(Object sender, FileSystemEventArgs e)
        {
            if (debug) log("loadXmlFileSystemEventHandler");
            log("loadXmlFileSystemEventHandler: Reloading " + xmlPath);
            loadXml();
        }

        /// <summary>
        /// Simple logging to the Application Event Log
        /// </summary>
        /// <param name="msg">String message, no longer than 31,000 characters</param>
        private static void log(String msg)
        {
            log(msg, EventLogEntryType.Information);
        }

        /// <summary>
        /// Simple logging to the Application Event Log
        /// </summary>
        /// <param name="msg">String message, no longer than 31,000 characters</param>
        private static void log(String msg, EventLogEntryType type)
        {
            EventLog.WriteEntry(ServName, msg, type);
        }

        /// <summary>
        /// Simple logging to the Application Event Log
        /// </summary>
        /// <param name="msg">String message, no longer than 31,000 characters</param>
        public static void log(String msg, ref Exception ex)
        {
            log(msg, ref ex, EventLogEntryType.Information);
        }

        /// <summary>
        /// Simple logging to the Application Event Log
        /// </summary>
        /// <param name="msg">String message, no longer than 31,000 characters</param>
        private static void log(String msg, ref Exception ex, EventLogEntryType type)
        {
            msg = msg + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace;
            log(msg, type);
        }

        /// <summary>
        /// The common constructor for the testing program or the service
        /// </summary>
        internal bool start(String[] args)
        {
            if (debug) log("start");
            debug = Properties.Settings.Default.Debug;
            // Setup the log
            try
            {
                if (!EventLog.SourceExists(ServName)) EventLog.CreateEventSource(ServName, "Application");
            }
            catch (Exception ex)
            {
                log("start: Error while starting with " + xmlPath, ref ex, EventLogEntryType.Error);
                stop();
                return false;
            }
            // Handle the initial config xml
            xmlPath = Properties.Settings.Default.ConfigFile.Trim();
            if (String.IsNullOrEmpty(xmlPath))
            {
                xmlPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                log("The ConfigFile was not set in the " + Path.Combine(xmlPath, "TaskSchedulerAssistant.exe.config") 
                    + " file. Please change this to the file's permenant location. As default, TSA will be using " 
                    + Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Information Products Inc\\TaskSchedulerAssistant.xml"));
                xmlPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Information Products Inc\\TaskSchedulerAssistant.xml");
               
            }
            // Resolve any special path variables
            foreach (DictionaryEntry e in Environment.GetEnvironmentVariables())
            {
                xmlPath = xmlPath.Replace("%" + e.Key + "%", (String)e.Value);
            }
            // Generate the file if it doesn't exist
            if (!File.Exists(xmlPath)) createConfigXml();
            // Setup the timer
            _xmlReload.Elapsed += new ElapsedEventHandler(loadXmlElapsedEventHandler);
            // Start everything
            if (!initConfigXml())
            {
                stop();
                return false;
            }
            log("Service started using " + xmlPath);
            base.OnStart(args);
            return true;
        }

        /// <summary>
        /// Initializes a set of FileSystemWatchers
        /// </summary>
        /// <param name="nodes">The array of tasks and paths to monitor</param>
        internal bool startFileSystemWatchers(XmlNodeList xNTasks)
        {
            if (debug) log("startFileSystemWatchers");
            killFileSystemWatchers(fsws);
            fsws.Clear();
            FileSystemWatcher fsw;
            String path, name;
            foreach (XmlNode xNTask in xNTasks)
            {
                try
                {
                    name = xNTask.Attributes["id"].Value;
                    foreach (XmlNode xNTriggers in xNTask.SelectNodes("//" + ServName + "/task[@id='" + name + "']/trigger"))
                    {
                        // Delegate the triggers
                        path = xNTriggers.Attributes["path"].Value.Trim();
                        if (String.IsNullOrEmpty(path)) continue;
                        fsw = new FileSystemWatcher();
                        fsw.Filter = path.Substring(path.LastIndexOf("\\") + 1);
                        fsw.Path = path.Substring(0, path.LastIndexOf("\\"));
                        fsw.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.LastAccess | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                        fsw.Changed += new FileSystemEventHandler(triggerFileSystemEventHandler);
                        fsw.Created += new FileSystemEventHandler(triggerFileSystemEventHandler);
                        fsw.Renamed += new RenamedEventHandler(triggerFileSystemEventHandler);
                        fsw.InternalBufferSize = 32768;
                        fsw.EnableRaisingEvents = true;
                        fsws.Add(fsw, name);
                        log("startFileSystemWatchers: Spawned watcher: " + fsw.Path + "\\" + fsw.Filter + " --> " + name);
                    }
                }
                catch (Exception ex)
                {
                    log("startFileSystemWatchers: Error, could not start " + xNTask.OuterXml, ref ex, EventLogEntryType.Error);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// The common destructor for the testing program or the service
        /// </summary>
        internal bool stop()
        {
            if (debug) log("stop");
            bool ret = killFileSystemWatchers(fsws);
            try
            {
                fswConfig.EnableRaisingEvents = false;
                fswConfig.Dispose();
            }
            catch (Exception ex)
            {
                log("stop: Error disposing fswConfig.", ref ex, EventLogEntryType.Error);
                ret = false;
            }
            _xmlReload.Enabled = false;
            log("Service stopped");
            base.OnStop();
            return ret;
        }

        /// <summary>
        /// Checks if a task has been triggered too soon after the previous
        /// </summary>
        /// <returns>True/False on whether to continue</returns>
        internal bool triggerAllow(String name)
        {
            if (debug) log("triggerAllow: " + name);
            if (lastRun.Key == name)
            {
                if (DateTime.UtcNow.Subtract(lastRun.Value).Seconds < 5)
                {
                    if (debug) log("triggerAllow: Denied " + name);
                    return false;
                }
            }
            lastRun = new KeyValuePair<String, DateTime>(name, DateTime.UtcNow);
            if (debug) log("triggerAllow: Allowed " + name);
            return true;
        }

        /// <summary>
        /// Initiates a task triggered by an fsw
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The System.IO.FileSystemEventArgs that contains the event data.</param>
        internal void triggerFileSystemEventHandler(Object sender, FileSystemEventArgs e)
        {
            if (debug) log("triggerFs");
            triggerTask(fsws[(FileSystemWatcher)sender]);
        }

        /// <summary>
        /// Initiates a task thread based on name
        /// </summary>
        /// <param name="name">Name of the service as it appears in the xml</param>
        internal bool triggerTask(String name)
        {
            if (debug) log("triggerTask");
            // Prevent doubling up
            if (!triggerAllow(name)) return false;
            bool ran = false;
            TaskService ts = new TaskService();
            foreach (Task t in ts.RootFolder.Tasks)
            {
                try
                {
                    if (name != t.Name) continue;
                    log("Running " + name);
                    t.Run();
                    ran = true;
                }
                catch (Exception ex)
                {
                    log("triggerTask: Error", ref ex,EventLogEntryType.Error);
                    return false;
                }
            }
            if (!ran) log("Task " + name + " not found.");
            return ran;
        }

        private void InitializeComponent()
        {
            // 
            // TaskSchedulerAssistantService
            // 
            this.ServiceName = "TaskSchedulerAssistant";

        }

    }
}

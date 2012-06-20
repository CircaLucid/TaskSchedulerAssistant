namespace TaskSchedulerAssistant
{
    using NUnit.Framework;

    [TestFixture]
    public class TaskSchedulerAssistantServiceTest
    {
        private TaskSchedulerAssistantService ts;
        private System.String _xmlFile = "D:\\TaskSchedulerAssistant.xml";

        [SetUp]
        public void SetUp()
        {
            ts = new TaskSchedulerAssistantService();
            ts.debug = true;
            System.IO.File.Copy("D:\\IPI_TSA.xml", _xmlFile);
        }

        [Test]
        public void createConfigXml_bool()
        {
            ts.xmlPath = _xmlFile;
            Assert.IsTrue(ts.createConfigXml());
            ts.xmlPath = "C:\\Windows\\System32\\test.xml";
            Assert.IsFalse(ts.createConfigXml());
        }

        [Test]
        public void initConfigXml_bool()
        {
            bool temp;
            ts.xmlPath = _xmlFile;
            temp = ts.initConfigXml();
            if (!ts.fswConfig.Equals(null))
            {
                ts.fswConfig.EnableRaisingEvents = false;
                ts.fswConfig.Dispose();
            }
            Assert.IsTrue(temp);

            ts.xmlPath = "C:\\asdfasdf\\test.xml";
            temp = ts.initConfigXml();
            if (!ts.fswConfig.Equals(null))
            {
                ts.fswConfig.EnableRaisingEvents = false;
                ts.fswConfig.Dispose();
            }
            Assert.IsFalse(temp);
        }

        [Test]
        public void killFileSystemWatchers_bool()
        {
            System.IO.FileSystemWatcher fsw;
            fsw = new System.IO.FileSystemWatcher("D:\\", "*.bat");
            fsw.EnableRaisingEvents = true;
            ts.fsws.Add(fsw, "*.bat");
            fsw = new System.IO.FileSystemWatcher("D:\\", "*.txt");
            fsw.EnableRaisingEvents = true;
            ts.fsws.Add(fsw, "*.txt");
            Assert.IsTrue(ts.killFileSystemWatchers(ts.fsws));
            ts.fsws.Add(new System.IO.FileSystemWatcher(), "*.txt");
            Assert.IsFalse(ts.killFileSystemWatchers(ts.fsws));
        }

        [Test]
        public void loadXml_bool()
        {
            ts.xmlPath = _xmlFile;
            Assert.IsTrue(ts.loadXml());
            ts.xmlPath = "C:\\Windows\\System32\\test.xml";
            Assert.IsFalse(ts.loadXml());
        }

        [Test]
        public void start_bool_true()
        {
            ts.xmlPath = _xmlFile;
            System.String[] test = {};
            Assert.IsTrue(ts.start(test));
        }

        [Test]
        public void startFileSystemWatchers_bool_true()
        {
            System.Xml.XmlDocument xmlSettings = new System.Xml.XmlDocument();
            xmlSettings.LoadXml("<?xml version=\"1.0\"?><test><task id=\"auto_email\"><trigger path=\"D:\\~Tools\\auto_email\\*.email.bat\" /></task></test>");
            Assert.IsTrue(ts.startFileSystemWatchers(xmlSettings.SelectNodes("//test/task")));
        }

        [Test]
        public void stop_bool()
        {
            ts.fswConfig = new System.IO.FileSystemWatcher();
            Assert.IsTrue(ts.stop());
        }

        [Test]
        public void triggerAllow_bool()
        {
            ts.lastRun = new System.Collections.Generic.KeyValuePair<System.String, System.DateTime>("nightly upkeep", System.DateTime.UtcNow);
            // Should allow first run
            Assert.IsTrue(ts.triggerAllow("auto_email"));
            // Should error second run
            Assert.IsFalse(ts.triggerAllow("auto_email"));
            // Should allow not existant
            Assert.IsTrue(ts.triggerAllow("asdasdfa"));
        }

        // TODO: Complete this
        [Test]
        public void triggerFileSystemEventHandler_bool()
        {
        }

        [Test]
        public void triggerTask_bool()
        {
            Assert.IsTrue(ts.triggerTask("auto_email"));
            // Should error, does not exist
            Assert.IsFalse(ts.triggerTask("asdasdasfd"));
        }

        [TearDown]
        public void TearDown()
        {
            if (System.IO.File.Exists(_xmlFile)) System.IO.File.Delete(_xmlFile);
        }

    }
}
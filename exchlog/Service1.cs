using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace exchlog
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MySource", "MyNewLog");
            }
            eventLog1.Source = "MySource";
            eventLog1.Log = "MyNewLog";
        }

        protected override void OnStart(string[] args)
        {
            //eventLog1.WriteEntry("In OnStart");
            MainEx execlog = new MainEx();
            //StreamWriter wrLog = File.AppendText("d:\\temp\\startserv.log");
            execlog.wrLogU("OnStart: Service;");
            execlog.LogWatcher();
        }

        protected override void OnStop()
        {
            
        }
    }
}

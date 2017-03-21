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
            
        }

        protected override void OnStart(string[] args)
        {
            //eventLog1.WriteEntry("In OnStart");
            MainEx execlog = new MainEx();
            //StreamWriter wrLog = File.AppendText("d:\\temp\\startserv.log");
            execlog.wrLogU("Service: OnStart;");
            execlog.LogWatcher();
        }

        protected override void OnStop()
        {
            ///execlog.wrLogU("Service: OnStop;");
        }
    }
}

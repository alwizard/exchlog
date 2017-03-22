using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Security.Permissions;

using System.Threading;


namespace exchlog
{
    class MainEx
    {
        private String path = "C:\\Program Files\\Microsoft\\Exchange Server\\V15\\TransportRoles\\Logs\\FrontEnd\\ProtocolLog\\SmtpReceive\\";
        private Thread exec;
        private string pFile;
        private string fileLog;// = "d:\\temp\\exchsql.log";
        private IniFile getConfig;
        //public StreamWriter wrLog;

        public void wrLogU(string message)
        {
            using (StreamWriter wrLog = File.AppendText(fileLog))
            {
                DateTime time = DateTime.Now;

                wrLog.WriteLine(time.ToString("yyyy-MM-dd HH:mm:ss") + ": " + message);
                wrLog.Close();
            }
        }
        public MainEx()
        {
            //wrLog = File.AppendText(fLog);
            getConfig = new IniFile();
            fileLog = getConfig.getCurrPath() + getConfig.IniReadValue("SETTINGS", "LOGFILE");
            wrLogU("SETTINGS: Read:" + fileLog);
        }
        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        public void LogWatcher()
        {
            //try { 
            FileSystemWatcher fw = new FileSystemWatcher("C:\\Program Files\\Microsoft\\Exchange Server\\V15\\TransportRoles\\Logs\\FrontEnd\\ProtocolLog\\SmtpReceive");
            //fw.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
            //| NotifyFilters.FileName;
            fw.Filter = "*.LOG";

            fw.Changed += new FileSystemEventHandler(onChange);
            fw.EnableRaisingEvents = true;
            wrLogU("Enable: Enable FileSystemWatcher");
            //}catch(EventHandler e){

            //}
        }

        private void onChange(object source, FileSystemEventArgs e)
        {
            String nameF = e.Name;
            pFile = path + ParseDate(nameF);
            wrLogU("onChange: " + pFile);
            try
            {
                exec = new Thread(new ThreadStart(execExport));
                exec.Start();
                wrLogU("Thread: start.");
                //exec.ThreadState.ToString();

            }
            catch (Exception ex)
            {

                wrLogU("Exception Thread: " + ex.Message);
            }

        }

        public void execExport()
        {

            logInDB(pFile);

        }

        public string ParseDate(string fName)
        {
            //string[] dtstr = new string[FilesName.Count];
            string dtpattern = "yyyyMMddHH";
            string oldName = "";
            oldName = fName;
            //int i = 0;
            DateTime tmpDT;
            fName = fName.Substring(4, 10);
            tmpDT = DateTime.ParseExact(fName, dtpattern, null);
            DateTime dtFile = tmpDT.AddHours(-1);
            fName = "RECV" + dtFile.ToString("yyyyMMddHH", CultureInfo.InvariantCulture) + "-1.log";
            //FilesDate.Add(tmpDT);
            //ConvDateString.Add(FilesDate[i].ToString("u"));

            wrLogU(oldName + " => " + fName);
            //sw.Close();
            return fName;
        }
        void addRows(string[,] strrows)
        {
            System.Data.DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("date-time");
            dt.Columns.Add("connector-id");
            dt.Columns.Add("session-id");
            dt.Columns.Add("sequence-number");
            dt.Columns.Add("local-endpoint");
            dt.Columns.Add("remote-endpoint");
            dt.Columns.Add("event");
            dt.Columns.Add("data");
            dt.Columns.Add("context");
            DataRow rows;
            for (int i = 0; i < strrows.GetLength(0); i++)
            {
                rows = dt.NewRow();
                rows["id"] = "";
                rows["date-time"] = strrows[i, 0];
                rows["connector-id"] = strrows[i, 1];
                rows["session-id"] = strrows[i, 2];
                rows["sequence-number"] = strrows[i, 3];
                rows["local-endpoint"] = strrows[i, 4];
                rows["remote-endpoint"] = strrows[i, 5];
                rows["event"] = strrows[i, 6];
                rows["data"] = strrows[i, 7];
                rows["context"] = strrows[i, 8];
                //rows[9] = strrows[i, 9];
                dt.Rows.Add(rows);
            }
            
            string dbServer = getConfig.IniReadValue("SETTINGS", "DBSERVER");
            string dbName = getConfig.IniReadValue("SETTINGS", "DBNAME");
            string dbTable = getConfig.IniReadValue("SETTINGS", "DBTABLE");
            wrLogU("SETTINGS: ReadDBSettings => " + dbServer + ", " + dbName + " , " + dbTable);
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection("Data Source=" + dbServer + ";Initial Catalog=" + dbName + ";User ID=exchlog;Password=sRlJIl9beFx1;Persist Security Info=False");
                //Data Source=srv16-sql;Initial Catalog=MyDB;User ID=exchlog

            }
            catch (SqlException e)
            {

                wrLogU("ERROR: ConnectToSQL => " + e.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Open();
                    wrLogU("SQL Connection: open;");
                    using (var sqlBulk = new SqlBulkCopy(conn))
                    {
                        sqlBulk.DestinationTableName = dbTable;
                        sqlBulk.WriteToServer(dt);
                    }
                    conn.Close();
                    wrLogU("SQL:  Inserting " + dt.Rows.Count.ToString());
                    wrLogU("SQL Connection: closed;");
                }
                else
                {
                    wrLogU("ERROR: No connection to SQL;");
                }


            }

            //conn = new SqlConnection("Data Source="+dbServer+";Initial Catalog="+dbName+ ";User ID=exchlog;Password=sRlJIl9beFx1;Persist Security Info=False");
            //Data Source=srv16-sql;Initial Catalog=MyDB;User ID=exchlog


        }

        string parseStr(string str)
        {
            int fpos, lpos;
            fpos = str.IndexOf("\"");
            if (fpos == -1)
            {
                return str;
            }
            else
            {
                string pstr = str;
                lpos = pstr.LastIndexOf("\"");
                pstr = pstr.Substring(fpos, lpos - fpos);
                pstr = pstr.Replace(",", "");
                pstr = pstr.Replace("\"", "");
                str = str.Remove(fpos, lpos + 1 - fpos);
                str = str.Insert(fpos, pstr);
                return str;
            }

        }

        void logInDB(string fName)
        {
            //String fName = "";
            //String teststr = "2017-03-02T14:00:03.373Z,SRV03-EX\\Default Frontend SRV03-EX,08D451F1B0A301DC,1,149.56.250.66:25,156.67.106.242:64634,>,\"220 SRV03 - EX.verta.media Microsoft ESMTP MAIL Service ready at Thu, 2 Mar 2017 09:00:02 - 0500\",";
            //openFileD.ShowDialog();
            //fName = openFileD.FileName;
            String[] buffFile = System.IO.File.ReadAllLines(fName);
            //label1.Text = buffFile.Length.ToString();
            Array.Reverse(buffFile);
            Array.Resize(ref buffFile, buffFile.Length - 5);
            Array.Reverse(buffFile);


            string[,] str = new string[buffFile.Length, 9];
            int i = 0;

            //DataRow rows;


            foreach (string s in buffFile)
            {
                string[] tBuff = new string[9];
                string pstr = s;
                pstr = parseStr(pstr);
                string[] sBuff = pstr.Split(new char[1] { ',' }, StringSplitOptions.None);
                Array.Copy(sBuff, tBuff, 0);
                sBuff.CopyTo(tBuff, 0);
                int size = tBuff.Length;
                for (int j = 0; j < 9; j++)
                {
                    str[i, j] = tBuff[j];
                    if (j == 0)
                    {
                        string[] buffdt = str[i, j].Split('T');
                        str[i, j] = buffdt[0] + " " + buffdt[1].Remove(12, 1);

                    }
                }
                i++;
            }
            //MessageBox.Show(str.GetLength(0).ToString());
            //inSQL(str);
            addRows(str);
        }

    }
    // exchlog sRlJIl9beFx1
}

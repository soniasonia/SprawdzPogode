using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;
using System.Reflection;

namespace SprawdzPogode
{
    public class Control
    {
        private bool error;
        private string description;

        public Control()
        {
            error = false;
            description = "";
        }

        public bool Error
        {
            get; set;
        }
        public string Description
        {
            get; set;
        }

    }
  
    public static class MyMethods
    {
        
      


        
        public static string CheckStatus(Control con, string s)
        {
            if (con.Error == true)
            {
                Console.WriteLine(s + ". Action failed.\n" + con.Description);
                return s + ". Action failed.\n" + con.Description;
            }
            else
            {
                Console.WriteLine(s + ". Action successful.");
                return s + ". Action successful.";
            }

        }
        
}
    public class LogWriter
    {
        public delegate string StatusDel(Control c, string s);
        private string m_exePath = string.Empty;
        public event StatusDel CheckStatus;
        public LogWriter(string logMessage)
        {
            m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
            {
                w.WriteLine();
                w.WriteLine(logMessage);
            }
        }
        public void LogWrite(Control con, string action)
        {
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    string logMessage = action + " " + con.Error + ": " + con.Description;
                    logMessage = CheckStatus(con, action);
                    w.WriteLine(logMessage);
                }
            }
            catch (Exception e)
            {
            }
        }
        public void LogWrite(string logMessage)
        {
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    w.WriteLine(logMessage);
                }
            }
            catch (Exception e)
            {
            }
        }
    }
}




using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using System.Configuration;
using System.Data;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using System.IO;
using Lowis_Reports_Testing.ObjectLibrary;
using System.Management;
using System.Management.Instrumentation;



namespace Lowis_Reports_Testing
{

    public class CodedUIBase
    {
        string lowisclientbinlocation = ConfigurationManager.AppSettings["binpath"];
        string lowisserver = ConfigurationManager.AppSettings["servername"];
        string lowisusername = ConfigurationManager.AppSettings["username"];
        string lowispassword = ConfigurationManager.AppSettings["password"];
        protected ApplicationUnderTest TestApp;

        #region TEST_INITIALIZE
   
        public void LaunchLowisServer()
        {
            LowisConnectDialog lconndlg = new LowisConnectDialog();
            LowisSettingsDialog lsettings = new LowisSettingsDialog();
            LowisMainWindow lwin = new LowisMainWindow();
            try
            {
                TestApp = ApplicationUnderTest.Launch(lowisclientbinlocation);
                lconndlg.selectServer(lowisserver);
                if (lowisusername.Length == 0 && lowispassword.Length == 0)
                {
                    //use default readio button Simply connect
                    lconndlg.Settings.Click();
                    lsettings.storepath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString(), "csstore", DateTime.Now.ToString("ddMMMyyyyhhmmss"));
                    lsettings.btnSavesettings.Click();
                    lconndlg.Connect.Click();
                    lwin.Maximized = true;
                    lwin.Analysis.WaitForControlReady();
                    lwin.Analysis.Click();

                }
                else
                {
                    lconndlg.usecredentails.Click();
                    lconndlg.txtuserName.Text = lowisusername;
                    lconndlg.txtuserName.Text = lowispassword;
                    lconndlg.Settings.Click();
                    lsettings.storepath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString(), "csstore", DateTime.Now.ToString("ddMMMyyyyhhmmss"));
                    lsettings.btnSavesettings.Click();
                    lconndlg.Connect.Click();
                    lwin.Maximized = true;
                    lwin.Analysis.WaitForControlReady();
                    lwin.Analysis.Click();

                }
            }
            catch (Exception ex)
            {
                TestContext.WriteLine("Encountered Exception: " + ex.Message);
            }
        }
        #endregion

        #region TEST_CLEANUP
    
        public void closeLowisCleint()
        {

            TestApp.Close();
            TerminateProcessByForce("Lowis.exe");
            TerminateProcessByForce("LowisClient.exe");
        }
        #endregion

        #region UnitTestsDefaults
        private  void TerminateProcessByForce(string strprocess)
        {
            try
            {
                string processName = strprocess;
                ConnectionOptions connectoptions = new ConnectionOptions();
                string ipAddress = "127.0.0.1";
                ManagementScope scope = new ManagementScope(@"\\" + ipAddress + @"\root\cimv2", connectoptions);
                SelectQuery query = new SelectQuery("select * from Win32_process where name = '" + processName + "'");
                using (ManagementObjectSearcher searcher = new
                            ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {

                        process.InvokeMethod("Terminate", null);

                    }
                }

            }
            catch (Exception ex)
            {
                //Log exception in exception log.
                //Logger.WriteEntry(ex.StackTrace);
              //  Console.WriteLine(ex.StackTrace);
                throw new Exception(ex.Message);

            }
        }
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;


        #endregion
    }
}

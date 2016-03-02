﻿using System;
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
using Lowis_Reports_Testing.StructureSheet;



namespace Lowis_Reports_Testing
{
    /// <summary>
    /// Test For Verifying the Chart Viwer Titles, Legends for each of Link 
    /// </summary>
    [CodedUITest]
    public class ESPConfig :CodedUIBase
    {

        #region TEST_INITIALIZE
        //-----------------------------
        [TestInitialize]
        public void InitializeTest()
        {
            base.LaunchLowisServer();
        }
        //-----------------------------
        #endregion

        #region TEST_CLEANUP
        [TestCleanup]
        public void Cleanuptest()
        {
           // base.closeLowisCleint();
        }
        #endregion

        #region ESPConfigureChildssdata
        [TestMethod, Description(@"ESPConfigureChildssdata")]
        [DeploymentItem(@"..\TestData\ESPConfig")]

        public void espwellgrpconfig()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string srchWell1 = ConfigurationManager.AppSettings["testwell1"];
            try
            {

                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started"+TestContext.TestName+"  " + repeat);
                Lwindow.All.DoubleClick();
                Lwindow.AllWels.Click();
                Lwindow.WellTypes.DoubleClick();
                Lwindow.AllESPWells.Click();
                Lwindow.RefreshWells.Click();
                Lwindow.Start.WaitForControlReady();
                Lwindow.Start.Click();
                Lwindow.clickMenuitem(".Configuration", "ESP Well Group Configuration");
                Lwindow.espconfigureparameters.Click();
                //Select a Sepecific Well that can have good data to test these
              //  Lwindow.SelectWellfromSearch(srchWell1);
             //   DataTable dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "All");
                string wellnamesfile = ConfigurationManager.AppSettings["wellnamesfile"];
                UIObect ui = new UIObect();
              //  System.IO.File.Open(wellnamesfile, FileMode.Open);
                StreamReader fs = new StreamReader(wellnamesfile);
                string line = "";
                while ((line = fs.ReadLine()) != null)
                {
                    Lwindow.SelectWellfromSearch(line.Trim());
                    Playback.Wait(2000);
                    ui.AddData(System.IO.Directory.GetCurrentDirectory() + "\\ESP_Config_Params.xls", "TC_AEPOC_step_1_3");
                }

              //  Lwindow.SelectWellfromSearch(srchWell1);




                hr.LogtoTextFile(repeat + "Test execution Ended" + TestContext.TestName + "  " + repeat);

            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }

            




        }
        #endregion

        #region ESPWGC
        [TestMethod, Description(@"ESPWGC")]
        [DeploymentItem(@"..\TestData\ESPConfig")]
        [Timeout(TestTimeout.Infinite)]

        public void espwellgrpconfigAddWell()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string pcptcist = System.IO.Directory.GetCurrentDirectory() + "\\ESPTC.txt";
            try
            {

                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started" + repeat);
                Lwindow.All.DoubleClick();
                Lwindow.AllWels.Click();
                Lwindow.WellTypes.DoubleClick();
                Lwindow.AllESPWells.Click();
                Lwindow.RefreshWells.Click();
                Lwindow.Start.WaitForControlReady();
                Lwindow.Start.Click();
                Lwindow.clickMenuitem(".Configuration", "ESP Well Group Configuration");

                string wellnamesfile = ConfigurationManager.AppSettings["wellnamesfile"];
                UIObect ui = new UIObect();
                StreamReader fs = new StreamReader(pcptcist);
                string line = "";
                while ((line = fs.ReadLine()) != null)
                {
                    // Lwindow.SelectWellfromSearch(line.Trim());
                    Playback.Wait(2000);
                    ui.AddData(System.IO.Directory.GetCurrentDirectory() + "\\ESP_ConfigWell_Params.xls", line.Trim());
                }
                hr.LogtoTextFile(repeat + "Test execution Ended" + repeat);

            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }






        }
        #endregion
        [TestMethod, Description(@"ESPPumpConfig")]
        [DeploymentItem(@"..\TestData\ESPConfig")]
        public void esppumpconfig()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string srchWell1 = ConfigurationManager.AppSettings["testwell1"];
            try
            {

                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started" + TestContext.TestName + "  " + repeat);
                Lwindow.All.DoubleClick();
                Lwindow.AllWels.Click();
                Lwindow.WellTypes.DoubleClick();
                Lwindow.AllESPWells.Click();
                Lwindow.RefreshWells.Click();
                Lwindow.Start.WaitForControlReady();
                Lwindow.Start.Click();
                Lwindow.clickMenuitem(".Configuration", "ESP Well Group Configuration");
                Lwindow.esppumpconfig.Click();
                //Select a Sepecific Well that can have good data to test these
                //  Lwindow.SelectWellfromSearch(srchWell1);
                //   DataTable dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "All");
                string wellnamesfile = ConfigurationManager.AppSettings["wellnamesfile"];
                UIObect ui = new UIObect();
                //  System.IO.File.Open(wellnamesfile, FileMode.Open);
                StreamReader fs = new StreamReader(wellnamesfile);
                string line = "";
                while ((line = fs.ReadLine()) != null)
                {
                    Lwindow.SelectWellfromSearch(line.Trim());
                    Playback.Wait(2000);
                    ui.AddData(System.IO.Directory.GetCurrentDirectory() + "\\ESP_pump_Config.xls", "TC_1");
                }

                //  Lwindow.SelectWellfromSearch(srchWell1);




                hr.LogtoTextFile(repeat + "Test execution Ended" + TestContext.TestName + "  " + repeat);

            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }

            


        }

     


       

       
        

      
       
      
    }

   
}

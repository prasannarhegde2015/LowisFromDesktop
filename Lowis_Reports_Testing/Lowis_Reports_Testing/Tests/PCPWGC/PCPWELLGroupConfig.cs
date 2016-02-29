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
    public class PCPConfig :CodedUIBase
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

        #region PCPWGC
        [TestMethod, Description(@"PCPWGC")]
        [DeploymentItem(@"..\TestData\PCPConfig")]
        [Timeout(TestTimeout.Infinite)]

        public void pcpwellgrpconfig()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string pcptcist = System.IO.Directory.GetCurrentDirectory() + "\\PCPTC.txt";
            try
            {

                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started" + repeat);
                Lwindow.All.DoubleClick();
                Lwindow.AllWels.Click();
                Lwindow.WellTypes.DoubleClick();
                Lwindow.AllPCPWells.Click();
                Lwindow.RefreshWells.Click();
                Lwindow.Start.WaitForControlReady();
                Lwindow.Start.Click();
                Lwindow.clickMenuitem(".Configuration", "PCP Well Group Configuration");
               
                string wellnamesfile = ConfigurationManager.AppSettings["wellnamesfile"];
                UIObect ui = new UIObect();
                StreamReader fs = new StreamReader(pcptcist);
                string line = "";
                while ((line = fs.ReadLine()) != null)
                {
                   // Lwindow.SelectWellfromSearch(line.Trim());
                    Playback.Wait(2000);
                    ui.AddData(System.IO.Directory.GetCurrentDirectory() + "\\PCP_Config_Params.xls", line.Trim());
                }
                hr.LogtoTextFile(repeat + "Test execution Ended" + repeat);

            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }

            




        }
        #endregion

     //   [TestMethod]
     //   public void getallhtmllinks()
      /*  {
            LReportPane lp2 = new LReportPane();
            Helper hp1 = new Helper();
            hp1.LogtoTextFile(lp2.getallhtmllinks());
        } */

        #region LocalTestMethods
        public void Chartclick(DataTable dt, LowisMainWindow lw1, LReportPane lp1)
        {
            foreach (DataRow dr in dt.Rows)
            {
                string tcname = dr["TestCaseID"].ToString();
                string linkname = dr["Name"].ToString();
                string chartindex= dr["ChartIndex"].ToString();
                string charttitle = dr["ChartTitle"].ToString();
                string chartytitle = dr["ChartYTitle"].ToString();
                string chartxtitle = dr["ChartXTilte"].ToString();
                string charty2title = dr["ChartY2Title"].ToString();
                string chartlegends = dr["ChartLegends"].ToString();
                // Verify
                lw1.lowisDwait();
                lp1.ClickHtmlLink(linkname);
                lw1.lowisDwait();
                lp1.VerifyChartPage(tcname, linkname, chartindex, charttitle, chartytitle, charty2title, chartxtitle, chartlegends);
            }
        }


        #endregion
       

        
        

      

       
      
    }

   
}

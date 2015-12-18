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


namespace Lowis_Reports_Testing
{
    /// <summary>
    /// Test For Verifying the Chart Viwer Titles, Legends for each of Link 
    /// </summary>
    [CodedUITest]
    public class BeamCharts :  CodedUIBase
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
           base.closeLowisCleint();
        }
        #endregion

        #region BeamChartsViewer
        [TestMethod, Description(@"Beam Charts Verification")]
        [DeploymentItem(@"..\TestData\BeamCharts")]

        public void VerifyBeamChartsLink()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string srchWell1 = ConfigurationManager.AppSettings["testwell1"];
            try
            {

                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started" + repeat);
                Lwindow.All.DoubleClick();
                Lwindow.AllWels.Click();
                Lwindow.WellTypes.DoubleClick();
                Lwindow.AllBeamWells.Click();
                Lwindow.RefreshWells.Click();
                Lwindow.Start.WaitForControlReady();
                Lwindow.Start.Click();
                Lwindow.clickMenuitem(".Charts", "Beam Chart Viewer");
                //Select a Sepecific Well that can have good data to test these
                Lwindow.SelectWellfromSearch(srchWell1);
                DataTable dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "All");
                Chartclick(dt, Lwindow, lpnae);
                dt.Clear();

                //dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "Welltest");
                //lpnae.tabWelltest.Click();
                //Chartclick(dt, Lwindow, lpnae);
                //dt.Clear();

                //dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "Analysis Charts");
                //lpnae.tabAlarms.Click();
                //Chartclick(dt, Lwindow, lpnae);
                //dt.Clear();

                //dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "WSM Charts");
                //lpnae.tabParameters.Click();
                //Chartclick(dt, Lwindow, lpnae);
                //dt.Clear();

                //dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "Charts By Controller");
                //lpnae.tabConfiguration.Click();
                //Chartclick(dt, Lwindow, lpnae);

                hr.LogtoTextFile(repeat + "Test execution Ended" + repeat);
            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }

            




        }
        #endregion

    
     

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
                Helper hlpp = new Helper();
                hlpp.LogtoTextFile(" result of licnk check " +lp1.Linkexists);
                if (lp1.Linkexists == "true")
                {
                    lw1.lowisDwait();
                    lp1.VerifyChartPage(tcname, linkname, chartindex, charttitle, chartytitle, charty2title, chartxtitle, chartlegends);
                }
            }
        }

        public void getallhtmllinks()
        {
            LReportPane lp2 = new LReportPane();
            Helper hp1 = new Helper();
            hp1.LogtoTextFile(lp2.getallhtmllinks());
        } 
        #endregion

       
      
    }

   
}

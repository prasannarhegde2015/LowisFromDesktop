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
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class BeamReports : CodedUIBase
    {

        #region TEST_INITIALIZE
        [TestInitialize]
       
        public void InitializeTest()
        {
            base.LaunchLowisServer();
        }
        #endregion

        #region TEST_CLEANUP
        [TestCleanup]
        public void Cleanuptest()
        {
            base.closeLowisCleint();
        }
        #endregion


        #region BeamReportsTest
        [TestMethod, Description(@"Beam Reports Links Verification")]
        [DeploymentItem(@"..\TestData\BeamReports")]
        public void VerifyBeamReportsLink()
        {

            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string repeat = new string('=', 50);
            hr.LogtoTextFile(repeat + "Test execution Started" + repeat);
            Lwindow.All.DoubleClick();
            Lwindow.AllWels.Click();
            Lwindow.WellTypes.DoubleClick();
            Lwindow.AllBeamWells.Click();
            Lwindow.RefreshWells.Click();
            Lwindow.Start.WaitForControlReady();
            Lwindow.Start.Click();
            Lwindow.clickMenuitem(".Reports");
            Lwindow.clickMenuitem("Beam Reports");
            DataTable dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "All");
            this.Reportclick(dt, Lwindow, lpnae);
            dt.Clear();
            dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "Status");
            lpnae.tabStatus.Click();
            Reportclick(dt, Lwindow, lpnae);
            dt.Clear();
            dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "Analysis");
            lpnae.tabAnalysis.Click();
            Reportclick(dt, Lwindow, lpnae);
            dt.Clear();
            dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "Welltest");
            lpnae.tabWelltest.Click();
            Reportclick2(dt, Lwindow, lpnae);
            dt.Clear();
            lpnae.btnLess.Click();
            dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "Alarms");
            lpnae.tabAlarms.Click();
            for (int i = 0; i < 3; i++)
            {
                lpnae.btnLess.Click();
            }
            Reportclick2(dt, Lwindow, lpnae);
            dt.Clear();
            dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "Parameter");
            lpnae.tabParameters.Click();
            Reportclick(dt, Lwindow, lpnae);
            dt.Clear();
            dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamReportsLinksName.xls", "Sheet1", "ReportTabPage", "Configuration");
            lpnae.tabConfiguration.Click();
            Reportclick(dt, Lwindow, lpnae);

            hr.LogtoTextFile(repeat + "Test execution Ended" + repeat);

        }
        #endregion

        #region LocalHelpMethods
        public void Reportclick(DataTable dt, LowisMainWindow lw1, LReportPane lp1)
        {
            foreach (DataRow dr in dt.Rows)
            {
                string tcname = dr["TestCaseID"].ToString();
                string linkname = dr["Name"].ToString();
                string reptype = dr["Type"].ToString();
                string tindex = dr["TitleHTMLTableIndex"].ToString();
                string ttext = dr["TitleText"].ToString();
                string colindex = dr["ColumnNameHTMLTableIndex"].ToString();
                string colnametext = dr["ColumnNameText"].ToString();
                string tableType = dr["TableType"].ToString();
                // Verify
                lw1.lowisDwait();
                lp1.ClickHtmlLink(linkname);
                lw1.lowisDwait();
                lp1.VerifyReportPage(tcname, linkname, tableType, reptype, tindex, ttext, colindex, colnametext);
            }
        }
        public void Reportclick2(DataTable dt, LowisMainWindow lw2, LReportPane lp2)
        {
            foreach (DataRow dr in dt.Rows)
            {
                string tcname = dr["TestCaseID"].ToString();
                string linkname = dr["Name"].ToString();
                string reptype = dr["Type"].ToString();
                string tindex = dr["TitleHTMLTableIndex"].ToString();
                string ttext = dr["TitleText"].ToString();
                string colindex = dr["ColumnNameHTMLTableIndex"].ToString();
                string colnametext = dr["ColumnNameText"].ToString();
                string tableType = dr["TableType"].ToString();
                // Verify
                lw2.lowisDwait();
                lp2.ClickHtmlLink(linkname);
                if (linkname == "Auto Welltest Evaluation Report")
                {
                    if (lp2.btnOKAutoReport.Exists)
                    {
                        lp2.btnOKAutoReport.Click();
                    }
                }
                else if (linkname == "Most Recent Coded Test Rpt (Wells)" || linkname == "Most Recent Coded Tests Rpt (Facility)")
                {
                    if (lp2.btnOKCodedReport.Exists)
                    {
                        lp2.btnOKCodedReport.Click();
                    }

                }

                lw2.lowisDwait();
                lp2.VerifyReportPage(tcname, linkname, tableType, reptype, tindex, ttext, colindex, colnametext);
            }
        }
        #endregion




    }
}

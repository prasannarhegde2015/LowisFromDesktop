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
using Lowis_Reports_Testing.StructureSheet;


namespace Lowis_Reports_Testing
{
    /// <summary>
    /// Test For Verifying the Chart Viwer Titles, Legends for each of Link 
    /// </summary>
    [CodedUITest]
    public class WellFloTest :CodedUIBase
    {

        #region TEST_INITIALIZE
        //-----------------------------
        [TestInitialize]
        public void InitializeTest()
        {
           // base.LaunchLowisServer();
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

        #region BeamChartsViewer
        [TestMethod, Description(@"Beam Charts Verification")]
        [DeploymentItem(@"..\TestData\ESPConfig")]

        public void TestPOC()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            UIObect ui = new UIObect();
            string srchWell1 = ConfigurationManager.AppSettings["testwell1"];
            try
            {
                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started" + repeat);
                ui.AddData(System.IO.Directory.GetCurrentDirectory() + "\\Wft1.xls", "T1");
                hr.LogtoTextFile(repeat + "Test execution Ended" + repeat);

            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }

            




        }
        #endregion

    
    

       
      
    }

   
}

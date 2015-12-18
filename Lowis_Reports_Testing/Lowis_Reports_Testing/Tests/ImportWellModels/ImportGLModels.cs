using System;
using System.Linq;
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
using System.Windows.Automation;



namespace Lowis_Reports_Testing
{
    /// <summary>
    /// Test For Importing wEll models
    /// </summary>
    [CodedUITest]
    public class ImportModelFiles : CodedUIBase
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

        #region WellModelImport
        [TestMethod, Description(@"Import Well Models")]
        [DeploymentItem(@"..\TestData\GLImport")]

        public void ImportWellmodelFiles()
        {
            LowisMainWindow Lwindow = new LowisMainWindow();
            LReportPane lpnae = new LReportPane();
            Helper hr = new Helper();
            string srchWell1 = ConfigurationManager.AppSettings["testwell1"];
            string wellmodelsfolderpath = ConfigurationManager.AppSettings["wellmodelsfolderpath"];
            try
            {

                string repeat = new string('=', 50);
                hr.LogtoTextFile(repeat + "Test execution Started" + repeat);
                //Lwindow.All.DoubleClick();
                //Lwindow.AllWels.Click();
                //Lwindow.WellTypes.DoubleClick();
                //Lwindow.AllESPWells.Click();
                //Lwindow.RefreshWells.Click();
                //Lwindow.Start.WaitForControlReady();
                //Lwindow.Start.Click();
                //Lwindow.clickMenuitem(".Configuration", "ESP Well Group Configuration");
                //Select a Sepecific Well that can have good data to test these
                //  Lwindow.SelectWellfromSearch(srchWell1);
                //   DataTable dt = hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\BeamChartsLinksName.xls", "Sheet1", "ReportTabPage", "All");
                UIObect ui = new UIObect();
                 DataTable dtwll =hr.dtFromExcelFile(System.IO.Directory.GetCurrentDirectory() + "\\GL_mapModelmap.xls", "Sheet1");

                  foreach(DataRow dr in dtwll.Rows)
                  {
                      string wellmodlename = dr["ModelFileName"].ToString();
                      int posuscore = wellmodlename.IndexOf("_");
                      string prewellnumber = wellmodlename.Substring(0, posuscore);
                      string intwellnumber = prewellnumber.Replace("GL", "");
                      DataRow[] result = dtwll.Select("WellNumber='" + intwellnumber + "'");
                      string mapwell = result[0]["WellName"].ToString();
                      Lwindow.SelectWellfromSearch(mapwell);
                      wellmodlename = Path.Combine(wellmodelsfolderpath, dr["ModelFileName"].ToString());
                      hr.LogtoTextFile("Model path updated " + wellmodlename);
                      hr.UpdateExcelFileColumn(System.IO.Directory.GetCurrentDirectory() + "\\GL_import.xls", "ExpectedData", "BrowseButton",
                      wellmodlename, "TestCase", "TC_import");
                      Playback.Wait(2000);
                      ui.AddData(System.IO.Directory.GetCurrentDirectory() + "\\GL_import.xls", "TC_import");
                  }

                   
               
                hr.LogtoTextFile(repeat + "Test execution Ended" + repeat);

            }
            catch (Exception ex)
            {
                hr.LogtoTextFile("Exeption occured : " + ex.Message.ToString());

            }






        }
        #endregion

        #region GetWellNamesList

        [TestMethod]
        public void getWellNamesList()
        {
            AutomationElement rootelem = AutomationElement.RootElement;

            Condition cnddatagridscollection = new System.Windows.Automation.PropertyCondition(
                AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataGrid);
            AutomationElementCollection  navgrids = rootelem.FindAll(TreeScope.Descendants, cnddatagridscollection);
            AutomationElement welldtgrid = navgrids[3];
            TablePattern tbl = (TablePattern)welldtgrid.GetCurrentPattern(TablePattern.Pattern);
            TestContext.WriteLine("Count of wells = " + tbl.Current.RowCount);
            int matchedColnumber = 0;
            string matchcolumnNmae = "Well Name";
            Condition cndheaderItems = new System.Windows.Automation.PropertyCondition(
                AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.HeaderItem);

            AutomationElementCollection headeritemcolection = welldtgrid.FindAll(TreeScope.Descendants, cndheaderItems);
            foreach (AutomationElement headerelem in headeritemcolection)
            {
                if (headerelem.Current.Name.ToLower() == matchcolumnNmae.ToLower())
                {
                    break;
                }
                matchedColnumber++;
            }
            TestContext.WriteLine("Match column number " + matchedColnumber);
            Condition cndListItems = new System.Windows.Automation.PropertyCondition(
                AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem);
             Condition cndtxtItems = new System.Windows.Automation.PropertyCondition(
                AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text);

            AutomationElementCollection listitems =  welldtgrid.FindAll(TreeScope.Descendants, cndListItems);
            int wllcount =1 ;
            foreach (AutomationElement listelem in listitems)
            {
                AutomationElementCollection cellitemsinrow = listelem.FindAll(TreeScope.Children, cndtxtItems);

                TestContext.WriteLine("(" + wllcount + ") WellName : " + cellitemsinrow[matchedColnumber].Current.Name);
                Helper hpp = new Helper();
                hpp.LogtoTextFile(cellitemsinrow[matchedColnumber].Current.Name);
                    wllcount++;
            }
        }

        [TestMethod]
        public void listwellmodelsinfolder()
        {
            string srcdir = ConfigurationManager.AppSettings["srcdir"];
            DirectoryInfo dir = new DirectoryInfo(srcdir);
            List<String> arrfiles = (from f in dir.GetFiles()
                                     orderby f.Name
                                     select f.Name).ToList();


            foreach (string innm in arrfiles)
            {
                File.AppendAllText(Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "list.txt"), innm + Environment.NewLine);
            }
        }

        #endregion






    }


}

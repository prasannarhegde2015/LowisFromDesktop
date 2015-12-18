using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CUITe.Controls.WinControls;
using CUITe.Controls.HtmlControls;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CUITe;
using CUITe.Controls.WpfControls;


namespace Lowis_Reports_Testing.ObjectLibrary
{
   public  class LReportPane : CUITe_WinWindow
    {
        Helper hlp = new Helper();
      //  LReportPane() :  base("ControlName=BrowserMain") { }
        public enum CompareType
        {
            Equal,
            Contains,
            Tolerance,
            Decimalround

        }
        LowisMainWindow lmain = new LowisMainWindow();

        #region ObjectStructure
        #region HtmlReportControls
        public CUITe_HtmlTable reptable
        {
            get
            {
                return lmain.Get<CUITe_WinClient>("ClassName=Internet Explorer_Server").Get<CUITe_HtmlDocument>("Id="+null).Get<CUITe_HtmlTable>("Id=eTbl;TagName=Table"); 
            }
        }

        public CUITe_HtmlDocument reportLinksPane
        {
            get
            {
                return lmain.Get<CUITe_WinClient>("ClassName=Internet Explorer_Server;Instance=2").Get<CUITe_HtmlDocument>("Id=" + null);
            }
        }
        public CUITe_HtmlDocument reportPane
        {
            get
            {
                return lmain.Get<CUITe_WinClient>("ClassName=Internet Explorer_Server").Get<CUITe_HtmlDocument>("Id=" + null);
            }
        }


        public CUITe_HtmlTable webtableatIndex(string index)
        {
            return lmain.Get<CUITe_WinClient>("ClassName=Internet Explorer_Server").Get<CUITe_HtmlDocument>("Id=" + null).Get<CUITe_HtmlTable>("TagName=Table;TagInstance=" + index); 
        }

        public CUITe_HtmlCell webtableatcellIndex(string index)
        {
            return lmain.Get<CUITe_WinClient>("ClassName=Internet Explorer_Server").Get<CUITe_HtmlDocument>("Id=" + null).Get < CUITe_HtmlCell>("TagName=TD;TagInstance=" + index); 
        }
        #endregion

        #region Winwindows
        public CUITe_WinWindow txtReportPane
        {
            get
            {
                return Get<CUITe_WinWindow>("ControlId=998");
            }
       }
        public CUITe_WinWindow wndUIOKButton1 { get { return Get<CUITe_WinWindow>("Name=OK;ClassName=Button"); } }
        public CUITe_WinWindow wndEnterwellTest { get { return Get<CUITe_WinWindow>("Name=Enter the number of Tests to display"); } }
        public CUITe_WinWindow wndUIOKButton2 { get { return wndEnterwellTest.Get<CUITe_WinWindow>("Name=OK;ClassName=Button"); } }
        public CUITe_WinWindow wndUIOLess { get { return  Get<CUITe_WinWindow>("ControlId=1;Instance=3"); } }
        public CUITe_WinWindow wndTabWindow { get { return Get<CUITe_WinWindow>("ControlId=242"); } }
        public CUITe_WinWindow WndGeorge { get { return lmain.Get<CUITe_WinWindow>("Name=George;ClassName=HwndWrapper"); } }
        #endregion

        #region D3ChartWindows
        public CUITe_WpfWindow wndEditAxisSettings { get { return Get<CUITe_WpfWindow>("Name~Axis Settings"); } }
        public CUITe_WpfWindow wndChartGeneralSettings { get { return Get<CUITe_WpfWindow>("Name~General Settings"); } }
#endregion

        #region Winbuttons
        public CUITe_WinButton btnOKAutoReport { get { return wndUIOKButton1.Get<CUITe_WinButton>("Name=OK"); } }
        public CUITe_WinButton btnOKCodedReport { get { return wndUIOKButton2.Get<CUITe_WinButton>("Name=OK"); } }
        public CUITe_WinButton btnLess { get { return wndUIOLess.Get<CUITe_WinButton>("Name=Less"); } }
        #endregion

        public CUITe_WinTabList tabList { get { return wndTabWindow.Get<CUITe_WinTabList>("Name=Tab1"); } }

        public CUITe_WinEdit reportEditor
        {
            get
            {
                return this.txtReportPane.Get<CUITe_WinEdit>("Instance=1");
            }
        }

        #region WinTabPages
        public CUITe_WinTabPage tabStatus { get { return tabList.Get<CUITe_WinTabPage>("Name=Status"); } }
        public CUITe_WinTabPage tabAnalysis { get { return tabList.Get<CUITe_WinTabPage>("Name=Analysis"); } }
        public CUITe_WinTabPage tabWelltest { get { return tabList.Get<CUITe_WinTabPage>("Name=Well Test"); } }
        public CUITe_WinTabPage tabAlarms { get { return tabList.Get<CUITe_WinTabPage>("Name=Alarms"); } }
        public CUITe_WinTabPage tabParameters { get { return tabList.Get<CUITe_WinTabPage>("Name=Parameter"); } }
        public CUITe_WinTabPage tabConfiguration { get { return tabList.Get<CUITe_WinTabPage>("Name=Configuration"); } }
        #endregion

        #region D3CHARTS
       
        public CUITe_WpfCustom custConainerParent1 { get { return WndGeorge.Get<CUITe_WpfCustom>("Instance=1"); } }

        public CUITe_WpfCustom custChartContainer1 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.Chart;Instance=1"); } }
        public CUITe_WpfCustom custChartContainer2 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.Chart;Instance=2"); } }
        public CUITe_WpfCustom custChartContainer3 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.Chart;Instance=3"); } }

        public CUITe_WpfCustom custSplitCenterContainer1 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.SplitPane;AutomationId=splitCenter;Instance=1"); } }
        public CUITe_WpfCustom custSplitCenterContainer2 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.SplitPane;AutomationId=splitCenter;Instance=2"); } }
        public CUITe_WpfCustom custSplitCenterContainer3 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.SplitPane;AutomationId=splitCenter;Instance=3"); } }

        public CUITe_WpfCustom custSplitRightContianer1 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.SplitPane;AutomationId=splitRight;Instance=1") ;}}
        public CUITe_WpfCustom custSplitRightContianer2 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.SplitPane;AutomationId=splitRight;Instance=2") ;}}
        public CUITe_WpfCustom custSplitRightContianer3 { get { return custConainerParent1.Get<CUITe_WpfCustom>("ClassName=Uia.SplitPane;AutomationId=splitRight;Instance=3"); } }

        public CUITe_WpfCustom custPlotter1 { get { return custChartContainer1.Get<CUITe_WpfCustom>("ClassName=Uia.Plotter;AutomationId=Plotter"); } }
        public CUITe_WpfCustom custPlotter2 { get { return custChartContainer2.Get<CUITe_WpfCustom>("ClassName=Uia.Plotter;AutomationId=Plotter"); } }
        public CUITe_WpfCustom custPlotter3 { get { return custChartContainer3.Get<CUITe_WpfCustom>("ClassName=Uia.Plotter;AutomationId=Plotter"); } }

        public CUITe_WpfCustom custcpChartContainer1 { get { return custSplitCenterContainer1.Get<CUITe_WpfCustom>("ClassName=Uia.ContentPane;AutomationId=cpChart"); } }
        public CUITe_WpfCustom custcpChartContainer2 { get { return custSplitCenterContainer2.Get<CUITe_WpfCustom>("ClassName=Uia.ContentPane;AutomationId=cpChart"); } }
        public CUITe_WpfCustom custcpChartContainer3 { get { return custSplitCenterContainer3.Get<CUITe_WpfCustom>("ClassName=Uia.ContentPane;AutomationId=cpChart"); } }


        public CUITe_WpfToolBar chtoolbar1 { get { return custSplitCenterContainer1.Get<CUITe_WpfToolBar>("Instance=1"); } }
        public CUITe_WpfToolBar chtoolbar2 { get { return custSplitCenterContainer2.Get<CUITe_WpfToolBar>("Instance=1"); } }
        public CUITe_WpfToolBar chtoolbar3 { get { return custSplitCenterContainer3.Get<CUITe_WpfToolBar>("Instance=1"); } }
       

        public CUITe_WpfCustom custlegend1 { get { return tablegpage1.Get<CUITe_WpfCustom>("ClassName=Uia.ContentPane;AutomationId=cpLegend"); } }
        public CUITe_WpfCustom custlegend2 { get { return tablegpage2.Get<CUITe_WpfCustom>("ClassName=Uia.ContentPane;AutomationId=cpLegend"); } }
        public CUITe_WpfCustom custlegend3 { get { return tablegpage3.Get<CUITe_WpfCustom>("ClassName=Uia.ContentPane;AutomationId=cpLegend"); } }

        public CUITe_WpfButton btnEditChartaxis1 { get { return custcpChartContainer1.Get<CUITe_WpfButton>("HelpText=Edit Axis Settings"); } }
        public CUITe_WpfButton btnEditChartaxis2 { get { return custcpChartContainer2.Get<CUITe_WpfButton>("HelpText=Edit Axis Settings"); } }
        public CUITe_WpfButton btnEditChartaxis3 { get { return custcpChartContainer3.Get<CUITe_WpfButton>("HelpText=Edit Axis Settings"); } }

        public CUITe_WpfButton btnEditGeneralSettings1 { get { return custcpChartContainer1.Get<CUITe_WpfButton>("HelpText=Edit General Chart Settings"); } }
        public CUITe_WpfButton btnEditGeneralSettings2 { get { return custcpChartContainer2.Get<CUITe_WpfButton>("HelpText=Edit General Chart Settings"); } }
        public CUITe_WpfButton btnEditGeneralSettings3 { get { return custcpChartContainer3.Get<CUITe_WpfButton>("HelpText=Edit General Chart Settings"); } }

        public CUITe_WpfToggleButton btnShowLegends1 { get { return custcpChartContainer1.Get<CUITe_WpfToggleButton>("HelpText=Show Legend"); } }
        public CUITe_WpfToggleButton btnShowLegends2 { get { return custcpChartContainer2.Get<CUITe_WpfToggleButton>("HelpText=Show Legend"); } }
        public CUITe_WpfToggleButton btnShowLegends3 { get { return custcpChartContainer3.Get<CUITe_WpfToggleButton>("HelpText=Show Legend"); } }

        public CUITe_WpfToggleButton btnHideLegends1 { get { return custcpChartContainer1.Get<CUITe_WpfToggleButton>("HelpText=Hide Legend"); } }
        public CUITe_WpfToggleButton btnHideLegends2 { get { return custcpChartContainer2.Get<CUITe_WpfToggleButton>("HelpText=Hide Legend"); } }
        public CUITe_WpfToggleButton btnHideLegends3 { get { return custcpChartContainer3.Get<CUITe_WpfToggleButton>("HelpText=Hide Legend"); } }

        public CUITe_WpfToggleButton btnchartOverflowBtn1 { get { return custcpChartContainer1.Get<CUITe_WpfToggleButton>("AutomationId=OverflowButton"); } }
        public CUITe_WpfToggleButton btnchartOverflowBtn2 { get { return custcpChartContainer2.Get<CUITe_WpfToggleButton>("AutomationId=OverflowButton"); } }
        public CUITe_WpfToggleButton btnchartOverflowBtn3 { get { return custcpChartContainer3.Get<CUITe_WpfToggleButton>("AutomationId=OverflowButton"); } }


        public CUITe_WpfButton btnCancelEditAxis { get { return grpAxes.Get<CUITe_WpfButton>("AutomationId=btnCancel"); } }
        public CUITe_WpfButton btnCancelGenChartAxis { get { return wndChartGeneralSettings.Get<CUITe_WpfButton>("AutomationId=btnCancel"); } }

        public CUITe_WpfGroup grpAxes { get { return wndEditAxisSettings.Get<CUITe_WpfGroup>("Name=Axes"); } }
        public CUITe_WpfGroup grpAxistitle { get { return grpAxes.Get<CUITe_WpfGroup>("Name=Axis Title"); } }
        public CUITe_WpfGroup grpChartHeader { get { return wndChartGeneralSettings.Get<CUITe_WpfGroup>("Name=Chart Header"); } }


        public CUITe_WpfComboBox cmbAxis { get { return grpAxes.Get<CUITe_WpfComboBox>("Automationid=cmbAxis"); } }


        public CUITe_WpfList lstLegBox1 { get { return custlegend1.Get<CUITe_WpfList>("AutomationId=legendList"); } }
        public CUITe_WpfList lstLegBox2 { get { return custlegend2.Get<CUITe_WpfList>("AutomationId=legendList"); } }
        public CUITe_WpfList lstLegBox3 { get { return custlegend3.Get<CUITe_WpfList>("AutomationId=legendList"); } }

       public CUITe_WpfTabList tablistLeg1 {get { return custSplitRightContianer1.Get<CUITe_WpfTabList>("AutomationId=tbGroupRight");}}
       public CUITe_WpfTabList tablistLeg2 {get { return custSplitRightContianer2.Get<CUITe_WpfTabList>("AutomationId=tbGroupRight");}}
       public CUITe_WpfTabList tablistLeg3 { get { return custSplitRightContianer3.Get<CUITe_WpfTabList>("AutomationId=tbGroupRight"); } }

       public CUITe_WpfTabPage tablegpage1 {get { return  tablistLeg1.Get<CUITe_WpfTabPage>("Name=Legend");}}
       public CUITe_WpfTabPage tablegpage2 {get { return  tablistLeg2.Get<CUITe_WpfTabPage>("Name=Legend");}}
       public CUITe_WpfTabPage tablegpage3 { get { return tablistLeg3.Get<CUITe_WpfTabPage>("Name=Legend"); } }




        public CUITe_WpfEdit tbaxistitle { get { return grpAxistitle.Get<CUITe_WpfEdit>("Automationid=txtAxisTitle"); } }
        public CUITe_WpfEdit tbchartTitle { get { return grpChartHeader.Get<CUITe_WpfEdit>("Automationid=txtHeader"); } } 


        #endregion


        #endregion


        #region ObjectMethods
        public string  getallhtmllinks()
        {

            string alvals = "";

            int rowsnum = this.reptable.RowCount;
            int colsnum = this.reptable.ColumnCount;
            for (int i = 0; i < rowsnum; i++)
            {
                for (int j = 0; j < colsnum; j++)
                {
                    alvals = alvals + this.reptable.GetCell(i, j).InnerText + "\t";
                }
                alvals = alvals + "\n";
            }
            return alvals;
        }

        private string _link = "";
        public string Linkexists
        {
          
            get 
             {
                return _link;
             }
            set 
             {
                _link = value;
             }
        }



        public void ClickHtmlLink(string Linktext)
        {
            if (this.reptable.Exists)
            {
                hlp.LogtoTextFile("Trying to click Link " + Linktext);

                if (this.reptable.Get<CUITe_HtmlSpan>("ControlType=Pane;Id=" + Linktext + ";TagName=SPAN").Exists)
                {
                    this.Linkexists = "true";
                    hlp.LogtoTextFile("link" + Linktext + "Exists: Found it");
                    this.reptable.Get<CUITe_HtmlSpan>("ControlType=Pane;Id=" + Linktext + ";TagName=SPAN;Name=" + null).Click();
                }
                else
                {
                    this.Linkexists = "false";
                    hlp.LogtoTextFile("Link by name " + Linktext + "Does not exist");
                    hlp.AreEqual(" ", Linktext, "Link Existanace", Linktext, "", CompareType.Equal);
                }
            }
            else
            {
                if (this.reportLinksPane.Get<CUITe_HtmlSpan>("ControlType=Pane;Id=" + Linktext + ";TagName=SPAN").Exists)
                {
                    this.Linkexists = "true";
                    hlp.LogtoTextFile("link" + Linktext + "Exists");
                    this.reportLinksPane.Get<CUITe_HtmlSpan>("ControlType=Pane;Id=" + Linktext + ";TagName=SPAN;Name=" + null).Click();
                }
                else
                {
                    this.Linkexists = "false";
                    hlp.LogtoTextFile("Link by name " + Linktext + "Does not exist");
                    hlp.AreEqual(" ", Linktext, "Link Existanace", Linktext, "", CompareType.Equal);
                }
            }

            hlp.LogtoFileCSV(hlp.dtRep);
        }
        public string getWebTableInnertext(string index)
        {
            string strtext = "";
            webtableatIndex(index).WaitForControlReady();
            strtext = webtableatIndex(index).InnerText;
            return strtext;
        }
        public string getWebtableCellInnertxt(string tagindex)
        {
            string rettxt = "";
            webtableatcellIndex(tagindex).WaitForControlReady();
            rettxt = this.webtableatcellIndex(tagindex).InnerText;
            return rettxt;
        }
        public string getWebTableColumnNames(string index)
        {
            string columnnames = "";
          
           
            for(int k=0;   k < webtableatIndex(index).ColumnCount ; k++ )
            {
                columnnames = columnnames +  webtableatIndex(index).GetCell(0,k).InnerText + ";";
            }

            return columnnames.Substring(0, columnnames.Length - 1);
        }

        public void VerifyReportPage( string testcaseNameOrId, string linkName , string tabletpye, string reporttype , string titleWebTableIndex,string expectedTitle,string columnWebTableIndex,string expectedColumnNametext)
        {
            try
            {
                if (reporttype.ToLower() == "html")
                {



                    if (tabletpye.ToLower() == "a")
                    {
                        string actualTitle = getWebTableInnertext(titleWebTableIndex);
                        string actualColumnNameText = getWebTableInnertext(columnWebTableIndex);
                        //   string actualColumnNameText = getWebTableColumnNames(columnWebTableIndex);
                        hlp.AreEqual(testcaseNameOrId, linkName, "Title", expectedTitle, actualTitle, CompareType.Contains);
                        string innertextcompare = expectedColumnNametext.Replace(";", "");
                        if (innertextcompare.Contains("_"))
                        {
                            innertextcompare = innertextcompare.Replace("_", " ");
                        }
                        if (actualColumnNameText.Contains("_"))
                        {
                            actualColumnNameText = actualColumnNameText.Replace("_", " ");
                        }
                        hlp.AreEqual(testcaseNameOrId, linkName, "ColumnNamesText", innertextcompare, actualColumnNameText, CompareType.Contains);
                        // hlp.LogtoTextFile("Col names retreived" + getWebTableColumnNames(columnWebTableIndex));
                    }
                    else if ((tabletpye.ToLower() == "b"))
                    {
                        string actualTitle = getWebtableCellInnertxt(titleWebTableIndex);
                        string actualColumnNameText = getWebTableInnertext(columnWebTableIndex);
                        //   string actualColumnNameText = getWebTableColumnNames(columnWebTableIndex);
                        hlp.AreEqual(testcaseNameOrId, linkName, "Title", expectedTitle, actualTitle, CompareType.Contains);
                        string innertextcompare = expectedColumnNametext.Replace(";", "");
                        if (innertextcompare.Contains("_"))
                        {
                            innertextcompare = innertextcompare.Replace("_", " ");
                        }
                        if (actualColumnNameText.Contains("_"))
                        {
                            actualColumnNameText = actualColumnNameText.Replace("_", " ");
                        }
                        hlp.AreEqual(testcaseNameOrId, linkName, "ColumnNamesText", innertextcompare, actualColumnNameText, CompareType.Contains);
                        // hlp.LogtoTextFile("Col names retreived" + getWebTableColumnNames(columnWebTableIndex));
                    }
                    else if ((tabletpye.ToLower() == "c"))
                    {
                        string actualTitle = this.reportPane.InnerText;
                        string actualColumnNameText = getWebTableInnertext(columnWebTableIndex);

                        //   string actualColumnNameText = getWebTableColumnNames(columnWebTableIndex);
                        if (actualTitle.Length > 120)
                        {
                            actualTitle = actualTitle.Substring(0, 150);
                        }
                        hlp.AreEqual(testcaseNameOrId, linkName, "Title", expectedTitle, actualTitle, CompareType.Contains);
                        string innertextcompare = expectedColumnNametext.Replace(";", "");
                        if (innertextcompare.Contains("_"))
                        {
                            innertextcompare = innertextcompare.Replace("_", " ");
                        }
                        if (actualColumnNameText.Contains("_"))
                        {
                            actualColumnNameText = actualColumnNameText.Replace("_", " ");
                        }
                        hlp.AreEqual(testcaseNameOrId, linkName, "ColumnNamesText", innertextcompare, actualColumnNameText, CompareType.Contains);
                        // hlp.LogtoTextFile("Col names retreived" + getWebTableColumnNames(columnWebTableIndex));
                    }
                }
                else //text reports
                {
                    this.reportEditor.WaitForControlReady();
                    this.reportEditor.Click();
                    Keyboard.SendKeys("^a");
                    Playback.Wait(1000);
                    Keyboard.SendKeys("^c");
                    Playback.Wait(2000);
                    string verifytext = System.Windows.Forms.Clipboard.GetText();
                    hlp.AreEqual(testcaseNameOrId, linkName, "Title", expectedTitle, verifytext, CompareType.Contains);
                    string innertextcompare = expectedColumnNametext.Replace(";", " ");
                    if (innertextcompare.Contains("_"))
                    {
                        innertextcompare = innertextcompare.Replace("_", " ");
                    }
                    if (verifytext.Contains("_"))
                    {
                        verifytext = verifytext.Replace("_", " ");
                    }
                    hlp.AreEqual(testcaseNameOrId, linkName, "ColumnNamesText", innertextcompare, verifytext, CompareType.Contains);
                }
                hlp.LogtoFileCSV(hlp.dtRep);
                hlp.dtRep.Clear();
            }
            catch (Exception ex)
            {
                hlp.LogtoTextFile("Exception from  fucntion VerifyReportPage " + ex.Message);
            }
        }



        public void VerifyChartPage(string testcaseNameOrId, string linkName, string chartIndex ,string expectedTitle, string expectedY1AxisTitle, string expectedY2AxisTitle , string expectedXaxisTitle ,string Expectedlegends)
        {

            string actualTitle = "";
            string actualY1axisTitle = "";
            string actualY2axisTitle = "";
            string actualXaxisTitle = "";
            string actaulLegends = "";
            try
            {

                switch (chartIndex)
                {

                    case "1":
                        {

                            if (this.chtoolbar1.Enabled == true)
                            {

                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn1.Click();
                                }
                                this.btnEditChartaxis1.Click();
                                if (expectedY1AxisTitle.Length > 0)
                                {

                                    // SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedY1AxisTitle);
                                    cmbAxis.SelectedIndex = 1; //y axis title
                                    actualY1axisTitle = this.tbaxistitle.Text;
                                }
                                if (expectedY2AxisTitle.Length > 0)
                                {
                                    // SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedY2AxisTitle);
                                    cmbAxis.SelectedIndex = 2; //y2 axis title
                                    actualY2axisTitle = this.tbaxistitle.Text;
                                }
                                if (expectedXaxisTitle.Length > 0)
                                {
                                    //  SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedXaxisTitle);
                                    cmbAxis.SelectedIndex = 0; //x axis title
                                    actualXaxisTitle = this.tbaxistitle.Text;
                                }


                                btnCancelEditAxis.Click();
                                //if ( this.wndEditAxisSettings.Exists )
                                //{
                                //    btnCancelEditAxis.Click();
                                //}
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn1.Click();
                                }
                                this.btnEditGeneralSettings1.Click();
                                actualTitle = tbchartTitle.Text;
                                this.btnCancelGenChartAxis.Click();


                                this.btnShowLegends1.Click();

                                UITestControlCollection arrlistitems = this.lstLegBox1.Items;

                                foreach (UITestControl indlistitem in arrlistitems)
                                {

                                    actaulLegends = actaulLegends + indlistitem.GetChildren()[0].GetChildren()[0].Name + ";";
                                }
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn1.Click();
                                }

                                this.btnHideLegends1.Click();
                                hlp.AreEqual(testcaseNameOrId, linkName, "ChartTitle", expectedTitle, actualTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "YaxisTitle", expectedY1AxisTitle, actualY1axisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "Y2AxisTitle", expectedY2AxisTitle, actualY2axisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "XAxisTitle", expectedXaxisTitle, actualXaxisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "LegendsText", Expectedlegends, actaulLegends, CompareType.Equal);
                            }
                            else
                            {
                                //no data Available Report Saying No Data found it can be bug or a data condition --- Fail it for reporting 
                                hlp.AreEqual(testcaseNameOrId, linkName, "ChartTitle", expectedTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "YaxisTitle", expectedY1AxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "Y2AxisTitle", expectedY2AxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "XAxisTitle", expectedXaxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "LegendsText", Expectedlegends, "No Data Avialable", CompareType.Equal);
                            }

                            break;
                        }
                    case "2":
                        {

                            if (this.chtoolbar2.Enabled == true)
                            {
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn2.Click();
                                }
                                this.btnEditChartaxis2.Click();
                                if (expectedY1AxisTitle.Length > 0)
                                {

                                    //  SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedY1AxisTitle);
                                    cmbAxis.SelectedIndex = 1; //y axis title
                                    actualY1axisTitle = this.tbaxistitle.Text;
                                }
                                if (expectedY2AxisTitle.Length > 0)
                                {

                                    //  SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedY2AxisTitle);
                                    cmbAxis.SelectedIndex = 2; //y2 axis title
                                    actualY2axisTitle = this.tbaxistitle.Text;
                                }
                                if (expectedXaxisTitle.Length > 0)
                                {

                                    //  SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedXaxisTitle);
                                    cmbAxis.SelectedIndex = 0; //x axis title
                                    actualXaxisTitle = this.tbaxistitle.Text;
                                }
                                btnCancelEditAxis.Click();
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn2.Click();
                                }
                                this.btnEditGeneralSettings2.Click();
                                actualTitle = tbchartTitle.Text;
                                this.btnCancelGenChartAxis.Click();


                                if (this.btnShowLegends2.Enabled == true)
                                {
                                    this.btnShowLegends2.Click();
                                }

                                UITestControlCollection arrlistitems = this.lstLegBox2.Items;

                                foreach (UITestControl indlistitem in arrlistitems)
                                {
                                    actaulLegends = actaulLegends + indlistitem.GetChildren()[0].GetChildren()[0].Name + ";";
                                }
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn2.Click();
                                }
                                this.btnHideLegends2.Click();
                                hlp.AreEqual(testcaseNameOrId, linkName, "ChartTitle", expectedTitle, actualTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "YaxisTitle", expectedY1AxisTitle, actualY1axisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "Y2AxisTitle", expectedY2AxisTitle, actualY2axisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "XAxisTitle", expectedXaxisTitle, actualXaxisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "LegendsText", Expectedlegends, actaulLegends, CompareType.Equal);
                            }
                            else
                            {
                                hlp.AreEqual(testcaseNameOrId, linkName, "ChartTitle", expectedTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "YaxisTitle", expectedY1AxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "Y2AxisTitle", expectedY2AxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "XAxisTitle", expectedXaxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "LegendsText", Expectedlegends, "No Data Avialable", CompareType.Equal);
                            }
                            break;
                        }
                    case "3":
                        {
                            if (chtoolbar3.Enabled == true)
                            {
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn3.Click();
                                }
                                this.btnEditChartaxis3.Click();
                                if (expectedY1AxisTitle.Length > 0)
                                {
                                    //SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedY1AxisTitle);
                                    cmbAxis.SelectedIndex = 1; //y axis title
                                    actualY1axisTitle = this.tbaxistitle.Text;
                                }
                                if (expectedY2AxisTitle.Length > 0)
                                {
                                    // SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedY2AxisTitle);
                                    cmbAxis.SelectedIndex = 2; //y2 axis title
                                    actualY2axisTitle = this.tbaxistitle.Text;
                                }
                                if (expectedXaxisTitle.Length > 0)
                                {
                                    //SelectComBoBoxCutomusingkeyboard(this.cmbAxis, expectedXaxisTitle);
                                    cmbAxis.SelectedIndex = 0; //x axis title
                                    actualXaxisTitle = this.tbaxistitle.Text;
                                }
                                btnCancelEditAxis.Click();
                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn3.Click();
                                }
                                this.btnEditGeneralSettings3.Click();
                                actualTitle = tbchartTitle.Text;
                                this.btnCancelGenChartAxis.Click();


                                this.btnShowLegends3.Click();

                                UITestControlCollection arrlistitems = this.lstLegBox3.Items;

                                foreach (UITestControl indlistitem in arrlistitems)
                                {
                                    actaulLegends = actaulLegends + indlistitem.GetChildren()[0].GetChildren()[0].Name + ";";
                                }

                                if (linkName == "Max/Min Run Time SPC Chart" || linkName == "Max_Min and Run Time")
                                {
                                    this.btnchartOverflowBtn3.Click();
                                }

                                this.btnHideLegends3.Click();
                                hlp.AreEqual(testcaseNameOrId, linkName, "ChartTitle", expectedTitle, actualTitle, CompareType.Contains);
                                hlp.AreEqual(testcaseNameOrId, linkName, "YaxisTitle", expectedY1AxisTitle, actualY1axisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "Y2AxisTitle", expectedY2AxisTitle, actualY2axisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "XAxisTitle", expectedXaxisTitle, actualXaxisTitle, CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "LegendsText", Expectedlegends, actaulLegends, CompareType.Equal);
                            }
                            else
                            {
                                hlp.AreEqual(testcaseNameOrId, linkName, "ChartTitle", expectedTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "YaxisTitle", expectedY1AxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "Y2AxisTitle", expectedY2AxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "XAxisTitle", expectedXaxisTitle, "No Data Avialable", CompareType.Equal);
                                hlp.AreEqual(testcaseNameOrId, linkName, "LegendsText", Expectedlegends, "No Data Avialable", CompareType.Equal);
                            }
                            break;
                        }
                    default : 
                        {
                            break;
                        }
                }

                hlp.LogtoFileCSV(hlp.dtRep);
                hlp.dtRep.Clear();
            }
            catch (Exception ex)
            {
                hlp.LogtoTextFile("Exception from  function VerifyChartPage: Error with LinkName " + linkName);
                hlp.LogtoTextFile("Exception from  function VerifyChartPage: " + ex.Message );
            }
        }

        #endregion

        public void SelectComBoBoxCutomusingkeyboard(CUITe_WpfComboBox cmbcntrl, string valuetoselect)
        {
          //  cmbcntrl.Click();
          //  Keyboard.SendKeys("{Enter}", System.Windows.Input.ModifierKeys.None);
            for (int i = 0; i < cmbcntrl.Items.Count; i++)
            {
                string  txt = cmbcntrl.Items[i].GetChildren()[0].FriendlyName;
                if (txt.ToLower().Trim() == valuetoselect.ToLower().Trim())
                {
                    cmbcntrl.SelectedIndex = i;
                    break;
                   
                }

              //  Keyboard.SendKeys("{Down}", System.Windows.Input.ModifierKeys.None);
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
    }
}

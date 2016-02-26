using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using CUITe.Controls.WpfControls;
using CUITe.Controls.WinControls;
using System.Windows.Input;


namespace Lowis_Reports_Testing.ObjectLibrary
{
   public  class LowisMainWindow : CUITe_WinWindow
    {
       Helper hlp = new Helper();
        #region ObjectStructure
        public LowisMainWindow() : base("ControlName=BrowserMain") { }

        #region Images
        public CUITe_WinControl<WinControl> Start { get { return Get<CUITe_WinControl<WinControl>>("Name=Start;ControlType=Image"); } }
        #endregion
        #region Buttons

        public CUITe_WinButton Analysis { get { return Get<CUITe_WinButton>("Name=Analysis"); } }
        public CUITe_WinButton RefreshWells { get { return AllWellstoolbar.Get<CUITe_WinButton>("Instance=1"); } }
        public CUITe_WinButton SearchWell { get { return AllWellstoolbar.Get<CUITe_WinButton>("Name=Search well"); } }
        public CUITe_WinButton btnCloseLowis { get { return this.titleBar_Lowis.Get<CUITe_WinButton>("Name=Close"); } }
      
        #endregion

        #region Menuitems
        public CUITe_WinMenuItem Reports
        {
            get { return LowisMenucontianer.Get<CUITe_WinMenuItem>("Name=.Reports");
            }
        }

        public WinMenuItem LReports
        {
            get
            {
                WinMenuItem item1 = new WinMenuItem(this.LMenu);
                item1.SearchProperties.Add("Name", ".Reports");
                item1.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
                return item1;
            }

        }
    
        public WinMenuItem LMenuItem2(string strnmae)
        {
            
            {
                WinMenuItem item2= new WinMenuItem(this.LMenu);
                item2.SearchProperties.Add("Name", strnmae);
                item2.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
                return item2;
            }

        }
        public CUITe_WinMenuItem BeamReports
        {
            get { return LowisMenucontianer.Get<CUITe_WinMenuItem>("Name=Beam Reports"); }
        }
        #endregion

        #region Toolbar

        public CUITe_WinToolBar AllWellstoolbar { get { return WndAllWellsToolbar.Get<CUITe_WinToolBar>("Instance=1"); } }

        #endregion

        #region Window
        public CUITe_WinWindow LowisMenucontianer { get { return Get<CUITe_WinWindow>("ClassName=#32768;AccessibleName=Context"); } }
        public CUITe_WinWindow WndAllWellsToolbar { get { return Get<CUITe_WinWindow>("ClassName=ToolbarWindow32;Instance=5"); } }

        public WinWindow LMenu
        {
            get
            {
                WinWindow Lwin = new WinWindow();
                Lwin.SearchProperties.Add("ClassName", "#32768", "AccessibleName", "Context");
                return Lwin;
            }

        }

      
        private WinWindow winLowisMain
        {
            get
            {
                WinWindow ParentWindow = new WinWindow();
                ParentWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, "LOWIS:", PropertyExpressionOperator.Contains);
                ParentWindow.SearchProperties.Add("ClassName", "WindowsForms10.Window", PropertyExpressionOperator.Contains);
                return ParentWindow;
            }
        }

        private WinWindow winLowisClock
        {
            get
            {
                WinWindow containerWindow = new WinWindow(this.winLowisMain);
                containerWindow.SearchProperties.Add("ControlId", "201");
                return containerWindow;
            }
        }

        public void clickMenuitem(string strtxt)
        {
             Mouse.Click(LMenuItem2(strtxt));
        }

        public void clickMenuitem(string strmenuitem1, string strmenuitem2)
        {
            {
                WinMenuItem item11 = new WinMenuItem(this.LMenu);
                item11.SearchProperties.Add("Name", strmenuitem1);
                item11.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
                Mouse.Click(item11);
                WinMenuItem item12 = new WinMenuItem(item11);
                item12.SearchProperties.Add("Name", strmenuitem2);
                item12.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
                Mouse.Click(item12);
            }
        }
        #endregion 

        #region CheckBoxes
        public CUITe_WinCheckBoxTreeItem WellTypes
        {
            get
            {
                return Get<CUITe_WinCheckBoxTreeItem>("Name=Well Types");
            }

        }

        public CUITe_WinCheckBoxTreeItem AllBeamWells
        {
            get
            {
                return Get<CUITe_WinCheckBoxTreeItem>("Name=All Beam Wells");
            }

        }
        public CUITe_WinCheckBoxTreeItem AllESPWells
        {
            get
            {
                return Get<CUITe_WinCheckBoxTreeItem>("Name=All ESP Wells");
            }

        }
        public CUITe_WinCheckBoxTreeItem AllPCPWells
        {
            get
            {
                return Get<CUITe_WinCheckBoxTreeItem>("Name=All PCP Wells");
            }

        }
        public CUITe_WinCheckBoxTreeItem All
        {
            get
            {
                return Get<CUITe_WinCheckBoxTreeItem>("Name=All;Instance=2");
            }

        }
        public CUITe_WinCheckBoxTreeItem AllWels
        {
            get
            {
                return Get<CUITe_WinCheckBoxTreeItem>("Name=All Wells;Instance=1");
            }

        }
        
        #endregion

        #region TitleBar
        public CUITe_WinTitleBar titleBar_Lowis { get { return Get<CUITe_WinTitleBar>("Instance=1"); } }
        #endregion 

        #region WinTextTimecounter
        public WinText lowisClock
        {
            get
            {
                WinText txt = null;
              //  WinToolBar tlbar = new WinToolBar(this.winLowisClock);
                txt = new WinText(this.winLowisClock);
                UITestControlCollection alltxt = txt.FindMatchingControls();
                string disptxt = ((WinText)alltxt[0]).DisplayText;
                WinText stxt = new WinText(this.winLowisClock);
                return stxt;
            }

        }
        #endregion

        #region WinEdit
      
        #endregion
        #region Wintabs
        public CUITe_WinTabPage wintaballseparatorreport { get { return Get<CUITe_WinTabPage>("Name=All Separator Reports"); } }
        #endregion 

        #region WinList

        #endregion
        #endregion

        // ========================================

        #region ObjectMethods
        public string getLowisTimeText()
        {
            string txtTime = "";
            if (this.lowisClock.Exists)
            {
              //  this.lowisClock.DrawHighlight();
                this.lowisClock.WaitForControlReady();
                 txtTime = this.lowisClock.DisplayText;
            }
            else
            {
                 txtTime = "";
            }
           return txtTime;
        }
        public void  lowisDwait()
        {
            string txttime1 = this.getLowisTimeText();
            Playback.Wait(1000);
            string txttime2 = this.getLowisTimeText();
            while (txttime1 == txttime2)
            {
                Playback.Wait(1000);
                txttime2 = this.getLowisTimeText();
            }
            Playback.Wait(1000);
        }
        public void SelectWellfromSearch(string wellname)
        {
            this.SearchWell.Click();
            //this.tbSearchWell.Click();
            //Keyboard.SendKeys(wellname, ModifierKeys.None);
            LowisSearchWindow lswindow = new LowisSearchWindow();
            lswindow.tbSearchWell.Text = wellname;
            lswindow.lstSearchWell.SetFocus();
            //lswindow.lstSearchWell.SelectedItem = wellname;
            int lstind =0;
            foreach(string item in lswindow.lstSearchWell.Items.GetNamesOfControls())
            {
                hlp.LogtoTextFile("Drop downItem text" + item);
                if (item.Trim() == wellname.Trim() )
                {
                    lswindow.lstSearchWell.SelectedIndex = lstind;
                    hlp.LogtoTextFile("Dropd owm match with" + wellname);
                }
                lstind++;
            }
            lswindow.btnSearchWellOK.Click();
        }
        public void ClickYestoAutoUpdates()
        {
            LowisAutoUpdate autoupdt = new LowisAutoUpdate();
            autoupdt.BtnAutoUpYes.Click();
        }
        #endregion
    }

   public class LowisSearchWindow : CUITe_WinWindow
   {
       public LowisSearchWindow() : base("Name=Search Well") { }
       #region WinWindow
       public CUITe_WinWindow WndSearchWell { get { return Get<CUITe_WinWindow>("Name=Search Well;ClassName=#32770"); } }
       public CUITe_WinWindow WndSearchWellEdit { get { return WndSearchWell.Get<CUITe_WinWindow>("ControlId=1007;Instance=1"); } }
       public CUITe_WinWindow WndSearchWellList { get { return WndSearchWell.Get<CUITe_WinWindow>("ControlId=1008;Instance=1"); } }
       public CUITe_WinWindow WndSearchWellOK { get { return WndSearchWell.Get<CUITe_WinWindow>("ControlId=1;Instance=1"); } }
       #endregion

       #region WinEdit
       public CUITe_WinEdit tbSearchWell { get { return this.WndSearchWellEdit.Get<CUITe_WinEdit>("Instance=1"); } }
       #endregion

       #region WinList
       public CUITe_WinList lstSearchWell { get { return this.WndSearchWellList.Get<CUITe_WinList>("Instance=1"); } }
       #endregion

       #region Buttons
       public CUITe_WinButton btnSearchWellOK { get { return this.WndSearchWellOK.Get<CUITe_WinButton>("Name=OK"); } }
       #endregion
   }

   public class LowisAutoUpdate : CUITe_WinWindow
   {
       public LowisAutoUpdate() : base("Name=Auto Updates") { }
       public CUITe_WinWindow WndAutoUpdate { get { return Get<CUITe_WinWindow>("Name=Auto Updates;ClassName=#32770"); } }
       public CUITe_WinButton BtnAutoUpYes { get { return Get<CUITe_WinButton>("Name=Yes"); } }

   }
   
}

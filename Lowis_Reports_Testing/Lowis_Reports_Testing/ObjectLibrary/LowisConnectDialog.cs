using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using CUITe.Controls.WinControls;
using CUITe.Controls.WpfControls;
using CUITe.Controls.HtmlControls;



namespace Lowis_Reports_Testing.ObjectLibrary
{

    public class LowisConnectDialog : CUITe_WpfWindow
    {
        public LowisConnectDialog() : base("Name=LOWIS: Connect") { }
        #region combobox
        public CUITe_WpfComboBox serverName { get { return Get<CUITe_WpfComboBox>("Instance=1"); } }
        #endregion
        #region Buttons
        public CUITe_WpfButton Connect { get { return Get<CUITe_WpfButton>("Instance=3"); } }
        public CUITe_WpfButton Settings { get { return Get<CUITe_WpfButton>("Instance=1"); } }

        #endregion 
        #region Textboxes
        public CUITe_WpfEdit serverNametb { get { return this.serverName.Get<CUITe_WpfEdit>("Automationid=PART_EditableTextBox"); } }
        public CUITe_WpfEdit txtuserName { get { return this.serverName.Get<CUITe_WpfEdit>("Automationid=PART_EditableTextBox"); } }
        public CUITe_WpfEdit txtPassword { get { return this.serverName.Get<CUITe_WpfEdit>("Automationid=PART_EditableTextBox"); } }
        #endregion
        #region radiobuttons
       // Use the following credentials:
        public CUITe_WinRadioButton usecredentails { get {  return Get<CUITe_WinRadioButton>("Name=Use the following credentials:"); } }
#endregion 


        #region ObjectMethods
        public void selectServer(string serverName)
        {
           // this.serverName.EditableItem = serverName;
         //   this.serverNametb.Text = serverName;
            this.serverNametb.Click();
            Keyboard.SendKeys("{Home}", ModifierKeys.None);
            Keyboard.SendKeys("{End}", ModifierKeys.Shift);
            Keyboard.SendKeys("{Del}", ModifierKeys.None);
            Keyboard.SendKeys(serverName, ModifierKeys.None);

            //System.Windows.Forms.SendKeys.SendWait("{Home}");
            //System.Windows.Forms.SendKeys.SendWait("+{End}");
            //System.Windows.Forms.SendKeys.SendWait("{Del}");
            //System.Windows.Forms.SendKeys.SendWait(serverName);

        }

        public void btnClick(string buttonname)
        {
            if (buttonname.ToLower() == "connect")
            {
                Connect.Click();
            }
            if (buttonname.ToLower() == "settings")
            {
                Settings.Click();
            }
        }
        #endregion
    }

    public class LowisSettingsDialog : CUITe_WpfWindow
    {
        public LowisSettingsDialog()  : base("Name=LOWIS: Settings") { }
        public CUITe_WpfEdit storepath { get { return Get<CUITe_WpfEdit>("Instance=1"); } }
        public CUITe_WpfButton btnSavesettings { get { return Get<CUITe_WpfButton>("Instance=3"); } }
    }
}

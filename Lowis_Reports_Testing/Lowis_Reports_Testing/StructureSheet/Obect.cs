using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lowis_Reports_Testing.ObjectLibrary;
using System.Data;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using System.Windows.Automation;
using System.Xml.Serialization;
using System.Threading;

namespace Lowis_Reports_Testing.StructureSheet
{
    class UIObect
    {


        Helper hlp = new Helper();

        public void AddData(string filename, string testcase)
        {

            try
            {
                DataTable dt1 = hlp.dtFromExcelFile(filename, "Template");
                DataTable dt2 = hlp.dtFromExcelFile(filename, "ExpectedData", "TestCase", testcase);
                UITestControl UIcurrentparent = null;
                bool inmemlist = false;
                string inmemprevsearchvalue = "";
                AutomationElementCollection objcol = null;
                AutomationElementCollection objcol2 = null;
                AutomationElement paneobject = null;
                List<UITestControl> pvmasks = new List<UITestControl>();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    string controlValue = "";
                    string searchBy2 = ""; string searchValue2 = "";
                    string parentType = dt1.Rows[i]["ParentType"].ToString();
                    string parentSearchBy = dt1.Rows[i]["parentSearchBy"].ToString();
                    string parentSearchValue = dt1.Rows[i]["parentSearchValue"].ToString();
                    string pTechnology = dt1.Rows[i]["Technology"].ToString();
                    string controlType = dt1.Rows[i]["ControlType"].ToString();
                    string technologyControl = dt1.Rows[i]["TechnologyControl"].ToString();
                    string field = dt1.Rows[i]["Field"].ToString();
                    string action = dt1.Rows[i]["Action"].ToString();
                    string index = dt1.Rows[i]["Index"].ToString();
                    string pindex = dt1.Rows[i]["pindex"].ToString();
                    string searchBy = dt1.Rows[i]["SearchBy"].ToString();
                    string searchValue = dt1.Rows[i]["SearchValue"].ToString();
                    string pOperator = dt1.Rows[i]["pOperator"].ToString();
                    string cOperator = dt1.Rows[i]["cOperator"].ToString();
                    string resetparent = dt1.Rows[i]["ResetParent"].ToString();
                    if (i != 0)
                    {
                        inmemprevsearchvalue = dt1.Rows[i - 1]["SearchValue"].ToString();
                    }
                    if (IsColumnPresent("SearchBy2", dt1))
                    {
                        searchBy2 = dt1.Rows[i]["SearchBy2"].ToString();
                    }
                    if (IsColumnPresent("SearchValue2", dt1))
                    {
                        searchValue2 = dt1.Rows[i]["SearchValue2"].ToString();
                    }
                    if (field.Length > 0)
                    {
                        controlValue = dt2.Rows[0][field].ToString();
                    }

                    if (parentType.Length > 0)
                    {
                        switch (parentType.ToLower())
                        {

                            #region Window
                            case "window":
                                {
                                    try
                                    {
                                        if (resetparent == "1")
                                        {
                                            UIcurrentparent = null;
                                        }
                                        if (pTechnology.ToLower() == "msaa")
                                        {
                                            WinWindow uiwindow = null;
                                            if (UIcurrentparent == null)
                                            {
                                                uiwindow = new WinWindow();
                                            }
                                            else
                                            {
                                                uiwindow = new WinWindow(UIcurrentparent);
                                            }
                                            if (pOperator == "=")
                                            {
                                                uiwindow.SearchProperties.Add(parentSearchBy, parentSearchValue);
                                                if (pindex.Length > 0)
                                                {
                                                    uiwindow.SearchProperties.Add("Instance", pindex);
                                                }
                                            }
                                            else if (pOperator == "~")
                                            {
                                                uiwindow.SearchProperties.Add(parentSearchBy, parentSearchValue, PropertyExpressionOperator.Contains);
                                            }
                                            UIcurrentparent = uiwindow;
                                        }
                                        else if (pTechnology == "uia")
                                        {
                                            WpfWindow uiwindow = new WpfWindow();
                                            if (pOperator == "=")
                                            {
                                                uiwindow.SearchProperties.Add(parentSearchBy, parentSearchValue);
                                            }
                                            else if (pOperator == "~")
                                            {
                                                uiwindow.SearchProperties.Add(parentSearchBy, parentSearchValue, PropertyExpressionOperator.Contains);
                                            }
                                            UIcurrentparent = uiwindow;

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        hlp.LogtoTextFile("error occured" + ex.Message);
                                    }
                                    break;
                                }
                            #endregion

                            #region client
                            case "client":
                                {
                                    if (pTechnology.ToLower() == "msaa")
                                    {
                                        WinClient uicleint = new WinClient(UIcurrentparent);
                                        if (pOperator == "=")
                                        {
                                            uicleint.SearchProperties.Add(parentSearchBy, parentSearchValue);
                                        }
                                        else if (pOperator == "~")
                                        {
                                            uicleint.SearchProperties.Add(parentSearchBy, parentSearchValue, PropertyExpressionOperator.Contains);
                                        }
                                        UIcurrentparent = uicleint;
                                    }
                                    else if (pTechnology.ToLower() == "uia")
                                    {
                                        // to do
                                    }
                                    break;
                                }
                            #endregion

                            #region docuemnt
                            case "document":
                                {
                                    if (pTechnology.ToLower() == "web")
                                    {
                                        HtmlDocument uidoc = new HtmlDocument(UIcurrentparent);
                                        if (pOperator == "=")
                                        {
                                            uidoc.SearchProperties.Add(parentSearchBy, parentSearchValue);
                                        }
                                        else if (pOperator == "~")
                                        {
                                            uidoc.SearchProperties.Add(parentSearchBy, parentSearchValue, PropertyExpressionOperator.Contains);
                                        }
                                        UIcurrentparent = uidoc;
                                    }
                                    else if (pTechnology.ToLower() == "uia")
                                    {

                                        // to do
                                    }
                                    break;
                                }

                            #endregion
                            #region Pane
                            case "pane":
                                {

                                    if (pTechnology.ToLower() == "web")
                                    {
                                        HtmlControl pane = new HtmlControl(UIcurrentparent);
                                        if (pOperator == "=")
                                        {
                                            pane.SearchProperties.Add(parentSearchBy, parentSearchValue);
                                        }
                                        else if (pOperator == "~")
                                        {
                                            pane.SearchProperties.Add(parentSearchBy, parentSearchValue, PropertyExpressionOperator.Contains);
                                        }
                                        UIcurrentparent = pane;
                                    }
                                    else if (pTechnology.ToLower() == "uia")
                                    {

                                        // to do
                                    }
                                    break;
                                }
                            #endregion
                            #region Custom
                            case "custom":
                                {

                                    if (pTechnology.ToLower() == "web")
                                    {
                                        HtmlCustom uicustm = new HtmlCustom(UIcurrentparent);
                                        if (pOperator == "=")
                                        {
                                            uicustm.SearchProperties.Add(parentSearchBy, parentSearchValue);
                                        }
                                        else if (pOperator == "~")
                                        {
                                            uicustm.SearchProperties.Add(parentSearchBy, parentSearchValue, PropertyExpressionOperator.Contains);
                                        }
                                        UIcurrentparent = uicustm;
                                    }
                                    else if (pTechnology.ToLower() == "uia")
                                    {

                                        // to do
                                    }
                                    break;
                                }
                            #endregion

                        }
                    }
                    if (controlType.Length > 0)
                    {
                        switch (controlType.ToLower())
                        {
                            #region edit
                            case "edit":
                                {
                                    if (technologyControl.ToLower() == "msaa")
                                    {
                                        WinEdit uiedit = new WinEdit(UIcurrentparent);
                                        if (cOperator == "=")
                                        {
                                            uiedit.SearchProperties.Add(searchBy, searchValue);
                                        }
                                        else if (cOperator == "~")
                                        {
                                            uiedit.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                        }

                                        if (controlValue.Length > 0)
                                        {

                                            uiedit.Text = controlValue;
                                        }
                                    }
                                    else if (technologyControl == "UIA")
                                    {
                                        // to do
                                    }
                                    else if (technologyControl == "Web")
                                    {
                                        HtmlEdit uiedit = new HtmlEdit(UIcurrentparent);
                                        if (cOperator == "=")
                                        {
                                            uiedit.SearchProperties.Add(searchBy, searchValue);
                                        }
                                        else if (cOperator == "~")
                                        {
                                            uiedit.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                        }

                                        if (controlValue.Length > 0)
                                        {

                                            uiedit.Text = controlValue;
                                        }
                                    }



                                    break;
                                }
                            #endregion
                            #region button
                            case "button":
                                {
                                    #region MSAAAButton
                                    try
                                    {
                                        if (technologyControl == "MSAA")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {
                                                WinButton ucntl = new WinButton(UIcurrentparent);
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("Instance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }

                                                if (controlValue.Length > 0)
                                                {
                                                    //  System.Drawing.Point p = new System.Drawing.Point(ucntl.BoundingRectangle.X,ucntl.BoundingRectangle.Y);
                                                    //   bool isvisible = ucntl.TryGetClickablePoint(out p);
                                                    Mouse.Click(ucntl);
                                                    hlp.LogtoTextFile(string.Format("Clicked Button {0}", field));
                                                }
                                            }
                                            else
                                            {
                                                hlp.LogtoTextFile("Parent of Control button was not constructed: Hence  Any Actions on this control are not peformed] ");
                                            }
                                        }
                                    #endregion
                                        else if (technologyControl == "UIA")
                                        {
                                            // to do
                                        }
                                        #region Webbutton
                                        else if (technologyControl == "Web")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {
                                                HtmlButton ucntl = new HtmlButton(UIcurrentparent);
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("TagInstance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }
                                                if (ucntl.Exists)
                                                {
                                                    if (controlValue.Length > 0)
                                                    {

                                                        Mouse.Click(ucntl);
                                                    }
                                                }
                                                else
                                                {
                                                    hlp.LogtoTextFile("Unable to Construct Control Button: [ Any Actions on this control are not peformed] ");
                                                }
                                            }
                                        }
                                        #endregion

                                    }
                                    catch (Exception ex)
                                    {
                                        hlp.LogtoTextFile("error while construcitng button [ UI may not exist ] " + ex.Message);
                                    }


                                    break;
                                }
                            #endregion
                            #region checkbox
                            case "checkbox":
                                {
                                    #region MSAAACheckbox
                                    try
                                    {
                                        if (technologyControl == "MSAA")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {
                                                WinCheckBox ucntl = new WinCheckBox(UIcurrentparent);
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("Instance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }

                                                if (controlValue.Length > 0)
                                                {
                                                    //  System.Drawing.Point p = new System.Drawing.Point(ucntl.BoundingRectangle.X,ucntl.BoundingRectangle.Y);
                                                    //   bool isvisible = ucntl.TryGetClickablePoint(out p);
                                                    if (controlValue == "0")
                                                    {
                                                        ucntl.Checked = false;
                                                    }
                                                    else if (controlValue == "1")
                                                    {
                                                        ucntl.Checked = false;
                                                    }
                                                    else
                                                    {
                                                        hlp.LogtoTextFile(string.Format("Nocation for unsupported value {0}", controlValue));
                                                    }
                                                    hlp.LogtoTextFile(string.Format("Clicked Button {0}", field));
                                                }
                                            }
                                            else
                                            {
                                                hlp.LogtoTextFile("Parent of Control button was not constructed: Hence  Any Actions on this control are not peformed] ");
                                            }
                                        }
                                    #endregion
                                        else if (technologyControl == "UIA")
                                        {
                                            // to do
                                        }
                                        #region WebbCheckbox
                                        else if (technologyControl == "Web")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {
                                                HtmlCheckBox ucntl = new HtmlCheckBox(UIcurrentparent);
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("TagInstance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }
                                                if (ucntl.Exists)
                                                {
                                                    if (controlValue.Length > 0)
                                                    {
                                                        //  System.Drawing.Point p = new System.Drawing.Point(ucntl.BoundingRectangle.X,ucntl.BoundingRectangle.Y);
                                                        //   bool isvisible = ucntl.TryGetClickablePoint(out p);
                                                        if (controlValue == "0")
                                                        {
                                                            ucntl.Checked = false;
                                                        }
                                                        else if (controlValue == "1")
                                                        {
                                                            ucntl.Checked = false;
                                                        }
                                                        else
                                                        {
                                                            hlp.LogtoTextFile(string.Format("Nocation for unsupported value {0}", controlValue));
                                                        }
                                                        hlp.LogtoTextFile(string.Format("Clicked Button {0}", field));
                                                    }
                                                }
                                                else
                                                {
                                                    hlp.LogtoTextFile("Unable to Construct Control Checkbox : [ Any Actions on this control are not peformed] ");
                                                }
                                            }
                                        }
                                        #endregion

                                    }
                                    catch (Exception ex)
                                    {
                                        hlp.LogtoTextFile("error while construcitng button [ UI may not exist ] " + ex.Message);
                                    }


                                    break;
                                }
                            #endregion
                            #region radiobutton
                            case "radiobutton":
                                {
                                    #region MSAAARadiobutton
                                    try
                                    {
                                        if (technologyControl == "MSAA")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {
                                                WinRadioButton ucntl = new WinRadioButton(UIcurrentparent);
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("Instance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }

                                                if (controlValue.Length > 0)
                                                {
                                                    //  System.Drawing.Point p = new System.Drawing.Point(ucntl.BoundingRectangle.X,ucntl.BoundingRectangle.Y);
                                                    //   bool isvisible = ucntl.TryGetClickablePoint(out p);
                                                   
                                                    hlp.LogtoTextFile(string.Format("Clicked Button {0}", field));
                                                }
                                            }
                                            else
                                            {
                                                hlp.LogtoTextFile("Parent of Control button was not constructed: Hence  Any Actions on this control are not peformed] ");
                                            }
                                        }
                                    #endregion
                                        else if (technologyControl == "UIA")
                                        {
                                            // to do
                                        }
                                        #region WebbRadio
                                        else if (technologyControl == "Web")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {

                                                hlp.LogtoTextFile("Performing action on:  " + field);
                                                HtmlRadioButton ucntl = new HtmlRadioButton(UIcurrentparent);
                                                
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (searchBy2 == "grouptext")
                                                    {
                                                        List<string> stringlist = searchValue2.Split(';').ToList<string>();
                                                        int result = stringlist.FindIndex(X => (X == controlValue));
                                                        hlp.LogtoTextFile("Got Id " + result);
                                                        ucntl.SearchProperties.Add("id", result.ToString());
                                                    }
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("TagInstance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }
                                                if (ucntl.Exists)
                                                {
                                                    
                                                 Mouse.Click(ucntl);           
                                                        
                                                }
                                                else
                                                {
                                                    hlp.LogtoTextFile("Unable to Construct Control RadioButton : [ Any Actions on this control are not peformed] ");
                                                }
                                            }
                                        }
                                        #endregion

                                    }
                                    catch (Exception ex)
                                    {
                                        hlp.LogtoTextFile("error while construcitng button [ UI may not exist ] " + ex.Message);
                                    }


                                    break;
                                }
                            #endregion
                            #region image
                            case "image":
                                {

                                    try
                                    {

                                        if (technologyControl == "MSAA")
                                        {
                                            // to do
                                        }
                                        else if (technologyControl == "UIA")
                                        {
                                            // to do
                                        }
                                        else if (technologyControl == "Web")
                                        {
                                            HtmlImage ucntl = new HtmlImage(UIcurrentparent);
                                            if (cOperator == "=")
                                            {
                                                ucntl.SearchProperties.Add(searchBy, searchValue);
                                                if (index.Length > 0)
                                                {
                                                    ucntl.SearchProperties.Add("TagInstance", index);
                                                }
                                            }
                                            else if (cOperator == "~")
                                            {
                                                ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                            }

                                            if (controlValue.Length > 0)
                                            {

                                                Mouse.Click(ucntl);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        hlp.LogtoTextFile("Exception frmo Control Type = Image " + "  Field Name : " + field + "Mesage: " + ex.Message);
                                    }
                                    break;
                                }
                            #endregion
                            #region Dropdown
                            case "dropdown":
                                {
                                    if (technologyControl == "MSAA")
                                    {
                                        // to do
                                    }
                                    else if (technologyControl == "UIA")
                                    {
                                        // to do
                                    }
                                    else if (technologyControl == "Web")
                                    {
                                        HtmlComboBox ucntl = new HtmlComboBox(UIcurrentparent);
                                        if (cOperator == "=")
                                        {
                                            if (searchValue.Length > 0)
                                            {
                                                ucntl.SearchProperties.Add(searchBy, searchValue);
                                            }
                                            if (index.Length > 0)
                                            {
                                                ucntl.SearchProperties.Add("TagInstance", index);
                                            }
                                        }
                                        else if (cOperator == "~")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                        }

                                        if (controlValue.Length > 0)
                                        {
                                            hlp.LogtoTextFile("web dropdown value to be selected " + controlValue);
                                            hlp.LogtoTextFile("web dropdown value curently  selected in dropdown " + ucntl.SelectedItem);
                                            ucntl.SelectedItem = controlValue;
                                            if (ucntl.SelectedItem != controlValue)
                                            {
                                                hlp.LogtoTextFile("Issue Encountered in setting combobox value");
                                            }
                                            else
                                            {
                                                hlp.LogtoTextFile("Entered Value in ComboBox");
                                            }
                                        }
                                    }
                                    break;
                                }
                            #endregion
                            #region TreeItem
                            case "treeitem":
                                {
                                    if (technologyControl == "MSAA")
                                    {
                                        WinTreeItem ucntl = new WinTreeItem(UIcurrentparent);
                                        if (cOperator == "=")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue);
                                            if (index.Length > 0)
                                            {
                                                ucntl.SearchProperties.Add("Instance", index);
                                            }
                                        }
                                        else if (cOperator == "~")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                        }

                                        if (controlValue.Length > 0)
                                        {

                                            Mouse.Click(ucntl);
                                        }
                                    }
                                    else if (technologyControl == "UIA")
                                    {
                                        // to do
                                    }
                                    else if (technologyControl == "Web")
                                    {

                                    }
                                    break;
                                }
                            #endregion
                            #region FileInput
                            case "fileinput":
                                {
                                    if (technologyControl == "MSAA")
                                    {
                                        WinButton ucntl = new WinButton(UIcurrentparent);
                                        if (cOperator == "=")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue);
                                            if (index.Length > 0)
                                            {
                                                ucntl.SearchProperties.Add("Instance", index);
                                            }
                                        }
                                        else if (cOperator == "~")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                        }

                                        if (controlValue.Length > 0)
                                        {

                                            Mouse.Click(ucntl);
                                        }
                                    }
                                    else if (technologyControl == "UIA")
                                    {
                                        // to do
                                    }
                                    else if (technologyControl == "Web")
                                    {
                                        HtmlFileInput ucntl = new HtmlFileInput(UIcurrentparent);
                                        if (cOperator == "=")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue);
                                            if (index.Length > 0)
                                            {
                                                ucntl.SearchProperties.Add("TagInstance", index);
                                            }
                                        }
                                        else if (cOperator == "~")
                                        {
                                            ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                        }

                                        if (controlValue.Length > 0)
                                        {
                                            Mouse.Click(ucntl);
                                            ucntl.FileName = controlValue;

                                        }
                                    }
                                    break;
                                }
                            #endregion
                            #region UPane
                            case "upane":
                                {
                                    AutomationElement reqlblelem = null;
                                    if (searchValue == inmemprevsearchvalue)
                                    {
                                        inmemlist = true;
                                    }
                                    else
                                    {
                                        inmemlist = false;
                                    }
                                    #region ConstructObjectCollection
                                    if (inmemlist == false)
                                    {
                                        hlp.LogtoTextFile(string.Format("Inside One Time construction of collection of labels and controls "));
                                        AutomationElement rootelem = AutomationElement.RootElement;
                                        Condition CondSysDateTimePick32 = null;
                                        switch (searchBy.ToLower())
                                        {
                                            case "classname":
                                                {
                                                    CondSysDateTimePick32 = new AndCondition(

                                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Pane),
                                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ClassNameProperty, searchValue)
                                                           );
                                                    break;
                                                }
                                        }
                                        for (int tk = 0; tk < 15; tk++)
                                        {
                                            objcol = rootelem.FindAll(TreeScope.Descendants, CondSysDateTimePick32);
                                            Thread.Sleep(1000);
                                            if (objcol.Count > 0)
                                            {
                                                hlp.LogtoTextFile(string.Format("Upane:  found in {0} Attempt....", tk));
                                                break;
                                            }
                                        }

                                        // get collection for Label Texts Or simply Control Type Text labels only for Dialogs not webpanes
                                        if (searchValue.ToLower() == "pvmaskedit" || searchValue.ToLower() == "pvnumeric")
                                        {
                                            AutomationElement dlgwindow = rootelem.FindFirst(TreeScope.Descendants,
                                                new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                                                (new PropertyCondition(AutomationElement.ClassNameProperty, "#32770")
                                                                )));

                                            Condition ConditionLabelSearch = null;
                                            ConditionLabelSearch =
                                                           new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text);
                                            objcol2 = dlgwindow.FindAll(TreeScope.Descendants, ConditionLabelSearch);
                                            hlp.LogtoTextFile(string.Format("SaerchBy and last asearchby valeus {0} {1}", searchValue, inmemprevsearchvalue));
                                        }

                                    }
                                    #endregion
                                    #region SearchByLabel
                                    if (searchBy2 == "label") //do search using labels wherever possible 
                                    {
                                        hlp.LogtoTextFile("Upane:  Searching by label...."+ searchValue2);
                                        hlp.LogtoTextFile(string.Format("Label Collection count: {0}", objcol2.Count));
                                        foreach (AutomationElement lbl in objcol2)
                                        {
                                           // hlp.LogtoTextFile(string.Format("Got label bname as {0} ",lbl.Current.Name));
                                            if (lbl.Current.Name == searchValue2)
                                            {
                                                reqlblelem = lbl; // wefound required label 
                                                hlp.LogtoTextFile("Upane:  found label ..." + searchValue2);
                                                break;
                                            }
                                        }
                                      //  hlp.LogtoTextFile(string.Format("Y cordinate of Label {0} was {1}", searchValue2, reqlblelem.Current.BoundingRectangle.Y));
                                        double lwoffset = reqlblelem.Current.BoundingRectangle.Y - 5;
                                        double hioffset = reqlblelem.Current.BoundingRectangle.Y + 5;
                                        hlp.LogtoTextFile(string.Format("Control  Collection count: {0}", objcol.Count));
                                        foreach (AutomationElement cntll in objcol)
                                        {
                                         //   hlp.LogtoTextFile(string.Format("Y cordinate of control {0} was {1}", searchValue2, cntll.Current.BoundingRectangle.Y));
                                            if ((cntll.Current.BoundingRectangle.Y > lwoffset) && (cntll.Current.BoundingRectangle.Y < hioffset))
                                            {
                                                paneobject = cntll; // wefound required control
                                                hlp.LogtoTextFile("Upane:  found contrl using label  ..." + searchValue2);
                                                break;
                                            }
                                        }
                                        if (paneobject != null)
                                        {
                                            ClickElement(paneobject, action);
                                            if (controlValue.Length > 0)
                                            {
                                                if (searchValue.ToLower() == "pvmaskedit")
                                                {
                                                    Keyboard.SendKeys("{Home}");
                                                    Keyboard.SendKeys("+{End}");
                                                    Keyboard.SendKeys("{Del}");
                                                    Playback.Wait(1000);
                                                }
                                                Keyboard.SendKeys(controlValue);
                                            }
                                        }
                                        else
                                        {
                                            hlp.LogtoTextFile("Unable to find desired UIautomation Pane using label as "+searchValue2);
                                        }

                                    }
                                    #endregion
                                    #region SearchByIndex
                                    else //else  search using indexs if not searchable by label
                                    {
                                        hlp.LogtoTextFile("U pane Searching by Index");
                                        if (objcol.Count > 0)
                                        {
                                            if (index.Length > 0)
                                            {
                                                paneobject = objcol[Int32.Parse(index)];
                                            }
                                            else
                                            {
                                                paneobject = objcol[0];
                                            }
                                            ClickElement(paneobject, action);
                                            if (controlValue.Length > 0)
                                            {
                                                if (searchValue.ToLower() == "pvmaskedit")
                                                {
                                                    Keyboard.SendKeys("{Home}");
                                                    Keyboard.SendKeys("+{End}");
                                                    Keyboard.SendKeys("{Del}");
                                                    Playback.Wait(1000);
                                                }
                                                Keyboard.SendKeys(controlValue);
                                            }
                                        }
                                        else
                                        {
                                            hlp.LogtoTextFile("Unable to find desired UIautomation Pane");
                                        }
                                    }
                                    #endregion
                                }
                                break;
                            #endregion
                            #region text
                            #endregion
                            #region Lowis_pvmaskedit
                            case "pvmaskedit":
                                {
                                    WinClient pvmaskedit = null;
                                    if (inmemlist == false) //construct List only once ...
                                    {

                                        UITestControlCollection ucol1 = UIcurrentparent.GetChildren();
                                        UITestControlCollection ucol2 = ucol1[1].GetChildren();
                                        UITestControlCollection ucol3 = ucol2[0].GetChildren();
                                        UITestControlCollection ucol4 = ucol3[0].GetChildren();
                                        UITestControlCollection ucol5 = ucol4[0].GetChildren();
                                        UITestControlCollection ucol6 = ucol5[0].GetChildren();
                                        foreach (UITestControl ictl2 in ucol6)
                                        {
                                            UITestControlCollection ucol7 = ictl2.GetChildren();
                                            UITestControlCollection ucol8 = ucol7[1].GetChildren();
                                            foreach (UITestControl in33 in ucol8)
                                            {
                                                UITestControlCollection ucol9 = in33.GetChildren();
                                                foreach (UITestControl in9 in ucol9)
                                                {

                                                    hlp.LogtoTextFile("Found control with control type " + in9.ControlType.ToString());
                                                    hlp.LogtoTextFile("Found control with clasanme: " + in9.ClassName.ToString());
                                                    try
                                                    {
                                                        pvmasks.Add(in9.GetChildren()[3].GetChildren()[0].GetChildren()[3]);
                                                        hlp.LogtoTextFile("added Control " + in9.GetChildren()[3].GetChildren()[0].GetChildren()[3].ControlType.ToString());
                                                    }
                                                    catch
                                                    {
                                                    }

                                                }
                                            }


                                        }
                                        inmemlist = true;
                                    }


                                    for (int it = 0; it < pvmasks.Count; it++)
                                    {
                                        if (it == Int32.Parse(index) && pvmasks[it].ControlType.ToString() == "Client")
                                        {
                                            pvmaskedit = pvmasks[it] as WinClient;
                                            break;
                                        }

                                    }
                                    pvmaskedit.SetFocus();
                                    //  pvmaskedit.DrawHighlight();
                                    System.Drawing.Point p = new System.Drawing.Point(pvmaskedit.BoundingRectangle.X, pvmaskedit.BoundingRectangle.Y);
                                    bool controlexist = pvmaskedit.TryGetClickablePoint(out p);
                                    if (controlexist)
                                    {
                                        System.Diagnostics.Trace.WriteLine("control detected Trying to click ");
                                        hlp.LogtoTextFile("control detected Trying to click ");
                                        Mouse.Click();
                                        Keyboard.SendKeys(controlValue);
                                    }
                                    else
                                    {
                                        hlp.LogtoTextFile(string.Format("Unable to detect Control pvmakedit using index {0}", index));
                                    }

                                    break;
                                }
                            #endregion
                            #region UPvcomboBox
                            case "upvcombobox":
                                {
                                    hlp.LogtoTextFile(string.Format("Inside upvcombobox"));
                                    AutomationElement rootelem = AutomationElement.RootElement;
                                    AutomationElement dlgwindow = null;
                                    for (int ik = 0; ik < 15; ik++)
                                    {
                                        dlgwindow = rootelem.FindFirst(TreeScope.Descendants,
                                                         new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                                                         (new PropertyCondition(AutomationElement.ClassNameProperty, "#32770")
                                                                         )));

                                        if (dlgwindow != null)
                                        {
                                            break;
                                        }
                                    }

                                    if (dlgwindow == null)
                                    {
                                        hlp.LogtoTextFile("Unable to get dialog window....");
                                        return;
                                    }
                                        Condition ConditionLabelSearch = null;
                                        ConditionLabelSearch =
                                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text);
                                        objcol2 = dlgwindow.FindAll(TreeScope.Descendants, ConditionLabelSearch);
                                        hlp.LogtoTextFile(string.Format("SaerchBy and last asearchby valeus {0} {1}", searchValue, inmemprevsearchvalue));
                                  //  }
                                        if (objcol2.Count > 0)
                                        {
                                            AutomationElement reqlblelem = null;
                                            hlp.LogtoTextFile("Upane:  Searching by label....");
                                            foreach (AutomationElement lbl in objcol2)
                                            {
                                                if (lbl.Current.Name == searchValue2)
                                                {
                                                    reqlblelem = lbl; // wefound required label 
                                                    break;
                                                }
                                            }
                                            hlp.LogtoTextFile(string.Format("Y cordinate of Label {0} was {1}", searchValue2, reqlblelem.Current.BoundingRectangle.Y));
                                            double lwoffset = reqlblelem.Current.BoundingRectangle.Y - 5;
                                            double hioffset = reqlblelem.Current.BoundingRectangle.Y + 5;
                                            AutomationElement reqcmb = null;
                                            Condition CondPVCombo = null;
                                            switch (searchBy.ToLower())
                                            {
                                                case "classname":
                                                    {
                                                        CondPVCombo = new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ComboBox);
                                                        for (int tk = 0; tk < 15; tk++)
                                                        {
                                                            objcol = rootelem.FindAll(TreeScope.Descendants, CondPVCombo);
                                                            Thread.Sleep(1000);
                                                            if (objcol.Count > 0)
                                                            {
                                                                hlp.LogtoTextFile(string.Format("Upane:  found in {0} Attempt....", tk));
                                                                break;
                                                            }
                                                        }

                                                        foreach (AutomationElement indcmb in objcol)
                                                        {
                                                            if ((indcmb.Current.ClassName.Contains(searchValue)) && (indcmb.Current.BoundingRectangle.Y > lwoffset) && (indcmb.Current.BoundingRectangle.Y < hioffset))
                                                            {
                                                                reqcmb = indcmb;
                                                                break;
                                                            }
                                                        }
                                                        break;
                                                    }
                                            }
                                            comboboxclick(reqcmb, controlValue);
                                        }
                                        else
                                        {
                                            hlp.LogtoTextFile("Upvcombobox : Unable to get label colection from above dialog....");
                                        }
                                    break;

                                }
                            #endregion
                            #region ScrollBar
                            case "scrollbar":
                                {
                                    inmemlist = false; // we need to get visible control collection again
                                     #region MSAAAScrollbar
                                    try
                                    {
                                        if (technologyControl == "Web")
                                        {
                                            if (UIcurrentparent.Exists)
                                            {
                                                HtmlScrollBar ucntl = new HtmlScrollBar(UIcurrentparent);
                                                if (cOperator == "=")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue);
                                                    if (index.Length > 0)
                                                    {
                                                        ucntl.SearchProperties.Add("Instance", index);
                                                    }
                                                }
                                                else if (cOperator == "~")
                                                {
                                                    ucntl.SearchProperties.Add(searchBy, searchValue, PropertyExpressionOperator.Contains);
                                                }

                                                if (controlValue.Length > 0)
                                                {
                                                    //  System.Drawing.Point p = new System.Drawing.Point(ucntl.BoundingRectangle.X,ucntl.BoundingRectangle.Y);
                                                    //   bool isvisible = ucntl.TryGetClickablePoint(out p);
                                                    ScrollSlide(ucntl, controlValue);
                                                    hlp.LogtoTextFile(string.Format("Clicked Button {0}", field));
                                                }
                                            }
                                            else
                                            {
                                                hlp.LogtoTextFile("Parent of Control button was not constructed: Hence  Any Actions on this control are not peformed] ");
                                            }
                                        }
                                    }
                                    catch
                                    {
                                    }
                                     #endregion
                                    break;
                                }
                            #endregion
                        }
                    }
                    #region ActionDefined
                    switch (action.ToLower())
                    {

                        case "dwait":
                            {

                                Playback.Wait(1000);
                                Lowis_Reports_Testing.ObjectLibrary.LowisMainWindow lmain = new Lowis_Reports_Testing.ObjectLibrary.LowisMainWindow();
                                lmain.lowisDwait();
                                Playback.Wait(1000);
                                break;
                            }
                        case "keystroke":
                            {
                                System.Windows.Forms.SendKeys.SendWait(controlValue);
                                break;
                            }

                        case "wait":
                            {
                                Playback.Wait(Int32.Parse(controlValue) * 1000);
                                break;
                            }
                    }
                    #endregion








                }
            }
            catch (Exception ex)
            {
                hlp.LogtoTextFile("generic error in AddData" + ex.Message);
                return;
            }


        }

        public void ClickElement(AutomationElement el, string actiontype)
        {
            double lx = el.Current.BoundingRectangle.Left;
            double lt = el.Current.BoundingRectangle.Top;
            System.Drawing.Point pt;
            System.Windows.Point wpt;

            if (actiontype == "clickoffset")
            {
                pt = new System.Drawing.Point(Convert.ToInt32(lx) + 20, Convert.ToInt32(lt) + 20);
            }
            else if (actiontype == "clickcentre" || actiontype.Length == 0)
            {
                wpt = el.GetClickablePoint();
                lx = wpt.X;
                lt = wpt.Y;
                pt = new System.Drawing.Point(Convert.ToInt32(lx), Convert.ToInt32(lt));
            }
            else
            {
                pt = new System.Drawing.Point(Convert.ToInt32(lx), Convert.ToInt32(lt));
            }
            Mouse.Click(pt);
        }

        public void comboboxclick(AutomationElement el, string _controlValue)
        {
            try
            {
                hlp.LogtoTextFile("Inside combox function=======");
                System.Drawing.Point p2 = new System.Drawing.Point(Convert.ToInt32(el.Current.BoundingRectangle.X), Convert.ToInt32(el.Current.BoundingRectangle.Y));
                Mouse.Click(p2);
                Thread.Sleep(2000);
                WinWindow cmbwin = new WinWindow();
                cmbwin.SearchProperties.Add(WinWindow.PropertyNames.Name, "ComboBox");
                while (cmbwin.Exists == false)
                {
                    hlp.LogtoTextFile("Did not get Combobx UI after clicking on Dropdown Retryuing  again =======");
                    Mouse.Click(p2);
                    Playback.Wait(2000);
                }
                hlp.LogtoTextFile("List obtained Confirmd");
                hlp.LogtoTextFile("Try Construct Element to click using AUelem collection");
                AutomationElement ae = AutomationElement.RootElement;
                Condition cond = new System.Windows.Automation.AndCondition(
                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "ComboBox", PropertyConditionFlags.IgnoreCase)
                    );
                AutomationElement cmbowin = ae.FindFirst(TreeScope.Descendants, cond);
                if (cmbwin != null)
                {
                    Condition cond2 =
                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem);
                    AutomationElementCollection alldataitems = cmbowin.FindAll(TreeScope.Descendants, cond2);
                    if (alldataitems.Count > 0)
                    {
                        hlp.LogtoTextFile("Got collection count =" + alldataitems.Count);
                        System.Diagnostics.Trace.WriteLine("Got collection count =" + alldataitems.Count);
                        foreach (AutomationElement inditem in alldataitems)
                        {
                            if (inditem.Current.Name == _controlValue)
                            {
                                InvokePattern invk = (InvokePattern)inditem.GetCurrentPattern(InvokePattern.Pattern);
                                invk.Invoke();
                                break;
                            }
                        }
                    }
                    else
                    {
                        hlp.LogtoTextFile("no dataimtes was invokved");
                    }
                }
                else
                {
                    hlp.LogtoTextFile("Attempt to reinvoke comboobox faield");
                }
            }
            catch (Exception ex)
            {
                hlp.LogtoTextFile("Exception "+ex.Message.ToString());
            }

        }

        public bool IsColumnPresent(string colname, DataTable testData)
        {
            try
            {
                bool IsColumnPresent = false;
                string colNameString = "";
                for (int ic = 0; ic < testData.Columns.Count; ic++)
                {
                    colNameString = colNameString + testData.Columns[ic].Caption.ToString() + ";";
                }
                if (colNameString.Contains(colname))
                {
                    IsColumnPresent = true;
                }

                return IsColumnPresent;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }
        public void ScrollSlide(UITestControl scrollBar,string numberofclick)
    {
        System.Drawing.Point bottomOfScrollBar = new System.Drawing.Point(scrollBar.Left + (scrollBar.Width / 2), scrollBar.Top + (scrollBar.Height - (scrollBar.Height / 20)));

        Mouse.Move(null, new System.Drawing.Point(scrollBar.Left + (scrollBar.Width / 2), scrollBar.Top + (scrollBar.Height - (scrollBar.Height / 15))));
        Mouse.Move(null, bottomOfScrollBar );
        for (int i = 0; i < Int32.Parse(numberofclick); i++)
        {
            Mouse.DoubleClick(null, bottomOfScrollBar);
        }
    }

    }
}

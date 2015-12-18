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

namespace Lowis_Reports_Testing.StructureSheet
{
    class UIObect
    {

        Helper hlp = new Helper();

        public void AddData(string filename, string testcase)
        {
            DataTable dt1 = hlp.dtFromExcelFile(filename, "Template");
            DataTable dt2 = hlp.dtFromExcelFile(filename, "ExpectedData", "TestCase", testcase);
            UITestControl UIcurrentparent = null;
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                string controlValue = "";
                string parentType = dt1.Rows[i]["ParentType"].ToString();
                string parentSearchBy = dt1.Rows[i]["parentSearchBy"].ToString();
                string parentSearchValue = dt1.Rows[i]["parentSearchValue"].ToString();
                string pTechnology = dt1.Rows[i]["Technology"].ToString();
                string controlType = dt1.Rows[i]["ControlType"].ToString();
                string technologyControl = dt1.Rows[i]["TechnologyControl"].ToString();
                string field = dt1.Rows[i]["Field"].ToString();
                string action= dt1.Rows[i]["Action"].ToString();
                string index = dt1.Rows[i]["Index"].ToString();
                string pindex = dt1.Rows[i]["pindex"].ToString();
                string searchBy = dt1.Rows[i]["SearchBy"].ToString();
                string searchValue = dt1.Rows[i]["SearchValue"].ToString();
                string pOperator = dt1.Rows[i]["pOperator"].ToString();
                string cOperator = dt1.Rows[i]["cOperator"].ToString();
                string resetparent = dt1.Rows[i]["ResetParent"].ToString();
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
                                #endregion 
                                else if (technologyControl == "UIA")
                                {
                                    // to do
                                }
                                #region Webbutton
                                else if (technologyControl == "Web")
                                {
                                    HtmlButton ucntl= new HtmlButton(UIcurrentparent);
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
                                #endregion

                                #region ActionDefined
                                if (action == "dwait")
                                {
                                    Playback.Wait(1000);
                                    Lowis_Reports_Testing.ObjectLibrary.LowisMainWindow lmain = new Lowis_Reports_Testing.ObjectLibrary.LowisMainWindow();
                                    lmain.lowisDwait();
                                    Playback.Wait(1000);
                                }
                                #endregion
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
                                    hlp.LogtoTextFile("Exception frmo Control Type = Image " + "  Field Name : " + field + "Mesage: "+ex.Message);
                                }
                                break;
                            }
                        #endregion
                        #region Dropdown
                        case "dropdown" :
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
                                    if (cOperator == "=" )
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

                                        ucntl.SelectedItem = controlValue;
                                        hlp.LogtoTextFile("Entered Value in ComboBox");
                                    }
                                }
                                break;
                            }
                        #endregion 
                        #region TreeItem
                        case "treeitem" :
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
                                AutomationElement rootelem = AutomationElement.RootElement;
                                AutomationElement paneobject = null;
                                Condition CondSysDateTimePick32 = null;
                                switch(searchBy.ToLower())
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

                                AutomationElementCollection objcol = rootelem.FindAll(TreeScope.Descendants, CondSysDateTimePick32);
                                if (index.Length > 0)
                                {
                                    paneobject = objcol[Int32.Parse(index)];
                                }
                                else
                                {
                                    paneobject = objcol[0];
                                }
                                ClickElement(paneobject,action);
                                if (controlValue.Length > 0)
                                {
                                    Keyboard.SendKeys(controlValue);
                                }
                               
                                }
                                break;
                        #endregion
                    }
                }








            }


        }

        public void ClickElement(AutomationElement el ,string actiontype)
        {
            double lx =   el.Current.BoundingRectangle.Left;
            double lt =   el.Current.BoundingRectangle.Top;
            System.Drawing.Point pt;
            System.Windows.Point wpt;

            if (actiontype == "clickoffset")
            {
                 pt = new System.Drawing.Point(Convert.ToInt32(lx) + 20, Convert.ToInt32(lt) + 20);
            }
            else if (actiontype == "clickcentre")
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
    }
}

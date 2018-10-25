using Microsoft.Office.Tools.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new AutoCopyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class AutoCopyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private ThisAddIn _AddIn;
        IStateSaver currentState = new TextSaver();

        public AutoCopyRibbon(ThisAddIn AddIn)
        {
            Debug.WriteLine("A new ribbon was created");
            currentState.Load();
            _AddIn = AddIn;
        }

        public void EnableAutoCopy(Office.IRibbonControl control, bool isEnabled)
        {
            if (control.Id == "togglebox")
            {
                currentState.IsEnabled = isEnabled;
                currentState.Save();

                _AddIn.Enabled = isEnabled;

                Debug.WriteLine($"Current State: {isEnabled}");

            }

        }

        public bool OnLoaded(Office.IRibbonControl control)
        {
            currentState.Load();
            return currentState.IsEnabled;
        }

        //public void LoadAutoCopy(Office.IRibbonControl control, bool isEnable)
        //{

        //}


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.AutoCopyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {

            

            currentState.Load();
            _AddIn.Enabled = currentState.IsEnabled;

            Debug.WriteLine($"Loaded State: {currentState.IsEnabled}");

            this.ribbon = ribbonUI;

            
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

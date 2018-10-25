using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using ConsoleApp4.Model;
using HtmlAgilityPack;

namespace OutlookAddIn1
{


    public partial class ThisAddIn
    {

        public static AutoCopyRibbon ribbon;
        public bool Enabled { get { return _Enabled; } set { _Enabled = value; } }

        private IStateSaver currentState = new ResourceSaver();
        private bool _Enabled;

        Outlook.MAPIFolder sentFolder;
        Outlook.Items sentItems;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentState.Load();

            sentFolder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
            sentItems = sentFolder.Items;

            sentItems.GetFirst();
            sentItems.ItemAdd += OnItemAdd;

            Enabled = currentState.IsEnabled;
        }
        
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return ribbon ?? (ribbon = new AutoCopyRibbon(this));
        }

        private void OnItemAdd(object Item)
        {
            if (Item is Outlook.MailItem mail)
            {
                
                Outlook.Recipients recips = mail.Recipients;
                string exchangeUserDetails;
                List<string> receivedUserAddressList = new List<string> { };
                
                foreach (Outlook.Recipient recip in recips)
                {
                    Outlook.ExchangeUser exchangeUser = recip.AddressEntry.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        exchangeUserDetails = $"{exchangeUser.Name} <{exchangeUser.Alias}@education.wa.edu.au>";
                    }

                    else
                    {
                        exchangeUserDetails = $"{recip.Address}";
                    }
                    receivedUserAddressList.Add(exchangeUserDetails);
                }


                var HTMLBody = new HtmlDocument();
                HTMLBody.LoadHtml(mail.HTMLBody);

                EmailTransformerModel email = new EmailTransformerModel();

                Debug.WriteLine(email.ConvertTo(HTMLBody));
                            
                string publishedSend = email.ConvertTo(HTMLBody);

                string allReceviedAddresses = string.Join(", ",receivedUserAddressList);
                string MailContentNew = $"From: {mail.SenderName}, {mail.Sender.GetExchangeUser().Name} \nSent : {mail.CreationTime}\nReceived By: {allReceviedAddresses}\nSubject: {mail.Subject}\n\n{publishedSend}";

                Debug.WriteLine("\n" + MailContentNew);

                if(Enabled)
                {
                    Debug.WriteLine("Enabled");
                    Debug.WriteLine(MailContentNew);
                    System.Windows.Forms.Clipboard.SetText(MailContentNew);
                }
            }
        }
        private void GetReceipts()
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

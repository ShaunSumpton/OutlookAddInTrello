using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Manatee.Trello;


namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAddAsync);
        }

        async void items_ItemAddAsync(object Item)
        {
           // string filter = "Approved for Print";

            Outlook.MailItem mail = (Outlook.MailItem)Item;

            TrelloAuthorization.Default.AppKey = "234d8eb40d3f3133b0812df057f7bdc3"; // Trello API key //
            TrelloAuthorization.Default.UserToken = "0e956ba7f0000d7ca7db8504e58a3301d45102e400297f230bfbdda2acc30e1e"; // Trello UserToken //
            string jnum = "";

            ITrelloFactory factory = new TrelloFactory();    // Get Trello board using board ID//
            var board = factory.Board("5db19603e4428377d77963b1");
            await board.Refresh();

            var TDList = factory.List("5db19603e4428377d77963b2");
            await TDList.Refresh();

            // 5db19603e4428377d77963b2 To Start Board ID
            // 5db19603e4428377d77963b3 On Proof Board ID
            // 5db19603e4428377d77963b4 Signed Off 

            // board = TDList.Contains("255705");

            var FoundList = board.Lists.FirstOrDefault(l => l.Name == "Swim Lane");
            var FoundCard = board.Cards.FirstOrDefault(l => l.Name == jnum);

            if (Item != null)
            {
                
                if (mail.Body.ToUpper().Contains("Approved for Print".ToUpper()))
                {
                    MessageBox.Show("Approved for Print");
                }
                else if (mail.Body.ToUpper().Contains("Awaiting Review".ToUpper()))
                {
                    MessageBox.Show("Awaiting Review");
                }
                else if(mail.Body.ToUpper().Contains("Amends".ToUpper()))
                {
                    MessageBox.Show("Amends");
                }
                else
                {
                    // non job mail
                }

            }

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

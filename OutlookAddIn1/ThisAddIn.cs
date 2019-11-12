using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Manatee.Trello;
using IQueryable = Manatee.Trello.IQueryable;


namespace OutlookAddIn1
{
    public partial class ThisAddIn

    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        public List CardList;
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

            bool test = false;
            var cardid = "";
            var LID = "";
            string jnum = "";
            int i = 0;
            var cardindex = "";

            try
            {

                Outlook.MailItem mail = (Outlook.MailItem)Item;

                TrelloAuthorization.Default.AppKey = "234d8eb40d3f3133b0812df057f7bdc3"; // Trello API key //
                TrelloAuthorization.Default.UserToken = "0e956ba7f0000d7ca7db8504e58a3301d45102e400297f230bfbdda2acc30e1e"; // Trello UserToken //


                ITrelloFactory factory = new TrelloFactory();    // Get Trello board using board ID//
                var board = factory.Board("5db19603e4428377d77963b1");
                await board.Refresh();

                var TDList = factory.List("5db19603e4428377d77963b2");
                //var TList = factory.List("");
                await TDList.Refresh();

                var Start = mail.Subject.IndexOf("t").ToString();
                jnum = mail.Subject.Substring(int.Parse(Start) + 1, 7);

                string[] ListID = new string[3];

                // 
                ListID[0] = "5db19603e4428377d77963b2"; //To Start Board ID
                ListID[1] = "5db19603e4428377d77963b3";// On Proof Board ID
                ListID[2] = "5db19603e4428377d77963b4";// Signed Off Board ID 

                //board = TDList.Contains("255705");

                // find Card



                string query = jnum;
                var search = factory.Search(query, 1, SearchModelType.Cards, new IQueryable[] { board });
                await search.Refresh();

                var CardList = search.Cards.ToList();

                foreach (var card in CardList)
                {
                    string tName = card.Name.Substring(0, 6);

                    if (tName == jnum.Trim())
                    {
                        cardid = card.Id;


                    }
                }

                var FoundCard = factory.Card(cardid);
                string FoundListid = FoundCard.List.Id;
                var fromlist = factory.List(FoundListid);
                Person p1 = new Person();
                p1.Name = "Shaun";


                //var FoundList = board.Lists.FirstOrDefault(l => l.Name == "Swim Lane");



                if (Item != null)
                {

                    if (mail.Body.ToUpper().Contains("Approved for Print".ToUpper()))

                    {
                        //var ToList = factory.List("5db19603e4428377d77963b4");
                        var ToList = board.Lists.FirstOrDefault(l => l.Name == "Signed Off");
                        FoundCard.List = ToList;
                        // from on proof


                        //MessageBox.Show("Approved for Print");
                    }
                    else if (mail.Body.ToUpper().Contains("Awaiting Review".ToUpper()))

                    {
                        //var ToList = factory.List("5db19603e4428377d77963b3");
                        var ToList = board.Lists.FirstOrDefault(l => l.Name == "On Proof");

                        FoundCard.List = ToList;

                        // from in progress or to start

                        // MessageBox.Show("Awaiting Review");
                    }
                    else if (mail.Body.ToUpper().Contains("Amends".ToUpper()))
                    {
                        var ToList = factory.List("5dc9442eb245e60a39b3d4a7");
                        FoundCard.List = ToList;

                        // from on proof
                        //MessageBox.Show("Amends");
                    }
                    else
                    {
                        // non job mail
                    }

                }

            }
            catch(Exception e)
            {
                //MessageBox.Show(e.Message);
            }


}
     
        class Person
        {
            private string _name;
           public string Name
            {
                get => _name;
                set => _name = value;

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

    }
}
        #endregion
    



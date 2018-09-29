using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
// using Microsoft.Office.Interop.Outlook.Application;

namespace ReadEmail
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        Microsoft.Office.Interop.Outlook.NameSpace outlookNamespace;
        MAPIFolder inboxFolder;
        Items mailItems;
        string[] bodyLines = new string[100];
        static Microsoft.Office.Interop.Outlook.MAPIFolder thisInBox;
        public Form1()
        {
            InitializeComponent();
        }
        private void ReadAccounts()
        {
            Microsoft.Office.Interop.Outlook.NameSpace outlookNamespace = null;
            Microsoft.Office.Interop.Outlook.Accounts accounts = null;
            Microsoft.Office.Interop.Outlook.Account account = null;
            string accountList = string.Empty;

            try
            {
               // ns = OutlookApp.Session;
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                accounts = outlookNamespace.Accounts;
                for (int i = 1; i <= accounts.Count; i++)
                {
                    account = accounts[i];
                    accountList += String.Format("{0} - {1}{2}",
                        account.UserName,
                        account.SmtpAddress,
                        Environment.NewLine);
                    //if (account != null)
                    //    Marshal.ReleaseComObject(account);
                }
                MessageBox.Show(accountList);
            }
            finally
            {
                //if (accounts != null)
                //    Marshal.ReleaseComObject(accounts);
                //if (ns != null)
                //    Marshal.ReleaseComObject(ns);
            }
        }
        private static void ReadMailItems()
        {
            Microsoft.Office.Interop.Outlook.Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;

            try
            {
                outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                string folderName = "TestFolder";
                thisInBox = (Microsoft.Office.Interop.Outlook.MAPIFolder)
                    outlookApplication.ActiveExplorer().Session.GetDefaultFolder
                    (Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

               // mailItems = inboxFolder.Items;
                outlookApplication.ActiveExplorer().CurrentFolder = thisInBox.
                    Folders[folderName];
                mailItems = thisInBox.Items;

                foreach (MailItem item in mailItems)
                {
                    var stringBuilder = new StringBuilder();
                    stringBuilder.AppendLine("From: " + item.SenderEmailAddress);
                    stringBuilder.AppendLine("To: " + item.To);
                    stringBuilder.AppendLine("CC: " + item.CC);
                    stringBuilder.AppendLine("Received  " + item.ReceivedTime);
                    stringBuilder.AppendLine("");
                    stringBuilder.AppendLine("Subject: " + item.Subject);
                    stringBuilder.AppendLine(item.Body);

                    Console.WriteLine(stringBuilder);
                  //  ReleaseComObject(item);
                }
            }
            catch { }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }
        private void SetCurrentFolder()
        {
            Microsoft.Office.Interop.Outlook.Folders folders = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder folder = null;
            string folderName = "Inbox";
            thisInBox = (Microsoft.Office.Interop.Outlook.MAPIFolder)
                this.outlookApplication.ActiveExplorer().Session.GetDefaultFolder
                (Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
           // folders = rootFolder.Folders;
            folders = thisInBox.Folders;

           
                folder = folders[1];
            folderName = folder.Name;
                try
            {
                this.outlookApplication.ActiveExplorer().CurrentFolder = thisInBox.Folders[folderName];
                this.outlookApplication.ActiveExplorer().CurrentFolder.Display();
            }
            catch
            {
                MessageBox.Show("There is no folder named " + folderName +
                    ".", "Find Folder Name");
            }
        }
        private void SearchInBox()
        {
            Microsoft.Office.Interop.Outlook.Folders folders = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder folder = null;
            string folderName = "Inbox";
            //Microsoft.Office.Interop.Outlook.MAPIFolder inbox = this.outlookApplication.ActiveExplorer().Session.
            //    GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            Microsoft.Office.Interop.Outlook.Items items = thisInBox.Items;
            Microsoft.Office.Interop.Outlook.MailItem mailItem = null;
            object folderItem;
            folders = thisInBox.Folders;

            folder = folders[1];
            folderName = folder.Name;
            outlookApplication.ActiveExplorer().CurrentFolder = thisInBox.
                    Folders[folderName];

            string subjectName = string.Empty, thisSubject = string.Empty, thisBody = string.Empty;
            
            int MailCount = outlookApplication.ActiveExplorer().CurrentFolder.Items.Count;
            folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items[1];
          
            int mailRead = 0;
            while (folderItem != null && mailRead < MailCount)
            {
                mailRead++;
                mailItem = folderItem as Microsoft.Office.Interop.Outlook.MailItem;
                if (mailItem != null)
                {
                    thisSubject = mailItem.Subject;
                    if (thisSubject.Contains("Daily Report:"))
                    {
                        thisBody = mailItem.Body;
                        ParseBody(thisBody);
                        subjectName += "\n" + mailItem.Subject;
                    }
                   
                }
                //folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items.GetLast();
                //folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items.GetNext();
                folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items[mailRead];
            }
            subjectName = " The following e-mail messages were found: " +
                subjectName;
            MessageBox.Show(subjectName);
        }
        void ParseBody(string thisBody)
        {
            int start = 0, end = 0;
            for(int i = 0; i < 100; i++)
            {
                end = thisBody.IndexOf("\n", start);
                //end = 20;
                if(end > start) bodyLines[i] = thisBody.Substring(start, end - start);
                start = end + 1;
            }
        }
        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                ReleaseComObject(obj);
                obj = null;
            }
        }
        private static void GetListOfStores()
        {
            Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            
            Microsoft.Office.Interop.Outlook.NameSpace outlookNamespace = null;
            Microsoft.Office.Interop.Outlook.Stores stores = null;
            Microsoft.Office.Interop.Outlook.Store store = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder rootFolder = null;
            Microsoft.Office.Interop.Outlook.Folders folders = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder folder = null;
            Microsoft.Office.Interop.Outlook.Accounts accounts = null;
            Microsoft.Office.Interop.Outlook.Account account = null;
            string folderList = string.Empty;
            int storeCount;

            try
            {
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                accounts = outlookNamespace.Accounts;
               // account = accounts[1];
                stores = outlookNamespace.Stores;
                for (int j = 1; j < 5; j++)
                {
                    account = accounts[j];
                    store = stores[j];
                    storeCount = stores.Count;
                    
                    rootFolder = store.GetRootFolder();
                    folders = rootFolder.Folders;
                    //folderList += String.Format("{0} - {1}{2}",
                    //    account.UserName,
                    //    account.SmtpAddress,
                    //    Environment.NewLine);
                    for (int i = 1; i < folders.Count; i++)
                    {
                        folder = folders[i];
                        folderList += folder.Name + Environment.NewLine;
                        //if (folder != null)
                        //    ReleaseComObject(folder);
                    }
                    MessageBox.Show(folderList);
                    folderList = string.Empty;
                }
            }
            finally
            {
                //if (folders != null)
                //    ReleaseComObject(folders);
                //if (folders != null)
                //    ReleaseComObject(folders);
                //if (rootFolder != null)
                //    ReleaseComObject(rootFolder);
                //if (store != null)
                //    ReleaseComObject(store);
                //if (stores != null)
                //    ReleaseComObject(stores);
                //if (outlookNamespace != null)
                //    ReleaseComObject(outlookNamespace);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ReadMailItems();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetCurrentFolder();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SearchInBox();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ReadAccounts();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            GetListOfStores();
        }
    }

}

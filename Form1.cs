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
        static Microsoft.Office.Interop.Outlook.MAPIFolder thisInBox;
        public Form1()
        {
            InitializeComponent();
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
            string folderName = "TestFolder";
            thisInBox = (Microsoft.Office.Interop.Outlook.MAPIFolder)
                this.outlookApplication.ActiveExplorer().Session.GetDefaultFolder
                (Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            try
            {
                this.outlookApplication.ActiveExplorer().CurrentFolder = thisInBox.
                    Folders[folderName];
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
            string folderName = "TestFolder";
            //Microsoft.Office.Interop.Outlook.MAPIFolder inbox = this.outlookApplication.ActiveExplorer().Session.
            //    GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            Microsoft.Office.Interop.Outlook.Items items = thisInBox.Items;
            Microsoft.Office.Interop.Outlook.MailItem mailItem = null;
            object folderItem;
            outlookApplication.ActiveExplorer().CurrentFolder = thisInBox.
                    Folders[folderName];

            string subjectName = string.Empty;
            //string filter = "[Subject] > 's' And [Subject] <'u'";
            //filter = "[Subject] > 's'";
           // folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items.Find(filter);
            int MailCount = outlookApplication.ActiveExplorer().CurrentFolder.Items.Count;
            folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items.GetFirst();
            while (folderItem != null)
            {
                mailItem = folderItem as Microsoft.Office.Interop.Outlook.MailItem;
                if (mailItem != null)
                {
                    subjectName += "\n" + mailItem.Subject;
                }
                folderItem = outlookApplication.ActiveExplorer().CurrentFolder.Items.GetNext();
            }
            subjectName = " The following e-mail messages were found: " +
                subjectName;
            MessageBox.Show(subjectName);
        }
        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                ReleaseComObject(obj);
                obj = null;
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
    }

}

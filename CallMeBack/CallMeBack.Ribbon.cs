using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;

namespace CallMeBack
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			var outlookObj = new Application();
			var tmp = outlookObj.CreateItem(OlIte‌​mType.olAppointmentItem);
			MAPIFolder calendar = outlookObj.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
			Explorer explorer = outlookObj.ActiveExplorer();
			Folder folder = explorer.CurrentFolder as Folder;
			View view = explorer.CurrentView as View;
			MailItem mail = view as MailItem;
			ShowForm();
        }

            
        

        private Form1 form1 = null;

        private void ShowForm()
        {
            if (form1 == null)
            {                
                form1 = new Form1(Globals.ThisAddIn.Application);                
            }
            form1.ShowDialog();
        }
        
    }
}

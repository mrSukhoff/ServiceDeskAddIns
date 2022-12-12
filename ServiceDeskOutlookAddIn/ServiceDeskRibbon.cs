using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ServiceDeskOutlookAddIn
{
    public partial class ServiceDeskRibbon
    {
        private void ServiceDeskRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void CreateClaimButton_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Помогите!", "Нужна помощь!");
            //var application = (Microsoft.Office.Interop.Outlook.Application)sender;
            //var cntrl = new UserControl1();
            //ThisAddIn.CreateSendItem(application);
        }
        

    }
}

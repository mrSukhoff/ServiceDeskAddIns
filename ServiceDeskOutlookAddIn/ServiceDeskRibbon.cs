using Microsoft.Office.Tools.Ribbon;

namespace ServiceDeskOutlookAddIn
{
    public partial class ServiceDeskRibbon
    {
        private void ServiceDeskRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void CreateClaimButton_Click(object sender, RibbonControlEventArgs e)
        {
            var cntrl = new RequestWindow();
            cntrl.Show();
        }
    }
}

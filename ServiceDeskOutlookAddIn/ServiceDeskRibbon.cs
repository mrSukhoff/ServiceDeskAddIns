﻿using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
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
            MessageBox.Show("Помогите!","Нужна помощь!");
        }
    }
}
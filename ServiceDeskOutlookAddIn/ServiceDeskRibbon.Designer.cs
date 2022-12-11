namespace ServiceDeskOutlookAddIn
{
    partial class ServiceDeskRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ServiceDeskRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ServiceDeskGroup = this.Factory.CreateRibbonGroup();
            this.CreateClaimButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ServiceDeskGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.ServiceDeskGroup);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // ServiceDeskGroup
            // 
            this.ServiceDeskGroup.Items.Add(this.CreateClaimButton);
            this.ServiceDeskGroup.Label = "ИТ-поддержка";
            this.ServiceDeskGroup.Name = "ServiceDeskGroup";
            this.ServiceDeskGroup.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupMailNew");
            // 
            // CreateClaimButton
            // 
            this.CreateClaimButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateClaimButton.Image = global::ServiceDeskOutlookAddIn.Properties.Resources.Ghost;
            this.CreateClaimButton.Label = "Создать обращение";
            this.CreateClaimButton.Name = "CreateClaimButton";
            this.CreateClaimButton.ShowImage = true;
            this.CreateClaimButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateClaimButton_Click);
            // 
            // ServiceDeskRibbon
            // 
            this.Name = "ServiceDeskRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ServiceDeskRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ServiceDeskGroup.ResumeLayout(false);
            this.ServiceDeskGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ServiceDeskGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateClaimButton;
    }

    partial class ThisRibbonCollection
    {
        internal ServiceDeskRibbon ServiceDeskRibbon
        {
            get { return this.GetRibbon<ServiceDeskRibbon>(); }
        }
    }
}

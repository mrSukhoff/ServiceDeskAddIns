using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace ServiceDeskOutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Примечание. Outlook больше не выдает это событие. Если имеется код, который 
            //    должно выполняться при завершении работы Outlook, см. статью на странице https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        public void CreateSendItem(Outlook.Application Application)
        {
            Outlook.MailItem mail = null;
            Outlook.Recipients mailRecipients = null;
            Outlook.Recipient mailRecipient = null;
            try
            {
                mail = Application.CreateItem(Outlook.OlItemType.olMailItem)
                    as Outlook.MailItem;
                mail.Subject = "Тестовое обращение";
                mailRecipients = mail.Recipients;
                mailRecipient = mailRecipients.Add("m.sukharev@pharmasyntez.com");
                mailRecipient.Resolve();
                if (mailRecipient.Resolved)
                {
                    mail.Send();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(
                        "There is no such record in your address book.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message,
                     "An exception is occured in the code of add-in.");
            }
            finally
            {
                if (mailRecipient != null) Marshal.ReleaseComObject(mailRecipient);
                if (mailRecipients != null) Marshal.ReleaseComObject(mailRecipients);
                if (mail != null) Marshal.ReleaseComObject(mail);
            }
        }
    }
}

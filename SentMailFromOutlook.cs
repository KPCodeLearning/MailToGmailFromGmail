using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office;

          try
            {
                Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mailItem.Subject = "This is the subject";
                mailItem.To = "karmanjayverma@gmail.com";
                mailItem.Body = "This is the message.";
                //mailItem.Attachments.Add(logPath);//logPath is a string holding path to the log.txt file
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                //mailItem.Display(false);
                mailItem.Send();
            }
            catch (Exception ex)
            {
                var msg = ex.Message;
            }

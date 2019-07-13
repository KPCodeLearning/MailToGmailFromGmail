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
namespace HealthCheckWinApp
{
    public partial class Form1 : Form
    {
        //List<urlClass> _lt = new List<urlClass>();

        //Dictionary<string,bool> _lt = new Dictionary<string, bool>();
        List<urlClass> _lt = new List<urlClass>()
            {
                new urlClass {WebURL = "ABCD", Status = false},
                new urlClass {WebURL = "ABCD1", Status = true}
            };
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
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
        }

        private void check_Click(object sender, EventArgs e)
        {
            var bindingList = new BindingList<urlClass>(_lt);
            var source = new BindingSource(bindingList, null);
            dataGridView1.DataSource = source;
            dataGridView1.Visible = true;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

    }

    public class urlClass 
    {
        public string WebURL { get; set; }
        public bool Status { get; set; }
    }
}

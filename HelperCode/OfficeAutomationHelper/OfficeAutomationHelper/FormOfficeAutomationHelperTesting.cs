using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OfficeAutomationHelperTesting
{
    public partial class FormOfficeAutomationHelperTesting : Form
    {
        OfficeAutomationHelper officeAutomationHelper = new OfficeAutomationHelper();
        Form formOfficeAutomationHelperTesting;

        public FormOfficeAutomationHelperTesting()
        {
            InitializeComponent();
            formOfficeAutomationHelperTesting = this;

            //Create email setup
            textBoxCreateMailTo.Text = Environment.UserName + "@ag.org";
            webBrowserCreateMail.DocumentText = "You would use dynamic HTML here in code";

            textBoxAppointmentTo.Text = Environment.UserName + "@ag.org";
            DateTimePickerExtendedAppointmentDate.Value = DateTime.Today.AddDays(1);

            comboBoxDataGridToExcelColor.DataSource = Enum.GetNames(typeof(OfficeAutomationHelper.ExcelDrawColors));

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Acct",typeof(int)); 
            dataTable.Columns.Add("Last",typeof(string)); 
            dataTable.Columns.Add("First",typeof(string)); 
            dataTable.Columns.Add("City",typeof(string)); 
            dataTable.Columns.Add("State",typeof(string));
            dataTable.Columns.Add("Amount", typeof(double));
            dataTable.Columns.Add("Date", typeof(DateTime)); 
            dataTable.Rows.Add(350299, "Beam", "Jon", "Fair Grove", "MO",123.45, DateTime.Now.AddMonths(1).AddDays(2).ToString("d"));
            dataTable.Rows.Add(398860, "Downs", "Larry", "Nixa", "MO",234.56, DateTime.Now.AddMonths(3).AddDays(5).ToString("d"));
            dataTable.Rows.Add(809323, "Yarbrough", "Darren", "Willard", "MO", 345.67, DateTime.Now.AddMonths(4).AddDays(8).ToString("d"));
            dataTable.Rows.Add(809323, "Balsters", "Brad", "Springfield", "MO",456.78, DateTime.Now.AddMonths(5).AddDays(3).ToString("d"));
            dataGridViewToExcel.DataSource = dataTable;
            dataGridViewToExcel.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridViewToExcel.RowHeadersVisible = false;
            dataGridViewToExcel.Rows[0].Height = 4;
            dataGridViewToExcel.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewToExcel.RowsDefaultCellStyle.BackColor = Color.White;
            dataGridViewToExcel.AlternatingRowsDefaultCellStyle.BackColor = Color.Wheat;
            dataGridViewToExcel.Columns["Acct"].DefaultCellStyle.Format = "#";
            dataGridViewToExcel.Columns["Amount"].DefaultCellStyle.Format = "c";
            dataGridViewToExcel.Columns["Date"].DefaultCellStyle.Format = "d";
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            comboBoxDataGridToExcelColor.SelectedItem = OfficeAutomationHelper.ExcelDrawColors.Blue;
        }

        private void ToolStripButtonSpellCheck_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.checkBoxGetActiveControlInForm.Checked == true)
                {
                    officeAutomationHelper.SpellCheckActiveControlUsingWord(this);
                }
                else
                {
                    this.textBoxSpellCheckActiveControl.Text = officeAutomationHelper.SpellCheckUsingWord(textBoxSpellCheckActiveControl.Text);
                }
            }
            catch (Exception ex)
            {
                //handle these exceptions in your code
                MessageBox.Show("Error or Warning: " + ex.Message);
            }
        }


        private void checkBoxCreateMailHTML_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCreateMailHTML.Checked == true)
            {
                this.webBrowserCreateMail.DocumentText = "<p>Body uses <FONT COLOR=\"#FF0000\"><b>HTML</b></FONT>.</p>";
            }
            else
            {
                this.webBrowserCreateMail.DocumentText = "Body uses plain text";
            }
        }

        private void ToolStripButtonEmail_Click(object sender, EventArgs e)
        {
            bool DisplayFirst = checkBoxCreateMailDisplayFirst.Checked;
            bool Important = checkBoxCreateMailImportant.Checked;
            bool CreateHtml = checkBoxCreateMailHTML.Checked;
            string To = textBoxCreateMailTo.Text;
	        string Cc = textBoxCreateMailCc.Text;
	        string Bcc = textBoxCreateMailBcc.Text;  
   	        string Subject = textBoxCreateMailSubject.Text;
            string Body = webBrowserCreateMail.DocumentText;
            List<string> Attachments = ListBoxCreateMailAttachments.Items.Cast<String>().ToList();
            officeAutomationHelper.OutlookNewMail(DisplayFirst,Important, To, Cc, Bcc, Subject, Body, CreateHtml, Attachments);
        }

        private void ToolStripButtonCreateEditAppointment_Click(object sender, EventArgs e)
        {
            if (!DateTimePickerExtendedAppointmentDate.IsDate()) 
            {
                MessageBox.Show("Enter a valid date");
            }
            else
            {
                DateTime date = (DateTime)DateTimePickerExtendedAppointmentDate.Value;
                DateTime time = (DateTime)DateTimePickerIncTimeAppointmentTime.Value;
                List<string> Attachments = ListBoxAppointments.Items.Cast<String>().ToList();

                OfficeAutomationHelper.AppointmentStructure aS = new OfficeAutomationHelper.AppointmentStructure();
                aS.EntryID = labelAppointmentEntryID.Text;
                aS.Subject = textBoxAppointmentSubject.Text;
                aS.Location = textBoxAppointmentLocation.Text;
                aS.Start = new DateTime(date.Year,date.Month,date.Day, time.Hour,time.Minute,0);
                aS.End = aS.Start.AddMinutes((double)NumericUpDownAppointmentDuration.Value); 
                aS.ReminderMinutes = NumericUpDownAppointmentReminder.Value; 
                aS.RequiredAttendees = textBoxAppointmentTo.Text;
                aS.OptionalAttendees = textBoxAppointmentCc.Text;
                aS.Resources = textBoxAppointmentLocation.Text;
                aS.Body = textBoxAppointmentBody.Text; 
                aS.ListAttachments = Attachments;
                aS.HighImportance = checkBoxAppointmentImportant.Checked;
                aS.ResponseRequested = checkBoxAppointmentResponseRequested.Checked;

                if (aS.EntryID.Length == 0)
                {
                    aS = officeAutomationHelper.OutlookNewAppointment(aS);
                }
                else
                {
                    aS = officeAutomationHelper.OutlookEditAppointment(true, true, aS);  
                }

                if (aS.EntryID != null && aS.EntryID.Length > 0)  //reload from outlook appointment if it was saved
                {
                    labelAppointmentEntryID.Text = aS.EntryID;
                    textBoxAppointmentSubject.Text = aS.Subject;
                    textBoxAppointmentLocation.Text = aS.Location;
                    DateTimePickerExtendedAppointmentDate.Value = aS.Start;
                    DateTimePickerIncTimeAppointmentTime.Value = aS.Start;
                    NumericUpDownAppointmentDuration.Value = aS.End.Subtract(aS.Start).Minutes;
                    NumericUpDownAppointmentReminder.Value = aS.ReminderMinutes;
                    textBoxAppointmentTo.Text = aS.RequiredAttendees;
                    textBoxAppointmentCc.Text = aS.OptionalAttendees;
                    textBoxAppointmentLocation.Text = aS.Resources;
                    textBoxAppointmentBody.Text = aS.Body;
                    //ListBoxAppointments.Items.Clear();
                    //ListBoxAppointments.Items.AddRange(aS.ListAttachments.ToArray());  //at this point, they're attached
                    checkBoxAppointmentImportant.Checked = aS.HighImportance;
                    checkBoxAppointmentResponseRequested.Checked = aS.ResponseRequested;
                }
            }
        }

        private void ToolStripButtonExportToExcel_Click(object sender, EventArgs e)
        {
            OfficeAutomationHelper.ExcelDrawColors excelDrawColors = OfficeAutomationHelper.ExcelDrawColors.None;
            excelDrawColors = (OfficeAutomationHelper.ExcelDrawColors)Enum.Parse(typeof(OfficeAutomationHelper.ExcelDrawColors), comboBoxDataGridToExcelColor.SelectedValue.ToString());
            officeAutomationHelper.DataGridViewToExcel(dataGridViewToExcel, checkBoxDataGridToExcelSelectedOnly.Checked, checkBoxDataGridViewToExcelShowWhileLoading.Checked, excelDrawColors);  
        }

        private void buttonSpellCheckHelp_Click(object sender, EventArgs e)
        {
            string message = "This is kind of handy because the user can just click the button to spellcheck text on the form"
                           + "Currently, I only have this checking Textbox and RickTextbox controls, but any control with a "
                           + "Text property should work.";
            MessageBox.Show(message);
        }

        private void buttonCreateMailHelp_Click(object sender, EventArgs e)
        {
            string message = "This is kind of handy because the user may have entered information into a form. Rather than the"
               + "user having to retype the information, we can grab the information and pre-format it for them.  I currently "
               + "use this for users to communicate with missionaries, name and address, and contributor services.  As a simple"
               + "scenario, if they lookup somebody in NA and they're not there, all the information and a request is set in the"
               + "body, the appropriate subject is set, and NA.ag.org is put in the to line.";
            MessageBox.Show(message);
        }

        private void buttonAppointmentHelp_Click(object sender, EventArgs e)
        {
            string message = "This is used in the AGUSM meeting application. The data is stored about the meeting.  When the"
                 + "appointment is made, the EntryID from the Outlook Appointment is stored in the database with the meeting."
                 + "If they need to edit the time, the Outlook Appointment is brought up with a click of a button with any "
                 + "edits that need to be applied. For example, if somebody is added, or the time is changed.";
            MessageBox.Show(message);
        }

        private void buttonDataGridToExcelHelp_Click(object sender, EventArgs e)
        {
            string message = "This is a general export from the datagridview.";
            MessageBox.Show(message);
        }

    }
}

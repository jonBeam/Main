using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;

//needs office developer tools modify Visual Studio
using Word = Microsoft.Office.Interop.Word;       // add a reference to Microsoft.Office.Interop.Word
using Excel = Microsoft.Office.Interop.Excel;     // add a reference to Microsoft.Office.Interop.Excel
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing; // add a reference to Microsoft.Office.Interop.Outlook

/// <summary>
/// This class does NOT require a reference to an Microsoft Office COM objects
/// </summary>
public class OfficeAutomationHelper : IDisposable
{
    private bool _IsDisposed = false;  //used for IDisposable interface implementation

    public enum OfficeApplicationType : byte {None = 0, Word = 1, Outlook = 2, Excel = 3}
    public enum ExcelDrawColors : byte { None = 0, Blue = 1, Orange = 2 } 

    private static Word._Application _WordApplication;
    private static Outlook._Application _OutlookApplication;
    private static Excel._Application _ExcelApplication;

    private static Excel._Workbook  _ExcelWorkbook;
    private static Excel._Worksheet _ExcelWorksheet;

    /// <summary>
    /// Structure for passing parameters in and out of the OutlookNewAppointment and OutlookEditAppointment method. 
    /// </summary>
    public struct AppointmentStructure  //used to populate the appointment
    {
        public string EntryID;
        public string Subject;
        public string Location;
        public DateTime Start;
        public DateTime End;
        public decimal ReminderMinutes;
        public string RequiredAttendees;
        public string OptionalAttendees;
        public string Resources;
        public string Body;
        public List<string> ListAttachments;
        public bool HighImportance;
        public bool ResponseRequested;
    }

    private const string _EntryID = "EntryID";  //this can be added to the Outlook Appointment to find the specific Entry later

    private string[] MSWordSpellingDialogWindowTitleDefault = new string[] {"Spelling: English (U.S.)","Spelling: English (United States)"};

    /// <summary>
    /// Will invoke Word spell checker dialog on the active control. 
    /// </summary>
    public string SpellCheckActiveControlUsingWord(Form form)
    {
        string SpellCheckTextIn = string.Empty;
        string SpellCheckTextOut = string.Empty;
        string Message = string.Empty;

        if (form.ActiveControl != null && form.ActiveControl != null)
        {
            if (object.ReferenceEquals(form.ActiveControl.GetType(), typeof(TextBox)) | object.ReferenceEquals(form.ActiveControl.GetType(), typeof(RichTextBox)))
            {
                SpellCheckTextIn = form.ActiveControl.Text.Trim();
                if (SpellCheckTextIn.Length > 0)
                {
                    SpellCheckTextOut = SpellCheckUsingWord(SpellCheckTextIn);
                    if (SpellCheckTextOut.Trim() == SpellCheckTextIn.Trim())
                    {
                        throw (new Exception("Spell check complete with no changes."));
                    }
                    else
                    {
                        form.ActiveControl.Text = SpellCheckTextOut.Trim();
                    }
                }
                else
                {
                    throw (new Exception("No text to spell check"));
                }
            }
            else
            {
                throw (new Exception("Please select a text entry control before running spell check."));
            }
        }
        else
        {
            throw (new Exception("Please select a control before running spell check."));
        }
        return Message;
    }

    /// <summary>
    /// Will invoke Word spell checker and pass to it the string TextToSpellCheck.   
    /// The string is passed back with any changes from the Word Spell checker.
    /// If the Word Spell check window title is different than the default, pass in the window title to be found.
    /// </summary>
    public string SpellCheckUsingWord(string TextToSpellCheck, string[] MSWordSpellingDialogWindowTitle)
    {
        string SpellCheckTextOut = string.Empty;

        Word._Document document = null;
        Word.Range range = null;

        if (_WordApplication == null)
        {
            _WordApplication = new Word.Application();
        }
        _WordApplication.Visible = false;

        try
        {
            Thread.Sleep(50);  //I don't know why I did this
            if (TextToSpellCheck.Length > 0)
            {
                object emptyItem = System.Reflection.Missing.Value;
                document = _WordApplication.Documents.Add(emptyItem, emptyItem, emptyItem, false);
                range = document.Range(0, 0);
                range.Text = TextToSpellCheck;

                ClassDisplayWindowUsingThread myClassDisplayWindow = new ClassDisplayWindowUsingThread();
                myClassDisplayWindow.StartThreadSpellingWindowOnTop(MSWordSpellingDialogWindowTitle);

                document.CheckSpelling();  //Activate The Spell Checker  //document.CheckGrammar() 
                SpellCheckTextOut = document.Range(0, document.Characters.Count - 1).ToString();
                SpellCheckTextOut = SpellCheckTextOut.Replace("\r\n", string.Empty);
                SpellCheckTextOut = SpellCheckTextOut.Replace("\r", string.Empty);
            }
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception e)
        {
            throw (e);
        }
        finally
        {
            range = null;
            object saveOptionsObject = Word.WdSaveOptions.wdDoNotSaveChanges;
            document.Close(ref saveOptionsObject);  //new interop will properly garbage collect
        }
        return SpellCheckTextOut;
    }

    /// <summary>
    /// Will invoke Word spell checker and pass to it the string TextToSpellCheck.   
    /// The string is passed back with any changes from the Word Spell checker.
    /// </summary>
    public string SpellCheckUsingWord(string TextToSpellCheck)
    {
        string SpellCheckTextOut = string.Empty;

        Word._Document document = null;
        Word.Range range = null;

        if (_WordApplication == null)
        {
            _WordApplication = new Word.Application();
            _WordApplication.Visible = false;
        }
        try
        {
            Thread.Sleep(50);
            if (TextToSpellCheck.Length > 0)
            {
                object emptyItem = System.Reflection.Missing.Value;
                document = _WordApplication.Documents.Add(emptyItem, emptyItem, emptyItem, false);
                int endRange = document.Characters.Count - 1;
                SpellCheckTextOut = document.Range(0, endRange).ToString();

                range = document.Range(0, 0);
                range.Text = TextToSpellCheck;

                ClassDisplayWindowUsingThread myClassDisplayWindow = new ClassDisplayWindowUsingThread();
                myClassDisplayWindow.StartThreadSpellingWindowOnTop(MSWordSpellingDialogWindowTitleDefault);

                document.CheckSpelling();
                SpellCheckTextOut = document.Range(0, document.Characters.Count - 1).Text.ToString();
                SpellCheckTextOut = SpellCheckTextOut.Replace("\r\n", string.Empty);
                SpellCheckTextOut = SpellCheckTextOut.Replace("\r", string.Empty);
            }
        }
        catch (COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception e)
        {
            throw (e);
        }
        finally
        {
            range = null;
            object saveOptionsObject = Word.WdSaveOptions.wdDoNotSaveChanges;
            document.Close(ref saveOptionsObject);
        }
        return SpellCheckTextOut;
    }

    /// <summary>
    /// Create a new Outlook email.   
    /// </summary>
    public bool OutlookNewMail(bool DisplayFirst, bool Important, string To, string Cc, string Bcc, string Subject, string Body, bool IsHTMLBody, List<string> ListOfAttachments)
    {
        Outlook._MailItem mailItem = null;
        bool returnSuccess = false;
        try
        {
            if (_OutlookApplication == null)
            {
                if (Process.GetProcessesByName("OUTLOOK").Length > 0)
                {
                    //long ProcessID = Process.GetProcessesByName("OUTLOOK")[0].Id;
                    //If running this code in visual studio as administrator, you must also run Outlook as administrator. 

                    _OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                }
                else
                {
                    _OutlookApplication = new Outlook.Application();
                }
            }
            mailItem = _OutlookApplication.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mailItem.Importance = (Important ? Outlook.OlImportance.olImportanceHigh : Outlook.OlImportance.olImportanceNormal);
            mailItem.To = To;
            mailItem.CC = Cc;
            mailItem.BCC = Bcc;
            mailItem.Subject = Subject;
            if (IsHTMLBody)
            {
                mailItem.HTMLBody = Body.TrimStart();
            }
            else
            {
                mailItem.Body = Body.TrimStart();
            }
            if (ListOfAttachments != null && ListOfAttachments.Count > 0)
            {
                string PathFile = null;
                foreach (string PathFile_loopVariable in ListOfAttachments)
                {
                    PathFile = PathFile_loopVariable;
                    if (File.Exists(PathFile))
                    {
                        mailItem.Attachments.Add(PathFile);
                    }
                }
            }
            if (DisplayFirst)
            {
                mailItem.Display();
            }
            else if (Subject.Length > 0)
            {
               mailItem.Send();
            }
            returnSuccess = true;
        }
        catch (COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
        }
        return returnSuccess;
    }


    /// <summary>
    /// Create a new Outlook appointment. The appointment's EntityID is passed back so it can be edited later, if needed.
    /// An Appointment structure must be defined and filled before passing to the appointment.
    /// </summary>
    public AppointmentStructure OutlookNewAppointment(AppointmentStructure OutlookAppointmentStructure)
    {
        Outlook._AppointmentItem appointment = null;

        try
        {
            if (_OutlookApplication == null)
            {
                _OutlookApplication = new Outlook.Application();
            }

            appointment = _OutlookApplication.CreateItem(Outlook.OlItemType.olAppointmentItem);

            //this sets the values from the passed in structure
            OutlookAppointmentSetData(ref appointment, OutlookAppointmentStructure);

            appointment.Display(true);

            if (appointment.EntryID != null)
            {
                Outlook.UserProperty myOutlookUserProperty = null;
                myOutlookUserProperty = appointment.UserProperties.Add(_EntryID, Outlook.OlUserPropertyType.olText);  //add property to appt
                myOutlookUserProperty.Value = appointment.EntryID;  //get the value of the property back out
                OutlookAppointmentStructure.EntryID = myOutlookUserProperty.Value;  //get the value of the property back out

                //get the data back out because the user may have changed something. and we don't have the EntryID
                OutlookAppointmentStructure = OutlookAppointmentStructureFromAppointment(appointment);
            }
        }
        catch (COMException ex)
        {
            if (ex.Message.Contains("dialog box is open"))
            {
                MessageBox.Show("Dialog box is open.  Please save or close." );
            }
            else
            {
              throw (new Exception("COM Exception: " + ex.Message));
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
        }

        return (OutlookAppointmentStructure);
    }

    /// <summary>
    /// Edit an existing Outlook appointment. The appointment's EntityID must be passed into an AppointmentStructure.
    /// Changes can be made to the existing appointment by passing them in the AppointmentStructure and setting the SetNewData parameter to true.
    /// </summary>
    public AppointmentStructure OutlookEditAppointment(bool Display, bool SetNewData, AppointmentStructure appointmentStructure)
    {

        Outlook._AppointmentItem appointment = null;

        try
        {
            if (_OutlookApplication == null)
            {
                _OutlookApplication = new Outlook.Application();
            }

            appointment = OutlookAppointmentItemByEntryID(appointmentStructure.EntryID);

            if (appointment != null)
            {
                if (SetNewData)
                {
                    OutlookAppointmentSetData(ref appointment, appointmentStructure);
                }

                if (Display)
                {
                    appointment.Display(true);   //.Recipients.ResolveAll()
                }
                else
                {
                    appointment.Save();
                }

                appointmentStructure = new AppointmentStructure();  //reset and get the appointment structure back out
                appointmentStructure = OutlookAppointmentStructureFromAppointment(appointment);
            }
            else
            {
                throw (new Exception("No appointment found"));
            }
        }
        catch (COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
        }
        return(appointmentStructure);
    }

    private bool OutlookAppointmentSetData(ref Outlook._AppointmentItem appointment, AppointmentStructure appointmentStructure)
    {
        bool returnSuccess = false;
        const string DoubleNewLine = "\r\n\r\n";

        try
        {

            if (appointment != null)
            {
                appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                appointment.Subject = appointmentStructure.Subject;
                appointment.Location = appointmentStructure.Location;
                appointment.Start = appointmentStructure.Start;
                appointment.End = appointmentStructure.End;
                appointment.ReminderMinutesBeforeStart = appointment.ReminderMinutesBeforeStart;
                appointment.ReminderSet = (appointmentStructure.ReminderMinutes > 0 ? true : false);

                //load attendees and resources
                if (appointmentStructure.RequiredAttendees != null)
                { 
                    string[] myListArrayReq = appointmentStructure.RequiredAttendees.Split(new Char[] { ';' });
                    foreach (string myStringReq in myListArrayReq)
                    {
                        if (myStringReq.Length > 0)
                        {
                            Outlook.Recipient recipient = appointment.Recipients.Add(myStringReq);
                            recipient.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
                        }
                    }
                }

                if (appointmentStructure.OptionalAttendees != null)
                {
                    string[] myListArrayOpt = appointmentStructure.OptionalAttendees.Split(new Char[] { ';' });
                    foreach (string myStringOpt in myListArrayOpt)
                    {
                        if (myStringOpt.Length > 0)
                        {
                            Outlook.Recipient recipient = appointment.Recipients.Add(myStringOpt);
                            recipient.Type = (int)Outlook.OlMeetingRecipientType.olOptional;
                        }
                    }
                }

                if (appointmentStructure.Resources != null)
                {
                    string[] myListArrayRes = appointmentStructure.Resources.Split(new Char[] { ';' });
                    foreach (string myStringRes in myListArrayRes)
                    {
                        if (myStringRes.Length > 0)
                        {
                            Outlook.Recipient recources = appointment.Recipients.Add(myStringRes);
                            recources.Type = (int)Outlook.OlMeetingRecipientType.olResource;
                        }
                    }
                }

                appointment.Recipients.ResolveAll();

                if (appointmentStructure.Body.TrimStart().EndsWith(DoubleNewLine))
                {
                    appointment.Body = appointmentStructure.Body.TrimStart();
                }
                else
                {
                    appointment.Body = appointmentStructure.Body.TrimStart() + DoubleNewLine;
                }

                while (appointment.Attachments.Count > 0)
                {
                    appointment.Attachments.Remove(1);
                }

                if (appointmentStructure.ListAttachments != null && appointmentStructure.ListAttachments.Count > 0)
                {
                    foreach (string PathFile in appointmentStructure.ListAttachments)
                    {
                        if (File.Exists(PathFile))
                        {
                            appointment.Attachments.Add(PathFile);
                        }
                    }
                }
                if (appointmentStructure.HighImportance)
                {
                    appointment.Importance = Outlook.OlImportance.olImportanceHigh;
                }
                else
                {
                    appointment.Importance = Outlook.OlImportance.olImportanceNormal;
                }
                appointment.BusyStatus = Outlook.OlBusyStatus.olBusy;
                appointment.ResponseRequested = appointmentStructure.ResponseRequested;
                returnSuccess = true;
            }
        }
        catch (COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
        }
        return (returnSuccess); 
    }

    private AppointmentStructure OutlookAppointmentGetData(string OutlookAppointmentEntryID)
    {
        Outlook._AppointmentItem appointment = null;
        AppointmentStructure appointmentStructure = new AppointmentStructure();

        if (OutlookAppointmentEntryID.Length > 0)
        {
            appointment = OutlookAppointmentItemByEntryID(OutlookAppointmentEntryID);
            if (appointment != null)
            {
                appointmentStructure = OutlookAppointmentStructureFromAppointment(appointment);
            }
        }
        return appointmentStructure;
    }

    private AppointmentStructure OutlookAppointmentStructureFromAppointment(Outlook._AppointmentItem outlookAppointmentItem)
    {
        AppointmentStructure appointmentStructure = new AppointmentStructure();
        if (outlookAppointmentItem != null & outlookAppointmentItem.EntryID.Length > 0)
        {
            appointmentStructure.EntryID = outlookAppointmentItem.EntryID;
            appointmentStructure.Subject = outlookAppointmentItem.Subject;
            appointmentStructure.Location = outlookAppointmentItem.Location;
            appointmentStructure.Start = outlookAppointmentItem.Start;
            appointmentStructure.End = outlookAppointmentItem.End;
            appointmentStructure.ReminderMinutes = outlookAppointmentItem.ReminderMinutesBeforeStart;
            appointmentStructure.RequiredAttendees = outlookAppointmentItem.RequiredAttendees;
            appointmentStructure.OptionalAttendees = outlookAppointmentItem.OptionalAttendees;
            appointmentStructure.Resources = outlookAppointmentItem.Resources;
            appointmentStructure.Body = outlookAppointmentItem.Body.ToString().Trim();
            appointmentStructure.ResponseRequested = outlookAppointmentItem.ResponseRequested;  
            if(outlookAppointmentItem.Importance == Outlook.OlImportance.olImportanceHigh)
            {
                appointmentStructure.HighImportance = true;
            }
            else
            {
                appointmentStructure.HighImportance = false;
            }
        }
        return appointmentStructure;
    }

    private Outlook._AppointmentItem OutlookAppointmentItemByEntryID(string EntryID)
    {
        //This function is needed because the EntryID or store may change with someone other that the 
        //originator of the appointment who saves the item.  Returns the appointment by given EntryID

        Outlook._AppointmentItem appointmentEntryID = null;
        try
        {
            Outlook.NameSpace nameSpace = _OutlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder folder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            if (folder != null)
            {
                appointmentEntryID = nameSpace.GetItemFromID(EntryID);


                //foreach (Outlook.AppointmentItem item in folder.Items)
                //{

                //    if (item.UserProperties[_EntryID].Value == EntryID)
                //    {
                //        appointmentEntryID = item;
                //        break;
                //    }
                //}
            }
        }
        catch (COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
        }

        return appointmentEntryID;
    }


    /// <summary>
    /// Send a DataGridView's data to an Excel worksheet.
    /// No formatting is set w/in Excel.
    /// </summary>
    public void DataGridViewToExcel(DataGridView myDataGridView, bool SelectedRowsOnly, bool ShowWhileLoading, ExcelDrawColors excelDrawColors)
    {
        System.Data.DataTable myDataTable = DataTableFromDataGridView(myDataGridView, SelectedRowsOnly);
        DateTableToExcel(myDataTable, ShowWhileLoading, excelDrawColors);
    }

    /// <summary>
    /// Build a DataTable from a DataGridView, connected or disconnected.
    /// </summary>
    public System.Data.DataTable DataTableFromDataGridView(DataGridView dataGridView, bool SelectedRowsOnly)
    {

        bool rowHasNonNulls = false;
        System.Data.DataTable dataTable = new System.Data.DataTable();
        for (int i = 0; i <= dataGridView.ColumnCount - 1; i++)
        {
            dataTable.Columns.Add(dataGridView.Columns[i].HeaderText, dataGridView.Columns[i].ValueType);
        }

        foreach (DataGridViewRow dataGridViewRow in dataGridView.Rows)
        {
            if (!SelectedRowsOnly || (SelectedRowsOnly && dataGridViewRow.Selected))
            {
                rowHasNonNulls = false;
                DataRow dataRow = dataTable.NewRow();
                for (int i = 1; i <= dataGridView.ColumnCount; i++)
                {
                    if (dataGridViewRow.Cells[i-1].Value != null)
                    {
                        rowHasNonNulls = true;
                        dataRow[i-1] = dataGridViewRow.Cells[i-1].Value;
                    }
                }
                if (rowHasNonNulls)
                {
                    dataTable.Rows.Add(dataRow);
                }
            }
        }

        return dataTable;
    }

    public System.Data.DataTable DataTableFromDataGridViewColumizeAddressLines(DataGridView dataGridView, bool SelectedRowsOnly)
    {

        System.Data.DataTable dataTable = new System.Data.DataTable();
        int donorAddressColumn = 0;
        int notesColumn = 0;
        int columnShift = 0;
        const char char160 = (char)160;

        for (int i = 1; i <= dataGridView.ColumnCount - 1; i++)
        {

            if (dataGridView.Columns[i].HeaderText == "DonorAddress")
            {
                donorAddressColumn = i;
                dataTable.Columns.Add("DonorAddress1", dataGridView.Columns[i].ValueType);
                dataTable.Columns.Add("DonorAddress2", dataGridView.Columns[i].ValueType);
                dataTable.Columns.Add("DonorAddress3", dataGridView.Columns[i].ValueType);
                dataTable.Columns.Add("DonorAddress4", dataGridView.Columns[i].ValueType);
                dataTable.Columns.Add("City State Zip", dataGridView.Columns[i].ValueType);
            }

            if (dataGridView.Columns[i].HeaderText == "Notes")
            {
                notesColumn = i;
            }
            dataTable.Columns.Add(dataGridView.Columns[i].HeaderText, dataGridView.Columns[i].ValueType);
        }

        foreach (DataGridViewRow dataGridViewRow in dataGridView.Rows)
        {
            columnShift = 0;
            if (!SelectedRowsOnly || (SelectedRowsOnly && dataGridViewRow.Selected))
            {
                DataRow dataRow = dataTable.NewRow();


                for (int i = 1; i <= dataGridView.ColumnCount - 1; i++)
                {
                    if (i == notesColumn)
                    {
                        if (dataGridViewRow.Cells[i].Value.ToString().Length > 10)
                        {
                            dataRow[i - 1 + columnShift] = dataGridViewRow.Cells[i].Value.ToString().Substring(0, 9) + "...";
                        }
                        else if (dataGridViewRow.Cells[i].Value.ToString().Length < 10)
                        {
                            dataRow[i - 1 + columnShift] = dataGridViewRow.Cells[i].Value;
                        }
                    }
                    else
                    {
                        dataRow[i - 1 + columnShift] = dataGridViewRow.Cells[i].Value;
                    }

                    if (i == donorAddressColumn)
                    {
                        columnShift = 5;
                        string[] address = dataGridViewRow.Cells[i].Value.ToString().Split(new Char[] {','});
                        for (int addAddress = 0; addAddress <= address.GetUpperBound(0); addAddress++)
                        {
                            if (address[addAddress].ToString().Contains(char160.ToString()) )
                            {
                                dataRow[i + 4] = address[addAddress].ToString();
                            }
                            else
                            {
                                dataRow[i + addAddress] = address[addAddress].ToString();
                            }
                        }
                    }
                }
                dataTable.Rows.Add(dataRow);
            }
        }

        return dataTable;
    }

    /// <summary>
    /// Send a DataTable to an Excel worksheet.
    /// No formatting is set w/in Excel.
    /// </summary>
    public void DateTableToExcel(System.Data.DataTable dataTable, bool ShowWhileLoading, ExcelDrawColors excelDrawColors)
    {
        _ExcelWorkbook = null;
        _ExcelWorksheet = null;

        long ColorBlueDark = 14136213; //ColorTranslator.ToOle((Color)colorConverter.ConvertFromString("#95B3D7")); System.Drawing.ColorConverter cc = new System.Drawing.ColorConverter();
        long ColorBlueLight = 15853276;
        long ColorOrangeDark = 4626167;
        long ColorOrangeLight = 14281213;

        long ColorHeader = 0;
        long ColorAltRow = 0;

        if (excelDrawColors == ExcelDrawColors.Blue)
        {
            ColorHeader = ColorBlueDark;
            ColorAltRow = ColorBlueLight; 
        }
        else if (excelDrawColors == ExcelDrawColors.Orange)
        {
            ColorHeader = ColorOrangeDark;
            ColorAltRow = ColorOrangeLight;
        }

        try
        {
            if (_ExcelApplication == null)
            {
                _ExcelApplication = new Excel.Application();
            }

            _ExcelWorkbook = _ExcelApplication.Workbooks.Add();
            _ExcelWorksheet = _ExcelWorkbook.Worksheets.Add(); 
            _ExcelWorksheet.Activate();

            _ExcelApplication.Visible = ShowWhileLoading;

            //get header
            for (int i = 0; i <= dataTable.Columns.Count - 1; i++)
            {

                _ExcelWorksheet.Cells[1, i + 1] = dataTable.Columns[i].Caption.ToString();
                _ExcelWorksheet.Cells[1, i + 1].Font.Bold = true;
                if (excelDrawColors != ExcelDrawColors.None)
                {
                    _ExcelWorksheet.Cells[1, i + 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    _ExcelWorksheet.Cells[1, i + 1].Interior.Color = ColorHeader; //System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SteelBlue);
                }

                switch (dataTable.Columns[i].DataType.ToString())
                {
                    case "System.Date":
                        _ExcelWorksheet.Columns[i + 1].NumberFormat = "[$-1809]yyyy-mm-dd;@";
                        break;
                    case "System.Decimal":
                        _ExcelWorksheet.Columns[i + 1].NumberFormat = "$#,##0.00;[Red]-$#,##0.00";
                        break;
                    default:
                        // "System.String"
                        _ExcelWorksheet.Columns[i + 1].NumberFormat = "@";
                        break;
                }
            }
            //_ExcelWorksheet.Rows[1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SteelBlue);

            //get data
            for (int dtRow = 0; dtRow <= dataTable.Rows.Count - 1; dtRow++)
            {
                DataRow dataRow = dataTable.Rows[dtRow];
                for (int dtColumn = 0; dtColumn <= dataTable.Columns.Count - 1; dtColumn++)
                {
                    if (dataTable.Columns[dtColumn].Caption.ToString().Contains("Address"))
                    {
                        _ExcelWorksheet.Cells[dtRow + 2, dtColumn + 1] = dataRow[dtColumn].ToString().Replace(",", "\r\n");
                    }
                    else
                    {
                        _ExcelWorksheet.Cells[dtRow + 2, dtColumn + 1] = dataRow[dtColumn].ToString();
                    }

                    if (excelDrawColors != ExcelDrawColors.None && dtRow % 2 == 0)
                    {
                        _ExcelWorksheet.Cells[dtRow + 2, dtColumn + 1].Interior.Color = ColorAltRow; //System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue);
                    }
                }

                if (dataTable.Rows.Count < 1000)  // a little faster to show just the 10s
                {
                    _ExcelApplication.StatusBar = "writing data row " + Convert.ToString(dtRow) + " of " + Convert.ToString(dataTable.Rows.Count) + ", please wait!";
                }
                else if (dtRow % 10 == 0)
                {
                    _ExcelApplication.StatusBar = "writing data row " + Convert.ToString(dtRow) + " of " + Convert.ToString(dataTable.Rows.Count) + ", please wait!";
                }
            }

            //Excel.FormatCondition format = _ExcelWorksheet.Rows.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2) = 0");
            //format.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Wheat);
            
            _ExcelApplication.StatusBar = "Auto-sizing columns";
            _ExcelWorksheet.Columns.AutoFit();
            _ExcelApplication.StatusBar = "";
            _ExcelApplication.Visible = true;

        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw (new Exception("COM Exception: " + ex.Message));
        }
        catch (Exception e)
        {
            throw (e);
        }
        finally
        {
        }
    }
    
    
    private void OfficeObjectAutomationQuitReleaseWord_Application()
    {
        if (_WordApplication != null)
        {
            try
            {
                object saveOptionsObject = Word.WdSaveOptions.wdDoNotSaveChanges;
                _WordApplication.Quit(saveOptionsObject); 
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_WordApplication);
            }
        }
    }

    private void OfficeObjectAutomationQuitReleaseExcel_Application()
    {
        if (_ExcelApplication != null)
        {
            try
            {
                object saveOptionsObject = Excel.XlSaveAction.xlDoNotSaveChanges; 
                foreach(Excel._Workbook _ExcelWorkbook in _ExcelApplication.Workbooks)
                {
                    _ExcelWorkbook.Close(saveOptionsObject);
                }
                _ExcelApplication.Quit();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_ExcelApplication);
            }
        }
    }

    private void OfficeObjectAutomationQuitReleaseOutlook_Application()
    {
        if (_OutlookApplication != null)
        {
            try
            {
                _OutlookApplication.Quit();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_OutlookApplication);
            }
        }
    }

    ~OfficeAutomationHelper()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(true);
    }

    protected virtual void Dispose(bool IsDisposing)
    {
        if (_IsDisposed)
        {
            return;
        }

        if (IsDisposing)
        {
            //_WordApplication.Quit();
            //_ExcelApplication.Quit();
            //_OutlookApplication.Quit();
            OfficeObjectAutomationQuitReleaseWord_Application();
            OfficeObjectAutomationQuitReleaseExcel_Application();
            OfficeObjectAutomationQuitReleaseOutlook_Application();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        _IsDisposed = true;  // Free any unmanaged resources in this section
    }
      
       
    
    public class ClassDisplayWindowUsingThread
    {
        public void StartThreadSpellingWindowOnTop(string[] WindowCaption)
        {
            Thread myThread = null;
            ClassDisplayWindow myClassDisplayWindowThreadStarter = new ClassDisplayWindow();
            myClassDisplayWindowThreadStarter.WindowCaption = WindowCaption;
            myThread = new Thread(myClassDisplayWindowThreadStarter.SpellingWindowOnTop);
            myThread.Start();
        }

        public void StartThreadSpellingWindowOnTop(string WindowApplicationName, string[] WindowCaption)
        {
            Thread myThread = null;
            ClassDisplayWindow myClassDisplayWindowThreadStarter = new ClassDisplayWindow();
            myClassDisplayWindowThreadStarter.WindowCaption = WindowCaption;
            myThread = new Thread(myClassDisplayWindowThreadStarter.SpellingWindowOnTop);
            myThread.Start();
        }

    }
        
    public class ClassDisplayWindow
    {
        private const int SW_SHOWNOACTIVATE = 4;
        private const int HWND_TOPMOST = -1;
        private const int HWND_NOTOPMOST = -2;
        private const uint SWP_NOACTIVATE = 16;
        private const int SWP_NOMOVE = 2;
        private const int SWP_NOSIZE = 1;
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        private const int SW_SHOWNORMAL = 1;
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowByCaption(IntPtr zero, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool ShowWindow(IntPtr hwnd, Int32 nCmdShow);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        private static extern bool SetWindowPos(int hWnd, int hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private string _WindowApplicationName;
        public string WindowApplicationName
        {
            get { return _WindowApplicationName; }
            set { _WindowApplicationName = value; }
        }

        private string[] _WindowCaption;
        public string[] WindowCaption
        {
            get { return _WindowCaption; }
            set { _WindowCaption = value; }
        }

        private IntPtr FindWindowByApplicationAndName(string WindowApplicationName, string WindowCaption)
        {
            IntPtr hWnd = new IntPtr(0);
            hWnd = FindWindow(WindowApplicationName, WindowCaption);
            return hWnd;
        }

        private IntPtr FindWindowByCaption(string WindowCaption)
        {
            IntPtr hWnd = new IntPtr(0);
            hWnd = FindWindow(null, WindowCaption);
            return hWnd;
        }

        private void ShowWindowByHandle(IntPtr WindowHandle, bool Show)
        {
            Int32 ShowValue = SW_HIDE;
            if (Show)
                ShowValue = SW_SHOWNORMAL;
            ShowWindow(WindowHandle, ShowValue);
        }

        private void ShowInactiveTopmost(IntPtr WindowHandle, bool OnTop)
        {
            int OnTopValue = HWND_NOTOPMOST;
            if (OnTop)
                OnTopValue = HWND_TOPMOST;
            ShowWindow(WindowHandle, SW_SHOWNOACTIVATE);
            SetWindowPos(WindowHandle.ToInt32(), OnTopValue, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE);
        }

        public void SpellingWindowOnTop()
        {
            IntPtr hWnd = default(IntPtr);
            int i = 1;

            while (i < 100)
            {
                Thread.Sleep(100);
                foreach (string myWindowCaption in _WindowCaption)
                {
                    if (_WindowApplicationName != string.Empty)
                    {
                        hWnd = FindWindowByApplicationAndName(_WindowApplicationName, myWindowCaption);
                    }
                    else
                    {
                        hWnd = FindWindowByCaption(myWindowCaption);
                    }
                    if (hWnd != new IntPtr(0))
                    {
                        break; 
                    }
                    else
                    {
                        i += 1;
                    }
                }
            }

            if (hWnd != new IntPtr(0))
            {
                ShowInactiveTopmost(hWnd, true);
                ShowInactiveTopmost(hWnd, false);
            }
        }
    }
}

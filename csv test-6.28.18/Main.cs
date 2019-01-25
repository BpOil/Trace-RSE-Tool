using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

namespace csv_test_6._28._18
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        public DataTable dataTable;
        StreamReader streamer;
        string contactpath = string.Empty;
        string reportpath = string.Empty;
        string path = string.Empty;
        string itemText = string.Empty;
        string fileType = string.Empty;

        //beginning of the attempt to get the loading gif to work 
        int currentYear = DateTime.Now.Year;
        Excel.Application xlApp;
        Excel.Workbook vishingNotesWB;
        Excel.Worksheet vishingNotesWS;
        long vishingNotesMaxRow = 0;
        int vishingVoicemailCount = 0;
        int vishingPassedCount = 0;
        int vishingFailedCount = 0;
        Excel.Workbook phishingResultsWB;
        Excel.Worksheet phishingResultsWS;
        int phishingResultsMaxRow = 0;
        int phishingTotalEmails = 0;
        int phishingFailedCount = 0;
        int phishingOpenedCount = 0;
        Word.Application wdApp;
        Word.Document reportDoc;
        int progressBarPercentage = 0;

        private void Main_Load(object sender, EventArgs e)
        {

        }

        // **METHOD THAT OPENS FILE EXPLORER AND FOCUSES THE NEWLY SAVED ITEM BY THE USER
        private void OpenFolder(string folderPath)
        {
            if (File.Exists(folderPath))
            {
                Process.Start(new ProcessStartInfo("explorer.exe", " /select, " + folderPath));
            }
        }

        // ***************************************************************************************************************************************************************
        // ***************************************************************************************************************************************************************
        // ********-----------------------------------------------------------------------------------------------------------------------------------------------********
        // ********--------------------------------------------C S V    S T U F F---------------------------------------------------------------------------------********
        // ********-----------------------------------------------------------------------------------------------------------------------------------------------********
        // ***************************************************************************************************************************************************************
        // ***************************************************************************************************************************************************************

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            ToolTip toolOpenFile = new ToolTip();
            toolOpenFile.ShowAlways = true;
            toolOpenFile.SetToolTip(btnOpenFile, "Open Scoping Form");
            try
            {
                OpenFileDialog openFile = new OpenFileDialog() { Filter = "Word Document|*.doc;*.docx", ValidateNames = true };
                DialogResult result = openFile.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //set fileType as Word and then create the Data Table that will be displayed in the preview window and store the Data Table into the class variable dataTable
                    fileType = "Word";
                    contactpath = openFile.FileName;
                    dataTable = wordDocToDataTable(contactpath);

                    btnPreview.Visible = true;
                    btnPreview2.Visible = false;
                    btnInsightCSV.Enabled = true;
                    btnKnowbe4CSV.Enabled = true;
                    txtUserGroup.Enabled = true;
                    btnCreateCallList.Enabled = true;
                    txtUserGroup.BackColor = System.Drawing.Color.White;
                    btnInsightCSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnInsightCSV.ForeColor = System.Drawing.Color.White;
                    btnKnowbe4CSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnKnowbe4CSV.ForeColor = System.Drawing.Color.White;
                    btnCreateCallList.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnCreateCallList.ForeColor = System.Drawing.Color.White;
                    lblPullContacts.Left = 3;
                    lblPullContacts.Top = 70;
                    lblPullContacts.Visible = true;
                    lblPullContacts.ForeColor = System.Drawing.Color.Lime;
                    lblPullContacts.Text = "Success extracting data";
                }

            }
            catch
            {
                btnPreview.Visible = false;
                btnPreview2.Visible = false;
                btnInsightCSV.Enabled = false;
                btnKnowbe4CSV.Enabled = false;
                txtUserGroup.Enabled = false;
                btnCreateCallList.Enabled = false;
                txtUserGroup.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnInsightCSV.ForeColor = System.Drawing.Color.LightGray;
                btnKnowbe4CSV.BackColor = System.Drawing.Color.Gray;
                btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnKnowbe4CSV.ForeColor = System.Drawing.Color.LightGray;
                btnCreateCallList.BackColor = System.Drawing.Color.Gray;
                btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnCreateCallList.ForeColor = System.Drawing.Color.LightGray;
                lblPullContacts.Left = 3;
                lblPullContacts.Top = 70;
                lblPullContacts.Visible = true;
                lblPullContacts.ForeColor = System.Drawing.Color.Red;
                lblPullContacts.Text = "Failed extracting data";
                MessageBox.Show("The file could not be loaded");
            }
        }

        //import excel workbooks to get employee contact info 
        private void btnOpenExcelFile_Click(object sender, EventArgs e)
        {
            ToolTip toolOpenFile = new ToolTip();
            toolOpenFile.ShowAlways = true;
            toolOpenFile.SetToolTip(btnOpenExcelFile, "Open Excel Sheet");
            try
            {
                OpenFileDialog openFile = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true };
                DialogResult result = openFile.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //set fileType as Excel and then create the Data Table that will be displayed in the preview window and store the Data Table into the class variable dataTable
                    fileType = "Excel";
                    contactpath = openFile.FileName;
                    dataTable = excelSheetToDataTable(contactpath, true);

                    btnPreview2.Visible = true;
                    btnPreview.Visible = false;
                    btnInsightCSV.Enabled = true;
                    btnKnowbe4CSV.Enabled = true;
                    txtUserGroup.Enabled = true;
                    btnCreateCallList.Enabled = true;
                    txtUserGroup.BackColor = System.Drawing.Color.White;
                    btnInsightCSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnInsightCSV.ForeColor = System.Drawing.Color.White;
                    btnKnowbe4CSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnKnowbe4CSV.ForeColor = System.Drawing.Color.White;
                    btnCreateCallList.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnCreateCallList.ForeColor = System.Drawing.Color.White;
                    lblPullContacts.Left = 213;
                    lblPullContacts.Top = 68;
                    lblPullContacts.Visible = true;
                    lblPullContacts.ForeColor = System.Drawing.Color.Lime;
                    lblPullContacts.Text = "Success extracting data";
                }

            }
            catch
            {
                btnPreview2.Visible = false;
                btnPreview.Visible = false;
                btnInsightCSV.Enabled = false;
                btnKnowbe4CSV.Enabled = false;
                txtUserGroup.Enabled = false;
                btnCreateCallList.Enabled = false;
                txtUserGroup.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnInsightCSV.ForeColor = System.Drawing.Color.LightGray;
                btnKnowbe4CSV.BackColor = System.Drawing.Color.Gray;
                btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnKnowbe4CSV.ForeColor = System.Drawing.Color.LightGray;
                btnCreateCallList.BackColor = System.Drawing.Color.Gray;
                btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnCreateCallList.ForeColor = System.Drawing.Color.LightGray;
                lblPullContacts.Left = 213;
                lblPullContacts.Top = 68;
                lblPullContacts.Visible = true;
                lblPullContacts.ForeColor = System.Drawing.Color.Red;
                lblPullContacts.Text = "Failed extracting data";
                MessageBox.Show("The file could not be loaded");
            }
        }

        //converts a Word Document Scoping form into a Data Table
        private DataTable wordDocToDataTable(string filePath)
        {
            Read displayData = new Read();
            displayData.NameFile = filePath;
            //get headers
            string[] copyHeader = displayData.WordTableHeader();
            //get table data 
            string[,] displayArray = displayData.WordDoc();

            //see if First Name and Last Name are in seperate columns
            int firstNameColumn = -1;
            int lastNameColumn = -1;
            string firstNameColumnName = null;
            string lastNameColumnName = null;
            for (int i = 0; i < copyHeader.Length; i++)
            {
                if (copyHeader[i].ToLower().Contains("first"))
                {
                    firstNameColumn = i;
                    firstNameColumnName = copyHeader[i];
                }
                else if (copyHeader[i].ToLower().Contains("last"))
                {
                    lastNameColumn = i;
                    lastNameColumnName = copyHeader[i];
                }
            }


            DataTable result = new DataTable();
            //create DataTable
            if (firstNameColumn != -1 & lastNameColumn != -1)
            { //create DataTable if First Name and Last Name are in SEPERATE columns
                //add headers to the columns in the DataTable
                result.Columns.Add("Name", typeof(String));
                for (int i = 0; i < copyHeader.Length; i++)
                {
                    if (!copyHeader[i].Equals(firstNameColumnName) & !copyHeader[i].Equals(lastNameColumnName))
                    {
                        result.Columns.Add(copyHeader[i], typeof(String));
                    }
                }
                //add employee info to the DataTable
                for (int i = 0; i < (displayArray.Length / copyHeader.Length); i++)
                {
                    DataRow row = result.NewRow();
                    row["Name"] = displayArray[i, firstNameColumn] + " " + displayArray[i, lastNameColumn];
                    for (int j = 0; j < copyHeader.Length; j++)
                    {
                        if (!copyHeader[j].Equals(firstNameColumnName) & !copyHeader[j].Equals(lastNameColumnName))
                        {
                            row[copyHeader[j]] = displayArray[i, j];
                        }
                    }
                    result.Rows.Add(row);
                }

            }
            else
            { //create DataTable if First Name and Last Name are in the SAME column
                //add headers to the columns in the DataTable
                for (int i = 0; i < copyHeader.Length; i++)
                {
                    result.Columns.Add(copyHeader[i], typeof(String));
                }
                //add employee info to the DataTable
                for (int i = 0; i < (displayArray.Length / copyHeader.Length); i++)
                {
                    DataRow row = result.NewRow();
                    for (int j = 0; j < copyHeader.Length; j++)
                    {
                        row[copyHeader[j]] = displayArray[i, j];
                    }
                    result.Rows.Add(row);
                }
            }
            return result;
        }

        //converts a Excel Workbook Scoping form into a Data Table
        private DataTable excelSheetToDataTable(string filePath, bool useFirstRowAsHeaders)
        {
            var file = new FileInfo(filePath);
            IExcelDataReader reader;
            FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read);
            if (file.Extension.Equals(".xls"))
                reader = ExcelReaderFactory.CreateBinaryReader(fs);
            else if (file.Extension.Equals(".xlsx"))
                reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            else if (file.Extension.Equals(".csv"))
                reader = ExcelReaderFactory.CreateCsvReader(fs);
            else
                throw new Exception("Invalid FileName");

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = useFirstRowAsHeaders
                }
            };

            var dataSet = reader.AsDataSet(conf);
            var dt = dataSet.Tables[0];
            reader.Close();

            List<string> initialHeaders = new List<string>();
            foreach (DataColumn column in dt.Columns)
            {
                initialHeaders.Add(column.ColumnName);
            }
            int firstNameColumn = -1;
            int lastNameColumn = -1;
            string firstNameColumnName = null;
            string lastNameColumnName = null;
            for (int i = 0; i < initialHeaders.Count; i++)
            {
                if (initialHeaders[i].ToLower().Contains("first"))
                {
                    firstNameColumn = i;
                    firstNameColumnName = initialHeaders[i];
                }
                else if (initialHeaders[i].ToLower().Contains("last"))
                {
                    lastNameColumn = i;
                    lastNameColumnName = initialHeaders[i];
                }
            }

            DataTable result = new DataTable();
            List<String> headers = new List<String>();
            if (firstNameColumn != -1 & lastNameColumn != -1)
            { //create DataTable if First Name and Last Name are in SEPERATE columns
                result.Columns.Add("Name", typeof(String));
                foreach (DataColumn column in dt.Columns)
                {
                    if (!column.ColumnName.Equals(firstNameColumnName) & !column.ColumnName.Equals(lastNameColumnName))
                    {
                        result.Columns.Add(column.ColumnName, typeof(String));
                    }
                }
                foreach (DataRow dr in dt.Rows)
                {
                    DataRow row = result.NewRow();
                    row["Name"] = dr[firstNameColumnName].ToString() + " " + dr[lastNameColumnName];
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (!column.ColumnName.Equals(firstNameColumnName) & !column.ColumnName.Equals(lastNameColumnName))
                        {
                            row[column.ColumnName] = dr[column.ColumnName];
                        }
                    }
                    result.Rows.Add(row);
                }
                foreach (DataColumn column in result.Columns)
                {
                    headers.Add(column.ColumnName);
                }
            }
            else
            { //create DataTable if First Name and Last Name are in the SAME column
                result = dt;
                headers = initialHeaders;
            }
            fs.Close();
            return result;
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            ToolTip toolPreview = new ToolTip();
            toolPreview.ShowAlways = false;
            toolPreview.SetToolTip(btnPreview, "Preview Extracted Data");
            Preview newpreview = new Preview();
            newpreview.dataTable = dataTable;
            newpreview.ShowDialog();
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            ToolTip toolPreview = new ToolTip();
            toolPreview.ShowAlways = false;
            toolPreview.SetToolTip(btnPreview, "Preview Extracted Data");
            Preview newpreview = new Preview();
            newpreview.dataTable = dataTable;
            newpreview.ShowDialog();
        }
        // ****************************************************************************************************************************************************************************************************
        // ****************************************************************************************************************************************************************************************************
        // ********------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------********
        // ********-----------------------------------------------------------------C R E A T E   C S V   F I L E   S T U F F--------------------------------------------------------------------------********
        // ********------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------********
        // ****************************************************************************************************************************************************************************************************
        // ****************************************************************************************************************************************************************************************************

        private void btnInsightCSV_Click(object sender, EventArgs e)
        {
            string userGroup = txtUserGroup.Text.ToString();
            Read reading = new Read();

            //   __________________________
            // ||__________________________||
            // ||                          ||
            // ||   W O R D   V A L U E S  ||
            // ||__________________________||
            // ||__________________________||
            if (fileType == "Word")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.WordTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.WordDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0) + 1;
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 7];
                try
                {
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < colcount; j++)
                        {
                            //   __________________________
                            // ||                          ||
                            // ||  B L A N K   V A L U E S ||
                            // ||__________________________||
                            // SET MIDDLE NAME VALUE
                            reorderData[i, 1] = " ";
                            // SET USERGROUP VALUE
                            if (!String.IsNullOrWhiteSpace(userGroup))
                            {
                                reorderData[i, 6] = userGroup;
                            }
                            else
                            {
                                reorderData[i, 6] = " ";
                            }
                            // SET FIRST NAME VALUE
                            if (copyHeader[j].Contains("first"))
                            {
                                reorderData[i, 0] = copyData[i, j];
                            }
                            // SET LAST NAME VALUE
                            else if (copyHeader[j].Contains("last"))
                            {
                                reorderData[i, 2] = copyData[i, j];
                            }
                            // SET TITLE VALUE
                            else if (copyHeader[j].Contains("title"))
                            {
                                reorderData[i, 3] = copyData[i, j];
                            }
                            // SET PHONE NUMBER VALUE
                            else if (copyHeader[j].Contains("phone"))
                            {
                                reorderData[i, 4] = copyData[i, j];
                            }
                            // SET EMAIL ADDRESS VALUE
                            else if (copyHeader[j].Contains("email"))
                            {
                                reorderData[i, 5] = copyData[i, j];
                            }
                            else
                            {

                            }
                        }
                    }
                    string thisfile = String.Empty;
                    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                    SaveFileDialog fileStream = new SaveFileDialog();
                    fileStream.FileName = "insightupload.csv";
                    fileStream.DefaultExt = ".csv";
                    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                    DialogResult result = fileStream.ShowDialog();
                    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                    if (result == DialogResult.OK)
                    {
                        thisfile = fileStream.FileName;
                    }
                    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                    // CALL CREATE CLASS'S CSV-MAKING METHOD
                    makeFile.InsightUpload();
                    OpenFolder(thisfile);
                }
                catch
                {
                    MessageBox.Show("Unable to export data to file properly.");
                }
            }

            //   __________________________
            // ||__________________________||
            // ||                          ||
            // ||  E X C E L   V A L U E S ||
            // ||__________________________||
            // ||__________________________||
            else if (fileType == "Excel")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.ExcelTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.ExcelDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0);
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 7];

                for (int i = 0; i < rowcount; i++)
                {
                    for (int j = 0; j < colcount; j++)
                    {
                        //   __________________________
                        // ||                          ||
                        // ||  B L A N K   V A L U E S ||
                        // ||__________________________||
                        // SET MIDDLE NAME VALUE
                        reorderData[i, 1] = " ";
                        // SET USERGROUP VALUE
                        if (!String.IsNullOrWhiteSpace(userGroup))
                        {
                            reorderData[i, 6] = userGroup;
                        }
                        else
                        {
                            reorderData[i, 6] = " ";
                        }
                        // SET FIRST NAME VALUE
                        if (copyHeader[j].Contains("first"))
                        {
                            reorderData[i, 0] = copyData[i, j];
                        }
                        // SET LAST NAME VALUE
                        else if (copyHeader[j].Contains("last"))
                        {
                            reorderData[i, 2] = copyData[i, j];
                        }
                        // SET TITLE VALUE
                        else if (copyHeader[j].Contains("title"))
                        {
                            reorderData[i, 3] = copyData[i, j];
                        }
                        // SET PHONE NUMBER VALUE
                        else if (copyHeader[j].Contains("phone"))
                        {
                            reorderData[i, 4] = copyData[i, j];
                        }
                        // SET EMAIL ADDRESS VALUE
                        else if (copyHeader[j].Contains("email"))
                        {
                            reorderData[i, 5] = copyData[i, j];
                        }
                        else
                        {

                        }
                    }
                }
                string thisfile = String.Empty;
                // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                SaveFileDialog fileStream = new SaveFileDialog();
                fileStream.FileName = "insightupload.csv";
                fileStream.DefaultExt = ".csv";
                fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                DialogResult result = fileStream.ShowDialog();
                // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                if (result == DialogResult.OK)
                {
                    thisfile = fileStream.FileName;
                }
                // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                // CALL CREATE CLASS'S CSV-MAKING METHOD
                makeFile.InsightUpload();
                OpenFolder(thisfile);
                //try
                //{
                //    for (int i = 0; i < rowcount; i++)
                //    {
                //        for (int j = 0; j < colcount; j++)
                //        {
                //            //   __________________________
                //            // ||                          ||
                //            // ||  B L A N K   V A L U E S ||
                //            // ||__________________________||
                //            // SET MIDDLE NAME VALUE
                //            reorderData[i, 1] = " ";
                //            // SET USERGROUP VALUE
                //            if (!String.IsNullOrWhiteSpace(userGroup))
                //            {
                //                reorderData[i, 6] = userGroup;
                //            }
                //            else
                //            {
                //                reorderData[i, 6] = " ";
                //            }
                //            // SET FIRST NAME VALUE
                //            if (copyHeader[j].Contains("first"))
                //            {
                //                reorderData[i, 0] = copyData[i, j];
                //            }
                //            // SET LAST NAME VALUE
                //            else if (copyHeader[j].Contains("last"))
                //            {
                //                reorderData[i, 2] = copyData[i, j];
                //            }
                //            // SET TITLE VALUE
                //            else if (copyHeader[j].Contains("title"))
                //            {
                //                reorderData[i, 3] = copyData[i, j];
                //            }
                //            // SET PHONE NUMBER VALUE
                //            else if (copyHeader[j].Contains("phone"))
                //            {
                //                reorderData[i, 4] = copyData[i, j];
                //            }
                //            // SET EMAIL ADDRESS VALUE
                //            else if (copyHeader[j].Contains("email"))
                //            {
                //                reorderData[i, 5] = copyData[i, j];
                //            }
                //            else
                //            {

                //            }
                //        }
                //    }
                //    string thisfile = String.Empty;
                //    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                //    SaveFileDialog fileStream = new SaveFileDialog();
                //    fileStream.FileName = "insightupload.csv";
                //    fileStream.DefaultExt = ".csv";
                //    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                //    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                //    DialogResult result = fileStream.ShowDialog();
                //    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                //    if (result == DialogResult.OK)
                //    {
                //        thisfile = fileStream.FileName;
                //    }
                //    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                //    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                //    // CALL CREATE CLASS'S CSV-MAKING METHOD
                //    makeFile.InsightUpload();
                //    OpenFolder(thisfile);

                //}
                //catch
                //{
                //    MessageBox.Show("Unable to export data to file properly.");
                //}
            }


        }
        private void btnKnowbe4CSV_Click(object sender, EventArgs e)
        {
            string userGroup = txtUserGroup.Text.ToString();
            Read reading = new Read();

            //   __________________________
            // ||__________________________||
            // ||                          ||
            // ||   W O R D   V A L U E S  ||
            // ||__________________________||
            // ||__________________________||
            if (fileType == "Word")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.WordTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.WordDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0) + 1;
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 15];

                try
                {
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < colcount; j++)
                        {
                            //   __________________________
                            // ||                          ||
                            // ||  B L A N K   V A L U E S ||
                            // ||__________________________||
                            // SET LOCATION VALUE
                            reorderData[i, 6] = " ";
                            // SET DIVISION VALUE
                            reorderData[i, 7] = " ";
                            // SET MANAGER NAME VALUE
                            reorderData[i, 8] = " ";
                            // SET MANAGER EMAIL VALUE
                            reorderData[i, 9] = " ";
                            // SET EMPLOYEE NUMBER VALUE
                            reorderData[i, 10] = " ";
                            // SET PASSWORD VALUE
                            reorderData[i, 12] = " ";
                            // SET MOBILE NUMBER VALUE
                            reorderData[i, 13] = " ";
                            // SET AD MANAGED VALUE
                            reorderData[i, 14] = " ";
                            // SET USERGROUP VALUE
                            if (!String.IsNullOrWhiteSpace(userGroup))
                            {
                                reorderData[i, 5] = userGroup;
                            }
                            else
                            {
                                reorderData[i, 5] = " ";
                            }
                            // SET EMAIL VALUE
                            if (copyHeader[j].Contains("email"))
                            {
                                reorderData[i, 0] = copyData[i, j];
                            }
                            // SET FIRST NAME VALUE
                            else if (copyHeader[j].Contains("first"))
                            {
                                reorderData[i, 1] = copyData[i, j];
                            }
                            // SET LAST NAME VALUE
                            else if (copyHeader[j].Contains("last"))
                            {
                                reorderData[i, 2] = copyData[i, j];
                            }
                            // SET PHONE NUMBER VALUE
                            else if (copyHeader[j].Contains("phone"))
                            {
                                reorderData[i, 3] = copyData[i, j];
                            }
                            // SET EXTENSION VALUE
                            else if (copyHeader[j].Contains("ext") || copyHeader[j].Contains("extension"))
                            {
                                reorderData[i, 4] = copyData[i, j];
                            }
                            // SET TITLE VALUE
                            else if (copyHeader[j].Contains("title"))
                            {
                                reorderData[i, 11] = copyData[i, j];
                            }
                            else
                            {

                            }
                        }
                    }

                    string thisfile = String.Empty;
                    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                    SaveFileDialog fileStream = new SaveFileDialog();
                    fileStream.FileName = "resellerupload.csv";
                    fileStream.DefaultExt = ".csv";
                    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                    DialogResult result = fileStream.ShowDialog();
                    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                    if (result == DialogResult.OK)
                    {
                        thisfile = fileStream.FileName;
                    }
                    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                    // CALL CREATE CLASS'S CSV-MAKING METHOD
                    makeFile.ResellerUpload();
                    OpenFolder(thisfile);
                }
                catch
                {
                    MessageBox.Show("Unable to export data to file properly.");
                }
            }
            //   __________________________
            // ||__________________________||
            // ||                          ||
            // ||  E X C E L   V A L U E S ||
            // ||__________________________||
            // ||__________________________||
            else if (fileType == "Excel")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.ExcelTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.ExcelDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0) + 1;
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 15];

                try
                {
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < colcount; j++)
                        {
                            //   __________________________
                            // ||                          ||
                            // ||  B L A N K   V A L U E S ||
                            // ||__________________________||
                            // SET LOCATION VALUE
                            reorderData[i, 6] = " ";
                            // SET DIVISION VALUE
                            reorderData[i, 7] = " ";
                            // SET MANAGER NAME VALUE
                            reorderData[i, 8] = " ";
                            // SET MANAGER EMAIL VALUE
                            reorderData[i, 9] = " ";
                            // SET EMPLOYEE NUMBER VALUE
                            reorderData[i, 10] = " ";
                            // SET PASSWORD VALUE
                            reorderData[i, 12] = " ";
                            // SET MOBILE NUMBER VALUE
                            reorderData[i, 13] = " ";
                            // SET AD MANAGED VALUE
                            reorderData[i, 14] = " ";
                            // SET USERGROUP VALUE
                            if (!String.IsNullOrWhiteSpace(userGroup))
                            {
                                reorderData[i, 5] = userGroup;
                            }
                            else
                            {
                                reorderData[i, 5] = " ";
                            }
                            // SET EMAIL VALUE
                            if (copyHeader[j].Contains("email"))
                            {
                                reorderData[i, 0] = copyData[i, j];
                            }
                            // SET FIRST NAME VALUE
                            else if (copyHeader[j].Contains("first"))
                            {
                                reorderData[i, 1] = copyData[i, j];
                            }
                            // SET LAST NAME VALUE
                            else if (copyHeader[j].Contains("last"))
                            {
                                reorderData[i, 2] = copyData[i, j];
                            }
                            // SET PHONE NUMBER VALUE
                            else if (copyHeader[j].Contains("phone"))
                            {
                                reorderData[i, 3] = copyData[i, j];
                            }
                            // SET EXTENSION VALUE
                            else if (copyHeader[j].Contains("ext") || copyHeader[j].Contains("extension"))
                            {
                                reorderData[i, 4] = copyData[i, j];
                            }
                            // SET TITLE VALUE
                            else if (copyHeader[j].Contains("title"))
                            {
                                reorderData[i, 11] = copyData[i, j];
                            }
                            else
                            {

                            }
                        }
                    }

                    string thisfile = String.Empty;
                    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                    SaveFileDialog fileStream = new SaveFileDialog();
                    fileStream.FileName = "resellerupload.csv";
                    fileStream.DefaultExt = ".csv";
                    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                    DialogResult result = fileStream.ShowDialog();
                    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                    if (result == DialogResult.OK)
                    {
                        thisfile = fileStream.FileName;
                    }
                    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                    // CALL CREATE CLASS'S CSV-MAKING METHOD
                    makeFile.ResellerUpload();
                    OpenFolder(thisfile);
                }
                catch
                {
                    MessageBox.Show("Unable to export data to file properly.");
                }
            }
        }

        // ****************************************************************************************************************************************************************************************************
        // ****************************************************************************************************************************************************************************************************
        // ********------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------********
        // ********------------------------------------------------------------------------P A Y L O A D   S T U F F-----------------------------------------------------------------------------------********
        // ********------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------********
        // ****************************************************************************************************************************************************************************************************
        // ****************************************************************************************************************************************************************************************************


        private void btnCopyClipboard_Click(object sender, EventArgs e)
        {
            if (cboPayloadPicker.SelectedIndex != -1 && (radInsight.Checked || radKnowBe4.Checked))
            {
                FileStream payLoad = new FileStream(cboPayloadPicker.SelectedValue.ToString(), FileMode.Open, FileAccess.Read);
                streamer = new StreamReader(payLoad);
                itemText = streamer.ReadToEnd().ToString();
                Clipboard.SetText(itemText);
            }
        }

        private void radInsight_CheckedChanged(object sender, EventArgs e)
        {
            btnCopyClipboard.Enabled = true;
            cboPayloadPicker.Enabled = true;
            btnCopyClipboard.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
            btnCopyClipboard.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
            btnCopyClipboard.ForeColor = System.Drawing.Color.White;
            cboPayloadPicker.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
            cboPayloadPicker.ForeColor = System.Drawing.Color.White;
            if (radInsight.Checked)
            {
                string key = string.Empty;
                string value = string.Empty;
                int length = path.Length;

                List<KeyValuePair<string, string>> data = new List<KeyValuePair<string, string>>();
                path = @"insightpayloads\";
                string[] files = Directory.GetFiles(path);
                foreach (var file in files)
                {
                    key = file.ToString();
                    string[] temp = file.ToString().Split('\\');
                    temp = temp[1].Split('.');
                    value = temp[0];
                    data.Add(new KeyValuePair<string, string>(key, value));
                }
                cboPayloadPicker.DataSource = null;
                cboPayloadPicker.Items.Clear();
                cboPayloadPicker.DataSource = new BindingSource(data, null);
                cboPayloadPicker.DisplayMember = "Value";
                cboPayloadPicker.ValueMember = "Key";
            }
        }

        private void radKnowBe4_CheckedChanged(object sender, EventArgs e)
        {
            btnCopyClipboard.Enabled = true;
            cboPayloadPicker.Enabled = true;
            btnCopyClipboard.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
            btnCopyClipboard.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
            btnCopyClipboard.ForeColor = System.Drawing.Color.White;
            cboPayloadPicker.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
            cboPayloadPicker.ForeColor = System.Drawing.Color.White;
            if (radKnowBe4.Checked)
            {
                string key = string.Empty;
                string value = string.Empty;
                int length = path.Length;

                List<KeyValuePair<string, string>> data = new List<KeyValuePair<string, string>>();
                path = @"knowbe4payloads\";
                string[] files = Directory.GetFiles(path);
                foreach (var file in files)
                {
                    key = file.ToString();
                    string[] temp = file.ToString().Split('\\');
                    temp = temp[1].Split('.');
                    value = temp[0];
                    data.Add(new KeyValuePair<string, string>(key, value));
                }
                cboPayloadPicker.DataSource = null;
                cboPayloadPicker.Items.Clear();
                cboPayloadPicker.DataSource = new BindingSource(data, null);
                cboPayloadPicker.DisplayMember = "Value";
                cboPayloadPicker.ValueMember = "Key";
            }
        }

        private void btnAddNew_Click(object sender, EventArgs e)
        {
            AddPayload addPayload = new AddPayload();
            addPayload.ShowDialog();
        }


        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                            PHONE CALL TAB                                                                    ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################

        //method that will create a new Excel Sheet that will be used when making calls to clients 
        private void btnCreateCallList_Click(object sender, EventArgs e)
        {
            NewCallList callList = new NewCallList();
            callList.dataTable = dataTable;
            callList.ShowDialog();
        }

        private void btnMakeCalls_Click(object sender, EventArgs e)
        {
            MakeCalls calls = new MakeCalls();
            if (calls.failed == false)
            {
                calls.ShowDialog();
                calls.driver.Quit();
            }
        }

        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                            CREATE REPORT TAB                                                                    ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################

        private void enableReportShell()
        {
            radPhone.Enabled = (txtClient.Text != "" && txtPOC.Text != "");
            radEmail.Enabled = (txtClient.Text != "" && txtPOC.Text != "" && DateTime.Compare(dateTimePicker1.Value.Date, new DateTime(2019, 1, 1)) > 0);
            radBoth.Enabled = (txtClient.Text != "" && txtPOC.Text != "" && DateTime.Compare(dateTimePicker1.Value.Date, new DateTime(2019, 1, 1)) > 0);
            btnReportShell.Enabled = (txtClient.Text != "" && txtPOC.Text != "" && (radEmail.Checked | radPhone.Checked | radBoth.Checked));
            if (btnReportShell.Enabled)
            {
                btnReportShell.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                btnReportShell.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                btnReportShell.ForeColor = System.Drawing.Color.White;
            }
            else
            {
                btnReportShell.BackColor = System.Drawing.Color.Gray;
                btnReportShell.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
                btnReportShell.ForeColor = System.Drawing.Color.LightGray;
            }

            if (!radPhone.Enabled)
            {
                radPhone.Checked = false;
            }
            if (!radEmail.Enabled)
            {
                radEmail.Checked = false;
            }
            if (!radBoth.Enabled)
            {
                radBoth.Checked = false;
            }

        }

        private void txtClient_TextChanged(object sender, EventArgs e)
        {
            enableReportShell();
        }

        private void txtPOC_TextChanged(object sender, EventArgs e)
        {
            enableReportShell();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            enableReportShell();
        }

        private void radEmail_Click(object sender, EventArgs e)
        {
            enableReportShell();
        }

        private void radPhone_Click(object sender, EventArgs e)
        {
            enableReportShell();
        }

        private void radBoth_Click(object sender, EventArgs e)
        {
            enableReportShell();
        }

        private void btnReportShell_Click(object sender, EventArgs e)
        {
            this.Location = new Point(0, 0);
            if (radPhone.Checked)
            {
                hideReportTab();
                setLoadingLabel("Open the Call List File");
                OpenFileDialog openCallList = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true, Title = "Pick the Call List .xlsx file that was made by the RSE Tool." };
                DialogResult result = openCallList.ShowDialog();
                if (result == DialogResult.OK)
                {

                    contactpath = openCallList.FileName;

                    if (!createVishingReport.IsBusy)
                    {
                        xlApp = new Excel.Application();
                        createVishingReport.RunWorkerAsync();
                    }
                }
                else
                {
                    showReportTab();
                    return;
                }
            }
            else if (radEmail.Checked) //------------------------------------------------------------------------------------------------------------------------------------------------------------
            {
                hideReportTab();
                setLoadingLabel("Open the Phishing Campaign Results File");
                OpenFileDialog openCampaignResults = new OpenFileDialog() { Filter = "Comma Seperated Values|*.csv", ValidateNames = true, Title = "Pick the Email Campaign Results .csv file for " + txtClient.Text + "'s RSE." };
                DialogResult result = openCampaignResults.ShowDialog();
                if (result == DialogResult.OK)
                {
                    contactpath = openCampaignResults.FileName;

                    if (!createPhishingReport.IsBusy)
                    {
                        xlApp = new Excel.Application();
                        createPhishingReport.RunWorkerAsync();
                    }
                }
                else
                {
                    showReportTab();
                    return;
                }
            }
            else if (radBoth.Checked) //--------------------------------------------------------------------------------------------------------------------------------------------
            {
                hideReportTab();
                setLoadingLabel("Open the Call List and Phishing Campaign Results Files"); //**
                //open the file that contains the email campaign results 
                OpenFileDialog openCampaignResults = new OpenFileDialog() { Filter = "Comma Seperated Values|*.csv", ValidateNames = true, Title = "Pick the Email Campaign Results .csv file for " + txtClient.Text + "'s RSE." };
                DialogResult phishingResult = openCampaignResults.ShowDialog();
                //open the file that contains the phone call results 
                OpenFileDialog openCallList = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true, Title = "Pick the Call List .xlsx file that was made by the RSE Tool." };
                DialogResult vishingResult = openCallList.ShowDialog();

                if (phishingResult == DialogResult.OK & vishingResult == DialogResult.OK)
                {
                    contactpath = openCampaignResults.FileName + "?" + openCallList.FileName;

                    if (!createBothReport.IsBusy)
                    {
                        xlApp = new Excel.Application();
                        createBothReport.RunWorkerAsync();
                    }
                }
                else
                {
                    showReportTab();
                    return;
                }
            }
        }

        private void incrementProgressBar(int addNum, BackgroundWorker backgroundworker) //**
        {
            if (addNum > 1)
            {
                for (int i = 0; i < addNum; i++)
                {
                    progressBarPercentage += 1;  //**
                    backgroundworker.ReportProgress(progressBarPercentage); //**
                    Thread.Sleep(100);
                }
            }
            else
            {
                progressBarPercentage += addNum;  //**
                backgroundworker.ReportProgress(progressBarPercentage); //**
            }
        }

        private void hideReportTab()
        {
            label2.Visible = false;
            label5.Visible = false;
            label7.Visible = false; //**
            //label7.Location = new Point(133, 73); //**
            label6.Visible = false;
            txtClient.Visible = false;
            txtPOC.Visible = false;
            dateTimePicker1.Visible = false;
            radEmail.Visible = false;
            radPhone.Visible = false;
            radBoth.Visible = false;
            btnReportShell.Visible = false;
            progressBar1.Visible = true;
            labelPercentage.Visible = true;
            labelCurrentAction.Visible = true;
        }

        private void showReportTab()
        {
            progressBar1.Visible = false;
            labelPercentage.Visible = false;
            labelCurrentAction.Visible = false;
            txtClient.Clear();
            txtPOC.Clear();
            dateTimePicker1.Value = new DateTime(2019, 1, 1);
            label2.Visible = true;
            label5.Visible = true;
            //label7.Text = "Point of Contact*"; //**
            //label7.Location = new Point(4, 73); //**
            label7.Visible = true; //**
            label6.Visible = true;
            txtClient.Visible = true;
            txtPOC.Visible = true;
            dateTimePicker1.Visible = true;
            radEmail.Visible = true;
            radPhone.Visible = true;
            radBoth.Visible = true;
            btnReportShell.Visible = true;
        }

        private void setLoadingLabel(string text) //**
        {
            labelCurrentAction.Text = text;
            labelCurrentAction.Location = new Point((397 - labelCurrentAction.Width) / 2, 73);
        }

        private void createVishingReport_DoWork(object sender, DoWorkEventArgs e)
        {
            incrementProgressBar(12, createVishingReport);
            calculateVishingResults(createVishingReport, contactpath);
            if (createVishingReport.CancellationPending == true)
            {
                e.Cancel = true;
                return;
            }
            incrementProgressBar(12, createVishingReport);

            //---------------------------------------------------------------- Specific to Vishing Campaign --------------------------------------------------------------------------------
            wdApp = new Word.Application();
            reportDoc = wdApp.Documents.Open("C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Report Template - Vishing.docx", ReadOnly: false);

            incrementProgressBar(5, createVishingReport);

            //setLoadingLabel("Updating Content Control fields");
            reportDoc.ContentControls[1].Range.Text = txtClient.Text.ToString(); //Client's Name
            incrementProgressBar(3, createVishingReport);
            reportDoc.ContentControls[4].Range.Text = txtPOC.Text; //Contact's Name
            incrementProgressBar(3, createVishingReport);
            reportDoc.ContentControls[6].Range.Text = (vishingPassedCount + vishingFailedCount + vishingVoicemailCount).ToString(); //Total Calls
            incrementProgressBar(3, createVishingReport);
            reportDoc.ContentControls[7].Range.Text = vishingPassedCount.ToString(); //Uncompromised
            incrementProgressBar(3, createVishingReport);
            reportDoc.ContentControls[8].Range.Text = vishingFailedCount.ToString(); //Compromised
            incrementProgressBar(3, createVishingReport);
            reportDoc.ContentControls[9].Range.Text = vishingVoicemailCount.ToString(); //Unanswered
            incrementProgressBar(3, createVishingReport);
            if (vishingFailedCount > 0)
            {
                reportDoc.ContentControls[10].DropdownListEntries[2].Select();
            }
            else
            {
                reportDoc.ContentControls[10].DropdownListEntries[1].Select();
            }
            incrementProgressBar(4, createVishingReport);

            Word.Chart vishingChart = reportDoc.Shapes[3].Chart;
            Excel.Workbook vishingChartWB = vishingChart.ChartData.Workbook;
            Excel.Worksheet vishingChartWS = vishingChartWB.Worksheets[1];
            vishingChartWS.Range["B2"].Value = vishingPassedCount; //Passed Value
            vishingChartWS.Range["B3"].Value = vishingFailedCount; //Failed Value
            vishingChartWS.Range["B4"].Value = vishingVoicemailCount; //Did not answer Value
            vishingChartWB.Close();

            incrementProgressBar(5, createVishingReport);

            //setLoadingLabel("Copying Vishing Notes to Report");
            vishingNotesWS.Range["A1", "D" + vishingNotesMaxRow].Copy();
            try
            {
                reportDoc.Paragraphs[43].Range.Paste();
                reportDoc.Tables[1].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }
            catch
            {
                Console.WriteLine("Vishing Paste Error, but could have work. The program sometime throws this error even though the vishing notes table is in the report.");
            }

            incrementProgressBar(5, createVishingReport);

            int currentTable = 1;
            for (int i = 1; i <= reportDoc.Tables[currentTable].Rows.Count; i = i + 4)
            {
                if (reportDoc.Tables[currentTable].Rows[i].Range.Information[Word.WdInformation.wdActiveEndPageNumber] != reportDoc.Tables[currentTable].Rows[i + 3].Range.Information[Word.WdInformation.wdActiveEndPageNumber])
                {
                    reportDoc.Tables[currentTable].Rows[i].Range.InsertBreak(Word.WdBreakType.wdPageBreak);
                    currentTable++;
                    i = -3;
                }

            }

            incrementProgressBar(8, createVishingReport);

            for (int i = 1; i <= reportDoc.Tables.Count; i++)
            {
                reportDoc.Tables[i].Rows[1].Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }

            incrementProgressBar(5, createVishingReport);
        }

        private void createVishingReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            labelPercentage.Text = e.ProgressPercentage.ToString() + "%";

            if (e.ProgressPercentage < 13)
            {
                setLoadingLabel("Opening Vishing Call List");
            }
            else if (e.ProgressPercentage < 25)
            {
                setLoadingLabel("Verifying Call List");
            }
            else if (e.ProgressPercentage < 32)
            {
                setLoadingLabel("Opening Vishing Notes Template");
            }
            else if (e.ProgressPercentage < 51)
            {
                setLoadingLabel("Calculating Vishing Results");
            }
            else if (e.ProgressPercentage < 65)
            {
                setLoadingLabel("Opening Vishing Report Template");
            }
            else if (e.ProgressPercentage < 83)
            {
                setLoadingLabel("Updating Template Data");
            }
            else if (e.ProgressPercentage < 101)
            {
                setLoadingLabel("Inserting Vishing Notes Table");
            }
        }

        private void createVishingReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                setLoadingLabel("Cancelling Vishing Report...");
                createVishingReport.Dispose();
                showReportTab();
            }
            else if (e.Error != null)
            {
                labelCurrentAction.Text = "Error: " + e.Error.Message;
            }
            else
            {
                xlApp.Visible = true;
                setLoadingLabel("Save the Vishing Notes File");
                this.BringToFront();
                int currentYear = DateTime.Now.Year;
                SaveFileDialog vishingNotesFileStream = new SaveFileDialog();
                vishingNotesFileStream.Title = "Vishing Notes/Phone Engagement Detail Table Excel File Save as";
                vishingNotesFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Vishing Notes.xlsx";
                vishingNotesFileStream.DefaultExt = ".xlsx";
                vishingNotesFileStream.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                DialogResult vishingNotesResult = vishingNotesFileStream.ShowDialog(); //**
                if (vishingNotesResult == DialogResult.OK) //**
                { //**
                    xlApp.DisplayAlerts = false;
                    string fileName = vishingNotesFileStream.FileName; //**
                    vishingNotesWB.SaveAs(fileName); //** this line in neccesary
                } //**
                setLoadingLabel("Exiting Excel...");
                xlApp.DisplayAlerts = false;
                vishingNotesWB.Close();
                xlApp.Quit();

                setLoadingLabel("Save the Vishing Report");
                this.BringToFront();
                wdApp.Visible = true;
                wdApp.Activate();
                reportDoc.Activate();
                SaveFileDialog vishingReportFileStream = new SaveFileDialog();
                vishingReportFileStream.Title = "Vishing Report File Save as";
                vishingReportFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Vishing Report.xlsx";
                vishingReportFileStream.DefaultExt = ".docx";
                vishingReportFileStream.Filter = "Word Document File (.docx)|*.docx";
                DialogResult vishingReportResult = vishingReportFileStream.ShowDialog();
                if (vishingReportResult == DialogResult.OK)
                {
                    string fileName = vishingReportFileStream.FileName;
                    reportDoc.SaveAs(fileName);
                }
                setLoadingLabel("Exiting Word...");
                reportDoc.Close();
                wdApp.Quit();
                setLoadingLabel("Success!");

                showReportTab();
            }
        }

        private void createPhishingReport_DoWork(object sender, DoWorkEventArgs e)
        {
            incrementProgressBar(12, createPhishingReport);
            calculatePhishingResults(createPhishingReport, contactpath);
            incrementProgressBar(13, createPhishingReport);

            //-------------------------------------------------------- Specific to Phishing Campaigns ------------------------------------------------------------
            //setLoadingLabel("Starting Word");
            wdApp = new Word.Application();
            reportDoc = wdApp.Documents.Open("C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Report Template - Phishing.docx", ReadOnly: false);

            incrementProgressBar(4, createPhishingReport);

            //setLoadingLabel("Updating Content Control Fields");
            reportDoc.ContentControls[1].Range.Text = txtClient.Text.ToString(); //Client's Name
            incrementProgressBar(4, createPhishingReport);
            reportDoc.ContentControls[12].Range.Text = txtPOC.Text; //Contact's Name
            reportDoc.ContentControls[4].Range.Text = (phishingTotalEmails).ToString(); //Total Emails
            incrementProgressBar(4, createPhishingReport);
            reportDoc.ContentControls[5].Range.Text = (phishingTotalEmails - phishingFailedCount).ToString(); //Passed Emails
            reportDoc.ContentControls[6].Range.Text = phishingFailedCount.ToString(); //Failed Emails
            incrementProgressBar(4, createPhishingReport);
            reportDoc.ContentControls[8].Range.Text = phishingOpenedCount.ToString(); //Opened Emails
            reportDoc.ContentControls[10].Range.Text = dateTimePicker1.Value.ToShortDateString(); //Phishing Testing Email Start Date
            incrementProgressBar(4, createPhishingReport);
            if (phishingFailedCount > 0)
            {
                reportDoc.ContentControls[7].DropdownListEntries[2].Select(); //an unsuccesful
            }
            else
            {
                reportDoc.ContentControls[7].DropdownListEntries[1].Select(); //a successful
            }

            incrementProgressBar(4, createPhishingReport);

            //setLoadingLabel("Updating Vishing Charts Data");
            Word.Chart vishingChart = reportDoc.Shapes[3].Chart;
            Excel.Workbook vishingChartWB = vishingChart.ChartData.Workbook;
            Excel.Worksheet vishingChartWS = vishingChartWB.Worksheets[1];
            vishingChartWS.Range["B2"].Value = (phishingTotalEmails - phishingOpenedCount); //Not Opened Emails
            vishingChartWS.Range["B3"].Value = phishingOpenedCount; //Opened Emails

            incrementProgressBar(8, createPhishingReport);

            vishingChart = reportDoc.Shapes[4].Chart;
            vishingChartWB = vishingChart.ChartData.Workbook;
            vishingChartWS = vishingChartWB.Worksheets[1];
            vishingChartWS.Range["B2"].Value = (phishingTotalEmails - phishingFailedCount); //Passed Emails
            vishingChartWS.Range["B3"].Value = phishingFailedCount; //Failed Emails

            incrementProgressBar(8, createPhishingReport);


            //setLoadingLabel("Pasting Phishing Email Engagement Table");
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Copy();
            reportDoc.Paragraphs[42].Range.Paste();
            reportDoc.Tables[1].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;

            incrementProgressBar(10, createPhishingReport);
        }

        private void createPhishingReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            labelPercentage.Text = e.ProgressPercentage.ToString() + "%";

            if (e.ProgressPercentage < 12)
            {
                setLoadingLabel("Opening Phishing Campaign File");
            }
            else if (e.ProgressPercentage < 21)
            {
                setLoadingLabel("Verifying Phishing Campaign File");
            }
            else if (e.ProgressPercentage < 26)
            {
                setLoadingLabel("Deleting Unnecessary Columns");
            }
            else if (e.ProgressPercentage < 38)
            {
                setLoadingLabel("Calculating Phishing Results");
            }
            else if (e.ProgressPercentage < 45)
            {
                setLoadingLabel("Adding Borders to Phishing Results Table");
            }
            else if (e.ProgressPercentage < 55)
            {
                setLoadingLabel("Opening Phishing Report Template");
            }
            else if (e.ProgressPercentage < 75)
            {
                setLoadingLabel("Updating Template Data");
            }
            else if (e.ProgressPercentage < 91)
            {
                setLoadingLabel("Updating Template Charts");
            }
            else if (e.ProgressPercentage < 101)
            {
                setLoadingLabel("Inserting Phishing Campaign Table");
            }
        }

        private void createPhishingReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                labelCurrentAction.Text = "This should never happen. backgroundWorker 1 was cancelled.";
            }
            else if (e.Error != null)
            {
                labelCurrentAction.Text = "Error: " + e.Error.Message;
            }
            else
            {
                setLoadingLabel("Exiting Excel...");
                xlApp.DisplayAlerts = false;
                phishingResultsWB.Close();
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                setLoadingLabel("Save the Phishing Report");
                this.BringToFront();
                wdApp.Visible = true;
                wdApp.Activate();
                reportDoc.Activate();
                SaveFileDialog phishingReportFileStream = new SaveFileDialog();
                phishingReportFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Phishing Report.xlsx";
                phishingReportFileStream.DefaultExt = ".docx";
                phishingReportFileStream.Filter = "Word Document File (.docx)|*.docx";
                DialogResult phishingReportResult = phishingReportFileStream.ShowDialog();
                if (phishingReportResult == DialogResult.OK)
                {
                    string fileName = phishingReportFileStream.FileName;
                    reportDoc.SaveAs(fileName);
                }

                setLoadingLabel("Exiting Word...");
                reportDoc.Close();
                wdApp.Quit();
                setLoadingLabel("Success!");

                showReportTab();
            }
        }

        private void createBothReport_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] splitter = contactpath.Split('?');

            //---------------------------------------------------------------- Email Calculations -------------------------------------------------------------------------------------

            calculatePhishingResults(createBothReport, splitter[0]);
            string emailResultRange = "A1:F" + phishingResultsMaxRow.ToString();
            //--------------------------------------------------------------------- Phone Call Calculations -------------------------------------------------------------------------------------------------

            calculateVishingResults(createBothReport, splitter[1]);
            string vishingNotesRange = "A1:D" + vishingNotesMaxRow.ToString();


            //------------------------------------------------------------------ Specific to Phishing and Vishing  -----------------------------------------------------------------------
            //setLoadingLabel("Starting Word");
            wdApp = new Word.Application();
            reportDoc = wdApp.Documents.Open("C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Report Template - Phishing and Vishing.docx", ReadOnly: false);

            incrementProgressBar(2, createBothReport);

            //setLoadingLabel("Updating the Content Control fields");
            reportDoc.ContentControls[1].Range.Text = txtClient.Text.ToString(); //Client's Name
            reportDoc.ContentControls[4].Range.Text = txtPOC.Text; //Contact's Name
            incrementProgressBar(2, createBothReport);
            reportDoc.ContentControls[6].Range.Text = (vishingPassedCount + vishingFailedCount + vishingVoicemailCount).ToString(); //Total Calls
            reportDoc.ContentControls[7].Range.Text = vishingPassedCount.ToString(); //Uncompromised
            incrementProgressBar(2, createBothReport);
            reportDoc.ContentControls[8].Range.Text = vishingFailedCount.ToString(); //Compromised
            reportDoc.ContentControls[9].Range.Text = vishingVoicemailCount.ToString(); //Unanswered
            incrementProgressBar(2, createBothReport);
            if (vishingFailedCount > 0) //Choose an item for Phone Calls
            {
                reportDoc.ContentControls[10].DropdownListEntries[2].Select(); //an unsuccesful
            }
            else
            {
                reportDoc.ContentControls[10].DropdownListEntries[1].Select(); //a successful
            }
            incrementProgressBar(2, createBothReport);
            reportDoc.ContentControls[12].Range.Text = (phishingTotalEmails).ToString(); //Total Emails
            reportDoc.ContentControls[13].Range.Text = (phishingTotalEmails - phishingFailedCount).ToString(); //Passed Emails
            incrementProgressBar(2, createBothReport);
            reportDoc.ContentControls[14].Range.Text = phishingFailedCount.ToString(); //Failed Emails
            reportDoc.ContentControls[16].Range.Text = phishingOpenedCount.ToString(); //Opened Emails
            incrementProgressBar(2, createBothReport);
            reportDoc.ContentControls[18].Range.Text = dateTimePicker1.Value.ToShortDateString(); //Phishing Test Start Date
            if (phishingFailedCount > 0)
            {
                reportDoc.ContentControls[15].DropdownListEntries[2].Select(); //an unsuccesful
            }
            else
            {
                reportDoc.ContentControls[15].DropdownListEntries[1].Select(); //a successful
            }
            incrementProgressBar(2, createBothReport);

            //setLoadingLabel("Updating Vishing and Phishing Charts' data");
            Word.Chart vishingChart = reportDoc.Shapes[5].Chart;
            Excel.Workbook vishingChartWB = vishingChart.ChartData.Workbook;
            Excel.Worksheet vishingChartWS = vishingChartWB.Worksheets[1];
            vishingChartWS.Range["B2"].Value = vishingPassedCount; //passed calls
            vishingChartWS.Range["B3"].Value = vishingFailedCount; //failed calls 
            vishingChartWS.Range["B4"].Value = vishingVoicemailCount; //did not answer
            System.Threading.Thread.Sleep(2000);
            vishingChartWB.Close();

            incrementProgressBar(5, createBothReport);

            Word.Chart phishingOpenChart = reportDoc.Shapes[6].Chart;
            Excel.Workbook phishingOpenChartWB = phishingOpenChart.ChartData.Workbook;
            Excel.Worksheet phishingOpenChartWS = phishingOpenChartWB.Worksheets[1];
            phishingOpenChartWS.Range["B2"].Value = (phishingTotalEmails - phishingOpenedCount); //Not Opened Emails
            phishingOpenChartWS.Range["B3"].Value = phishingOpenedCount; //Opened Emails
            System.Threading.Thread.Sleep(2000);
            phishingOpenChartWB.Close();

            incrementProgressBar(3, createBothReport);

            Word.Chart phishingResultChart = reportDoc.Shapes[7].Chart;
            Excel.Workbook phishingResultChartWB = phishingResultChart.ChartData.Workbook;
            Excel.Worksheet phishingResultChartWS = phishingResultChartWB.Worksheets[1];
            phishingResultChartWS.Range["B2"].Value = (phishingTotalEmails - phishingFailedCount); //Passed Emails
            phishingResultChartWS.Range["B3"].Value = phishingFailedCount; //Failed Emails
            System.Threading.Thread.Sleep(2000);
            phishingResultChartWB.Close();

            incrementProgressBar(3, createBothReport);

            //setLoadingLabel("Pasting Vishing Notes Summary into Report");
            //paste Vishing Call Notes into Paragraph 72 //
            vishingNotesWS.Range[vishingNotesRange].Copy();
            System.Threading.Thread.Sleep(2000);
            try
            {
                reportDoc.Paragraphs[72].Range.Paste();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Thrown: Error pasting vishing notes but this could have worked. Check report document");
            }

            incrementProgressBar(3, createBothReport);

            //setLoadingLabel("Pasting Email Engagement Table into Report");
            //paste Email Results into Paragraph 70 //
            phishingResultsWS.Range[emailResultRange].Copy();
            System.Threading.Thread.Sleep(2000);
            reportDoc.Paragraphs[70].Range.Paste();
            xlApp.DisplayAlerts = false;
            phishingResultsWB.Close();
            xlApp.DisplayAlerts = true;

            incrementProgressBar(3, createBothReport);

            //setLoadingLabel("Formatting Report for Pretty Printing");
            for (int i = 1; i <= reportDoc.Tables.Count; i++)
            {
                //System.Threading.Thread.Sleep(1000);
                reportDoc.Tables[i].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }

            incrementProgressBar(3, createBothReport);

            int phoneEngagementDetailParagraph = 0;
            for (int i = 70; i < reportDoc.Paragraphs.Count; i++)
            {
                string style = ((Word.Style)reportDoc.Paragraphs[i].get_Style()).NameLocal;
                if (style.Contains("Heading"))
                {
                    phoneEngagementDetailParagraph = i;
                    break;
                }
            }

            incrementProgressBar(3, createBothReport);

            int emailTableLastRow = reportDoc.Tables[1].Rows.Count;
            //if "Phone Engagement Details" HEADER is on the same page as the last row of the Email Engagement Detail TABLE then insert a page break
            if (reportDoc.Paragraphs[phoneEngagementDetailParagraph].Range.Information[Word.WdInformation.wdActiveEndPageNumber] == reportDoc.Tables[1].Rows[emailTableLastRow].Range.Information[Word.WdInformation.wdActiveEndPageNumber])
            {
                reportDoc.Paragraphs.Add(reportDoc.Paragraphs[phoneEngagementDetailParagraph].Range);
                reportDoc.Paragraphs[phoneEngagementDetailParagraph].Range.InsertBreak(Word.WdBreakType.wdPageBreak);
                reportDoc.Paragraphs[phoneEngagementDetailParagraph].set_Style(reportDoc.Styles["Normal"]);
            }

            incrementProgressBar(3, createBothReport);

            int currentTable = 2;
            for (int i = 1; i <= reportDoc.Tables[currentTable].Rows.Count; i = i + 4)
            {
                if (reportDoc.Tables[currentTable].Rows[i].Range.Information[Word.WdInformation.wdActiveEndPageNumber] != reportDoc.Tables[currentTable].Rows[i + 3].Range.Information[Word.WdInformation.wdActiveEndPageNumber])
                {
                    reportDoc.Tables[currentTable].Rows[i].Range.InsertBreak(Word.WdBreakType.wdPageBreak);
                    currentTable++;
                    i = -3;
                }

            }

            incrementProgressBar(3, createBothReport);

            for (int i = 1; i <= reportDoc.Tables.Count; i++)
            {
                reportDoc.Tables[i].Rows[1].Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }

            incrementProgressBar(5, createBothReport);
        }

        private void createBothReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            labelPercentage.Text = e.ProgressPercentage.ToString() + "%";

            if (e.ProgressPercentage < 5)
            {
                setLoadingLabel("Opening Phishing Campaign File");
            }
            else if (e.ProgressPercentage < 12)
            {
                setLoadingLabel("Deleting Unnecessary Columns");
            }
            else if (e.ProgressPercentage < 20)
            {
                setLoadingLabel("Calculating Phishing Results");
            }
            else if (e.ProgressPercentage < 26)
            {
                setLoadingLabel("Adding Borders to Phishing Results Table");
            }
            else if (e.ProgressPercentage < 30)
            {
                setLoadingLabel("Opening Vishing Call List");
            }
            else if (e.ProgressPercentage < 40)
            {
                setLoadingLabel("Verifying Call List");
            }
            else if (e.ProgressPercentage < 45)
            {
                setLoadingLabel("Opening Vishing Notes Template");
            }
            else if (e.ProgressPercentage < 51)
            {
                setLoadingLabel("Calculating Vishing Results");
            }
            else if (e.ProgressPercentage < 56)
            {
                setLoadingLabel("Opening Phishing/Vishing Report Template");
            }
            else if (e.ProgressPercentage < 67)
            {
                setLoadingLabel("Updating Template Data");
            }
            else if (e.ProgressPercentage < 78)
            {
                setLoadingLabel("Updating Template Charts");
            }
            else if (e.ProgressPercentage < 84)
            {
                setLoadingLabel("Inserting Phishing and Vishing Notes Tables");
            }
            else if (e.ProgressPercentage < 101)
            {
                setLoadingLabel("Formatting Tables");
            }
        }

        private void createBothReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                labelCurrentAction.Text = "This should never happen. backgroundWorker 1 was cancelled.";
            }
            else if (e.Error != null)
            {
                labelCurrentAction.Text = "Error: " + e.Error.Message;
            }
            else
            {
                setLoadingLabel("Save the Vishing Notes File");
                xlApp.Visible = true;
                vishingNotesWB.Activate();
                int currentYear = DateTime.Now.Year;
                SaveFileDialog vishingNotesFileStream = new SaveFileDialog();
                vishingNotesFileStream.Title = "Vishing Notes/Phone Engagement Detail Table File Save as";
                vishingNotesFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Vishing Notes.xlsx";
                vishingNotesFileStream.DefaultExt = ".xlsx";
                vishingNotesFileStream.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                DialogResult vishingNotesResult = vishingNotesFileStream.ShowDialog();
                if (vishingNotesResult == DialogResult.OK)
                {
                    string fileName = vishingNotesFileStream.FileName;
                    vishingNotesWB.SaveAs(fileName);
                }
                setLoadingLabel("Exiting Excel...");
                vishingNotesWB.Close();
                xlApp.Quit();


                setLoadingLabel("Save the Phishing and Vishing Report");
                wdApp.Visible = true;
                SaveFileDialog phishingAndVishingReportFileStream = new SaveFileDialog();
                phishingAndVishingReportFileStream.Title = "Phishing and Vishing Report File Save as";
                phishingAndVishingReportFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Phishing and Vishing Report.docx";
                phishingAndVishingReportFileStream.DefaultExt = ".docx";
                phishingAndVishingReportFileStream.Filter = "Word Document File (.docx)|*.docx";
                DialogResult phishingAndVishingReportResult = phishingAndVishingReportFileStream.ShowDialog();
                if (phishingAndVishingReportResult == DialogResult.OK)
                {
                    string fileName = phishingAndVishingReportFileStream.FileName;
                    reportDoc.SaveAs(fileName);
                }
                setLoadingLabel("Exiting Word...");
                reportDoc.Close();
                wdApp.Quit();

                showReportTab();
            }
        }

        private void calculateVishingResults(BackgroundWorker currentBackgroundWorker, string callListFileName)
        {
            dataTable = excelSheetToDataTable(callListFileName, false); //** changed callListPath to contactpath to callListFileName

            incrementProgressBar(2, currentBackgroundWorker);

            //check that the Call List was made by the RSE Tool so that there are no errors later 
            //setLoadingLabel("Verifying Call List");
            try
            {
                if (!dataTable.Rows[0][0].Equals("Calling As") & !dataTable.Rows[1][0].Equals("Phone # Displayed") & !dataTable.Rows[2][0].Equals("Name Drop")
                    & !dataTable.Rows[3][0].Equals("Engagements Needed") & !dataTable.Rows[4][0].Equals("Engagements per Day") & !dataTable.Rows[5][0].Equals("Current Engagements")
                    & !dataTable.Rows[6][0].Equals("Business Hours"))
                {
                    MessageBox.Show("Please use a Call List that was created by the RSE Tool.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dataTable = null;
                    xlApp.ActiveWorkbook.Close();
                    xlApp.Quit();
                    currentBackgroundWorker.CancelAsync(); //**
                    //exit method to return to the Main Class
                    return;
                }
            } catch (Exception ex)
            {
                MessageBox.Show("Please use a Call List that was created by the RSE Tool.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dataTable = null;
                xlApp.Quit();
                currentBackgroundWorker.CancelAsync(); //**
                //exit method to return to the Main Class
                return;
            }
            incrementProgressBar(2, currentBackgroundWorker);

            dataTable.Rows[0].Delete();
            dataTable.Rows[1].Delete();
            incrementProgressBar(2, currentBackgroundWorker);
            dataTable.Rows[2].Delete();
            dataTable.Rows[3].Delete();
            incrementProgressBar(2, currentBackgroundWorker);
            dataTable.Rows[4].Delete();
            dataTable.Rows[5].Delete();
            incrementProgressBar(2, currentBackgroundWorker);
            dataTable.Rows[6].Delete();
            dataTable.Rows[7].Delete();
            dataTable.AcceptChanges();
            incrementProgressBar(2, currentBackgroundWorker);

            //setLoadingLabel("Starting Excel");
            incrementProgressBar(2, currentBackgroundWorker);
            //xlApp.Visible = true;
            string path = "C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Vishing Notes Template.xlsx";
            vishingNotesWB = xlApp.Workbooks.Open(path, ReadOnly: false);
            vishingNotesWS = vishingNotesWB.Worksheets[1];
            Excel.Worksheet tempWS = vishingNotesWB.Worksheets[2];
            incrementProgressBar(3, currentBackgroundWorker);

            int dtMaxRow = dataTable.Rows.Count;
            int dtResultCol = 0;
            bool hasExtension = false;
            int dtExtensionCol = 0;
            //find the result column in the datatable 
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (dataTable.Rows[0][i].Equals("Result"))
                {
                    dtResultCol = i;
                }

                if (dataTable.Rows[0][i].Equals("Extension"))
                {
                    hasExtension = true;
                    dtExtensionCol = i;
                }
            }
            Console.WriteLine("DataTable Result Column: " + dtResultCol);

            incrementProgressBar(2, currentBackgroundWorker);

            //setLoadingLabel("Creating Vishing Notes summary");
            long j = 2;
            string tempDate = null;
            List<string> dates = new List<string>();
            string tempDescrip = null;
            incrementProgressBar(2, currentBackgroundWorker);
            for (int i = 1; i < dtMaxRow; i++) //i = current DataTable row
            {
                tempDate = null;
                dates.Clear();
                tempDescrip = "temp";
                if (!dataTable.Rows[i][dtResultCol].Equals(DBNull.Value))
                {
                    if (dataTable.Rows[i][dtResultCol].Equals("PASSED"))
                    {
                        vishingPassedCount++;
                    }
                    if (dataTable.Rows[i][dtResultCol].Equals("FAILED"))
                    {
                        vishingFailedCount++;
                    }

                    vishingNotesWS.Range["A" + j.ToString()].Value = dataTable.Rows[i][dtResultCol]; //Final Result
                    vishingNotesWS.Range["B" + j.ToString()].Value = dataTable.Rows[i][0]; //Name
                    vishingNotesWS.Range["C" + j.ToString()].Value = dataTable.Rows[i][1]; //Phone 
                    if (hasExtension == true)
                    {
                        vishingNotesWS.Range["D" + j.ToString()].Value = dataTable.Rows[i][dtExtensionCol]; //Extension
                    }

                    for (int k = dataTable.Columns.Count - 1; k > dtResultCol; k--) //k = current DataTable Column to the right of Result Column
                    {
                        if (!dataTable.Rows[i][k].Equals(DBNull.Value))
                        {
                            if (dataTable.Rows[i][k].Equals("Voicemail"))
                            {
                                vishingVoicemailCount++;
                            }

                            tempDate = dataTable.Rows[0][k].ToString();
                            string[] tempArray = tempDate.Split(' ');
                            tempDate = tempArray[0];
                            dates.Add(tempDate);
                            if (tempDescrip.Equals("temp"))
                            {
                                tempDescrip = dataTable.Rows[i][k].ToString();
                            }
                            //break;
                        }
                    }

                    if (tempDescrip.Equals("Voicemail") || tempDescrip.Equals("temp"))
                    {
                        tempDescrip = null;
                    }

                    vishingNotesWS.Range["A" + (j + 1).ToString()].Value = "Dates:  " + String.Join(", ", dates); //Dates:
                    vishingNotesWS.Range["A" + (j + 2).ToString()].Value = "Description: " + tempDescrip; //Description
                    tempWS.Range["A1"].Value = "Description: " + tempDescrip; //Temp Description
                    double rowHeight = tempWS.Range["A1"].RowHeight;
                    vishingNotesWS.Range["A" + (j + 2).ToString()].RowHeight = rowHeight; //Description

                    vishingNotesMaxRow = j + 2;
                    j = j + 4;
                }
            }
            incrementProgressBar(4, currentBackgroundWorker);

            Console.WriteLine("Vishing Results");
            Console.WriteLine("Unanswered: " + vishingVoicemailCount);
            Console.WriteLine("Compromised: " + vishingFailedCount);
            Console.WriteLine("Uncompromised: " + vishingPassedCount);
        }

        private void calculatePhishingResults(BackgroundWorker currentBackgroundWorker, string phishingResultsFileName)
        {
            //setLoadingLabel("Starting Excel");
            phishingResultsWB = xlApp.Workbooks.Open(phishingResultsFileName, ReadOnly: false);
            phishingResultsWS = phishingResultsWB.Worksheets[1];

            incrementProgressBar(2, currentBackgroundWorker);

            Excel.Range last = phishingResultsWS.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            phishingResultsMaxRow = last.Row;

            incrementProgressBar(2, currentBackgroundWorker);

            //setLoadingLabel("Identifying Phising Platform");
            phishingTotalEmails = phishingResultsMaxRow - 1;

            if ("Email".Equals(Convert.ToString(phishingResultsWS.Range["A1"].Value2)) & "Clicked at".Equals(Convert.ToString(phishingResultsWS.Range["B1"].Value2)) & "Data entered at".Equals(Convert.ToString(phishingResultsWS.Range["C1"].Value2))
                & "Attachment opened at".Equals(Convert.ToString(phishingResultsWS.Range["D1"].Value2)) & "Macro enabled at".Equals(Convert.ToString(phishingResultsWS.Range["E1"].Value2)) & "Opened at".Equals(Convert.ToString(phishingResultsWS.Range["F1"].Value2))
                & "Delivered at".Equals(Convert.ToString(phishingResultsWS.Range["G1"].Value2)) & "Bounced at".Equals(Convert.ToString(phishingResultsWS.Range["H1"].Value2)) & "First Name".Equals(Convert.ToString(phishingResultsWS.Range["I1"].Value2))
                & "Last Name".Equals(Convert.ToString(phishingResultsWS.Range["J1"].Value2)) & "Job Title".Equals(Convert.ToString(phishingResultsWS.Range["K1"].Value2)) & "Group".Equals(Convert.ToString(phishingResultsWS.Range["L1"].Value2))
                & "Manager Name".Equals(Convert.ToString(phishingResultsWS.Range["M1"].Value2)) & "Manager Email".Equals(Convert.ToString(phishingResultsWS.Range["N1"].Value2)) & "Location".Equals(Convert.ToString(phishingResultsWS.Range["O1"].Value2))
                & "Division".Equals(Convert.ToString(phishingResultsWS.Range["P1"].Value2)) & "Employee number".Equals(Convert.ToString(phishingResultsWS.Range["Q1"].Value2)) & "IP Address".Equals(Convert.ToString(phishingResultsWS.Range["R1"].Value2))
                & "IP Location".Equals(Convert.ToString(phishingResultsWS.Range["S1"].Value2)) & "Browser".Equals(Convert.ToString(phishingResultsWS.Range["T1"].Value2)) & "Browser Version".Equals(Convert.ToString(phishingResultsWS.Range["U1"].Value2))
                & "Operating System".Equals(Convert.ToString(phishingResultsWS.Range["V1"].Value2)) & "Email Template".Equals(Convert.ToString(phishingResultsWS.Range["W1"].Value2)))
            {
                //delete columns: C - E; G - H (D & E); K - V
                Excel.Range range = phishingResultsWS.Range["C1", "E1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["D1", "E1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["F1", "Q1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["A1", "F1"];
                range.EntireColumn.AutoFit();

                incrementProgressBar(1, currentBackgroundWorker);

                for (int i = 2; i <= phishingResultsMaxRow; i++)
                {
                    if (Convert.ToString(phishingResultsWS.Range["B" + i].Value2) != null)
                    {
                        phishingFailedCount++;
                    }
                    if (Convert.ToString(phishingResultsWS.Range["C" + i].Value2) != null)
                    {
                        phishingOpenedCount++;
                    }
                }
                Console.WriteLine("Phishing Results");
                Console.WriteLine("total emails: " + phishingTotalEmails);
                Console.WriteLine("failed count: " + phishingFailedCount);
                Console.WriteLine("opened count: " + phishingOpenedCount);

                //copy columns A1 - F[maxRow] and paste into the report doc 
            }
            else if ("First Name".Equals(Convert.ToString(phishingResultsWS.Range["A1"].Value2)) & "Last Name".Equals(Convert.ToString(phishingResultsWS.Range["B1"].Value2)) & "Email Address".Equals(Convert.ToString(phishingResultsWS.Range["C1"].Value2))
              & "Group".Equals(Convert.ToString(phishingResultsWS.Range["D1"].Value2)) & "Viewed Images / Opened Email".Equals(Convert.ToString(phishingResultsWS.Range["E1"].Value2)) & "Passed".Equals(Convert.ToString(phishingResultsWS.Range["F1"].Value2))
              & "Failed".Equals(Convert.ToString(phishingResultsWS.Range["G1"].Value2)) & "Failed Date".Equals(Convert.ToString(phishingResultsWS.Range["H1"].Value2)) & "Campaign".Equals(Convert.ToString(phishingResultsWS.Range["I1"].Value2))
              & "campaign type".Equals(Convert.ToString(phishingResultsWS.Range["J1"].Value2)) & "Payload".Equals(Convert.ToString(phishingResultsWS.Range["K1"].Value2)) & "Payload Type".Equals(Convert.ToString(phishingResultsWS.Range["L1"].Value2))
              & "Group(s)".Equals(Convert.ToString(phishingResultsWS.Range["M1"].Value2)) & "Clicked Link".Equals(Convert.ToString(phishingResultsWS.Range["N1"].Value2)))
            {
                //delete columns: D, F, H - J, L - N
                Excel.Range range = phishingResultsWS.Range["D1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["E1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["F1", "H1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["G1", "I1"];
                range.EntireColumn.Delete(Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                incrementProgressBar(1, currentBackgroundWorker);

                range = phishingResultsWS.Range["A1", "F1"];
                range.EntireColumn.AutoFit();

                for (int i = 2; i <= phishingResultsMaxRow; i++)
                {
                    if (phishingResultsWS.Range["E" + i].Value.Equals("Yes"))
                    {
                        phishingFailedCount++;
                    }
                    if (phishingResultsWS.Range["D" + i].Value.Equals("Yes"))
                    {
                        phishingOpenedCount++;
                    }
                }
                Console.WriteLine("Phishing Results:");
                Console.WriteLine("total emails: " + phishingTotalEmails);
                Console.WriteLine("failed count: " + phishingFailedCount);
                Console.WriteLine("opened count: " + phishingOpenedCount);
            }
            else
            {
                MessageBox.Show("Please use an UNEDITED phishing campaign file (.csv) that was downloaded from Insight or KnowBe4.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                showReportTab(); //**
                //exit method to return to the Main Class
                return;
            }

            incrementProgressBar(5, currentBackgroundWorker);
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            incrementProgressBar(2, currentBackgroundWorker);
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            incrementProgressBar(2, currentBackgroundWorker);
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            incrementProgressBar(2, currentBackgroundWorker);
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            incrementProgressBar(2, currentBackgroundWorker);
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            incrementProgressBar(2, currentBackgroundWorker);
            phishingResultsWS.Range["A1", "F" + phishingResultsMaxRow].Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            incrementProgressBar(2, currentBackgroundWorker);
        }
    }
}

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
            radEmail.Enabled = (txtClient.Text != "" && txtPOC.Text != "" &&  DateTime.Compare(dateTimePicker1.Value.Date, new DateTime(2019,1,1)) > 0);
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
            if (radPhone.Checked)
            {
                hideReportTab();
                setLoadingLabel("Open the Call List File.");
                OpenFileDialog openCallList = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true, Title = "Pick the Call List .xlsx file that was made by the RSE Tool." };
                DialogResult result = openCallList.ShowDialog();
                if (result == DialogResult.OK)
                {

                    contactpath = openCallList.FileName;
                    dataTable = excelSheetToDataTable(contactpath, false);

                    //check that the Call List was made by the RSE Tool so that there are no errors later 
                    setLoadingLabel("Verifying Call List");
                    if (!dataTable.Rows[0][0].Equals("Calling As") & !dataTable.Rows[1][0].Equals("Phone # Displayed") & !dataTable.Rows[2][0].Equals("Name Drop")
                    & !dataTable.Rows[3][0].Equals("Engagements Needed") & !dataTable.Rows[4][0].Equals("Engagements per Day") & !dataTable.Rows[5][0].Equals("Current Engagements")
                     & !dataTable.Rows[6][0].Equals("Business Hours"))
                    {
                        MessageBox.Show("Please use a Call List that was created by the RSE Tool.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        showReportTab();
                        //exit method to return to the Main Class
                        return;
                    }

                    dataTable.Rows[0].Delete();
                    dataTable.Rows[1].Delete();
                    dataTable.Rows[2].Delete();
                    dataTable.Rows[3].Delete();
                    dataTable.Rows[4].Delete();
                    dataTable.Rows[5].Delete();
                    dataTable.Rows[6].Delete();
                    dataTable.Rows[7].Delete();
                    dataTable.AcceptChanges();

                    setLoadingLabel("Starting Excel");
                    Excel.Application xlApp = new Excel.Application();
                    //xlApp.Visible = true;
                    string path = "C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Vishing Notes Template.xlsx";
                    Excel.Workbook wb = xlApp.Workbooks.Open(path, ReadOnly: false);
                    Excel.Worksheet ws = wb.Worksheets[1];
                    Excel.Worksheet tempWS = wb.Worksheets[2];

                    int dtMaxRow = dataTable.Rows.Count;
                    int dtResultCol = 0;
                    Boolean hasExtension = false;
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

                    setLoadingLabel("Creating Vishing Notes summary");
                    long j = 2;
                    long maxRow = 0;
                    string tempDate = null;
                    List<string> dates = new List<string>();
                    string tempDescrip = null;
                    int voicemailCount = 0;
                    int passedCount = 0;
                    int failedCount = 0;
                    for (int i = 1; i < dtMaxRow; i++) //i = current DataTable row
                    {
                        tempDate = null;
                        dates.Clear();
                        tempDescrip = "temp";

                        if (!dataTable.Rows[i][dtResultCol].Equals(DBNull.Value))
                        {
                            if (dataTable.Rows[i][dtResultCol].Equals("PASSED"))
                            {
                                passedCount++;
                            }
                            if (dataTable.Rows[i][dtResultCol].Equals("FAILED"))
                            {
                                failedCount++;
                            }

                            ws.Range["A" + j.ToString()].Value = dataTable.Rows[i][dtResultCol]; //Final Result
                            ws.Range["B" + j.ToString()].Value = dataTable.Rows[i][0]; //Name
                            ws.Range["C" + j.ToString()].Value = dataTable.Rows[i][1]; //Phone 
                            if (hasExtension == true)
                            {
                                ws.Range["D" + j.ToString()].Value = dataTable.Rows[i][dtExtensionCol]; //Extension
                            }

                            for (int k = dataTable.Columns.Count - 1; k > dtResultCol; k--) //k = current DataTable Column to the right of Result Column
                            {
                                if (!dataTable.Rows[i][k].Equals(DBNull.Value))
                                {
                                    if (dataTable.Rows[i][k].Equals("Voicemail"))
                                    {
                                        voicemailCount++;
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

                            ws.Range["A" + (j + 1).ToString()].Value = "Dates:  " + String.Join(", ", dates); //Dates:
                            ws.Range["A" + (j + 2).ToString()].Value = "Description: " + tempDescrip; //Description
                            tempWS.Range["A1"].Value = "Description: " + tempDescrip; //Temp Description
                            double rowHeight = tempWS.Range["A1"].RowHeight;
                            ws.Range["A" + (j + 2).ToString()].RowHeight = rowHeight; //Description

                            maxRow = j + 2;
                            j = j + 4;
                        }
                    }

                    ws.Activate();

                    Console.WriteLine("Vishing Results");
                    Console.WriteLine("Unanswered: " + voicemailCount);
                    Console.WriteLine("Compromised: " + failedCount);
                    Console.WriteLine("Uncompromised: " + passedCount);


                    setLoadingLabel("Save the Vishing Notes File");
                    //excelApp.ScreenUpdating = false
                    int currentYear = DateTime.Now.Year;
                    SaveFileDialog vishingNotesFileStream = new SaveFileDialog();
                    vishingNotesFileStream.Title = "Vishing Notes/Phone Engagment Detail Tabke File Save as";
                    vishingNotesFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Vishing Notes.xlsx";
                    vishingNotesFileStream.DefaultExt = ".xlsx";
                    vishingNotesFileStream.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                    DialogResult vishingNotesResult = vishingNotesFileStream.ShowDialog();
                    if (vishingNotesResult == DialogResult.OK)
                    {
                        string fileName = vishingNotesFileStream.FileName;
                        wb.SaveAs(fileName);
                    }


                    //---------------------------------------------------------------- Specific to Vishing Campaign --------------------------------------------------------------------------------
                    setLoadingLabel("Starting Word");
                    Word.Application wordApp = new Word.Application();
                    //wordApp.Visible = true;
                    Word.Document reportDoc = wordApp.Documents.Open("C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Report Template - Vishing.docx", ReadOnly: false);

                    setLoadingLabel("Updating Content Control fields");
                    reportDoc.ContentControls[1].Range.Text = txtClient.Text.ToString(); //Client's Name
                    reportDoc.ContentControls[4].Range.Text = txtPOC.Text; //Contact's Name
                    reportDoc.ContentControls[6].Range.Text = (passedCount + failedCount + voicemailCount).ToString(); //Total Calls
                    reportDoc.ContentControls[7].Range.Text = passedCount.ToString(); //Uncompromised
                    reportDoc.ContentControls[8].Range.Text = failedCount.ToString(); //Compromised
                    reportDoc.ContentControls[9].Range.Text = voicemailCount.ToString(); //Unanswered
                    if (failedCount > 0)
                    {
                        reportDoc.ContentControls[10].DropdownListEntries[2].Select();
                    }
                    else
                    {
                        reportDoc.ContentControls[10].DropdownListEntries[1].Select();
                    }

                    Word.Chart vishingChart = reportDoc.Shapes[3].Chart;
                    Excel.Workbook vishingChartWB = vishingChart.ChartData.Workbook;
                    Excel.Worksheet vishingChartWS = vishingChartWB.Worksheets[1];
                    vishingChartWS.Range["B2"].Value = passedCount; //Passed Value
                    vishingChartWS.Range["B3"].Value = failedCount; //Failed Value
                    vishingChartWS.Range["B4"].Value = voicemailCount; //Did not answer Value
                    vishingChartWB.Close();

                    setLoadingLabel("Copying Vishing Notes to Report");
                    ws.Range["A1", "D" + maxRow].Copy();
                    try
                    {
                        reportDoc.Paragraphs[43].Range.Paste();
                        reportDoc.Tables[1].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    } catch
                    {
                        Console.WriteLine("Vishing Paste Error, but could have worked. ");
                    }

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

                    for (int i = 1; i <= reportDoc.Tables.Count; i++)
                    {
                        reportDoc.Tables[i].Rows[1].Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }

                    setLoadingLabel("Save the Vishing Report");
                    wordApp.Visible = true;
                    SaveFileDialog vishingReportFileStream = new SaveFileDialog();
                    vishingReportFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Vishing Report.xlsx";
                    vishingReportFileStream.DefaultExt = ".docx";
                    vishingReportFileStream.Filter = "Word Document File (.docx)|*.docx";
                    DialogResult vishingReportResult = vishingReportFileStream.ShowDialog();
                    if (vishingReportResult == DialogResult.OK)
                    {
                        string fileName = vishingReportFileStream.FileName;
                        reportDoc.SaveAs(fileName);
                    }

                    setLoadingLabel("Exiting Word and Excel");
                    Console.WriteLine("Label7 Width: " + label7.Width);
                    xlApp.DisplayAlerts = false;
                    wb.Close();
                    xlApp.Quit();
                    reportDoc.Close();
                    wordApp.Quit();
                    setLoadingLabel("Success!");
                    System.Threading.Thread.Sleep(3000);
                    showReportTab();

                }
            }
            else if (radEmail.Checked) //------------------------------------------------------------------------------------------------------------------------------------------------------------
            {
                hideReportTab();
                setLoadingLabel("Starting Excel");
                Excel.Application xlApp = new Excel.Application();
                OpenFileDialog openCampaignResults = new OpenFileDialog() { Filter = "Comma Seperated Values|*.csv", ValidateNames = true, Title = "Pick the Email Campaign Results .csv file for " + txtClient.Text + "'s RSE." };
                DialogResult result = openCampaignResults.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //xlApp.Visible = true;
                    contactpath = openCampaignResults.FileName;
                    Excel.Workbook campaignResultsWB = xlApp.Workbooks.Open(contactpath, ReadOnly: false);
                    Excel.Worksheet campaignResultsWS = campaignResultsWB.Worksheets[1];

                    Excel.Range last = campaignResultsWS.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int maxRow = last.Row;
                    Console.WriteLine("Last row used in the email campaign sheet is: " + maxRow);

                    setLoadingLabel("Identifying Phising Platform");
                    int totalEmails = maxRow - 1;
                    int failedEmails = 0;
                    int openedEmails = 0;

                    if ("Email".Equals(Convert.ToString(campaignResultsWS.Range["A1"].Value2)) & "Clicked at".Equals(Convert.ToString(campaignResultsWS.Range["B1"].Value2)) & "Data entered at".Equals(Convert.ToString(campaignResultsWS.Range["C1"].Value2))
                        & "Attachment opened at".Equals(Convert.ToString(campaignResultsWS.Range["D1"].Value2)) & "Macro enabled at".Equals(Convert.ToString(campaignResultsWS.Range["E1"].Value2)) & "Opened at".Equals(Convert.ToString(campaignResultsWS.Range["F1"].Value2))
                        & "Delivered at".Equals(Convert.ToString(campaignResultsWS.Range["G1"].Value2)) & "Bounced at".Equals(Convert.ToString(campaignResultsWS.Range["H1"].Value2)) & "First Name".Equals(Convert.ToString(campaignResultsWS.Range["I1"].Value2))
                        & "Last Name".Equals(Convert.ToString(campaignResultsWS.Range["J1"].Value2)) & "Job Title".Equals(Convert.ToString(campaignResultsWS.Range["K1"].Value2)) & "Group".Equals(Convert.ToString(campaignResultsWS.Range["L1"].Value2))
                        & "Manager Name".Equals(Convert.ToString(campaignResultsWS.Range["M1"].Value2)) & "Manager Email".Equals(Convert.ToString(campaignResultsWS.Range["N1"].Value2)) & "Location".Equals(Convert.ToString(campaignResultsWS.Range["O1"].Value2))
                        & "Division".Equals(Convert.ToString(campaignResultsWS.Range["P1"].Value2)) & "Employee number".Equals(Convert.ToString(campaignResultsWS.Range["Q1"].Value2)) & "IP Address".Equals(Convert.ToString(campaignResultsWS.Range["R1"].Value2))
                        & "IP Location".Equals(Convert.ToString(campaignResultsWS.Range["S1"].Value2)) & "Browser".Equals(Convert.ToString(campaignResultsWS.Range["T1"].Value2)) & "Browser Version".Equals(Convert.ToString(campaignResultsWS.Range["U1"].Value2))
                        & "Operating System".Equals(Convert.ToString(campaignResultsWS.Range["V1"].Value2)) & "Email Template".Equals(Convert.ToString(campaignResultsWS.Range["W1"].Value2)))
                    {
                        //delete columns: C - E; G - H (D & E); K - V
                        Excel.Range range = campaignResultsWS.Range["C1", "E1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["D1", "E1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["F1", "Q1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["A1", "F1"];
                        range.EntireColumn.AutoFit();

                        for (int i = 2; i <= maxRow; i++)
                        {
                            if (Convert.ToString(campaignResultsWS.Range["B" + i].Value2) != null)
                            {
                                failedEmails++;
                            }
                            if (Convert.ToString(campaignResultsWS.Range["C" + i].Value2) != null)
                            {
                                openedEmails++;
                            }
                        }
                        Console.WriteLine("Phishing Results");
                        Console.WriteLine("total emails: " + totalEmails);
                        Console.WriteLine("failed count: " + failedEmails);
                        Console.WriteLine("opened count: " + openedEmails);

                        //copy columns A1 - F[maxRow] and paste into the report doc 
                    }
                    else if ("First Name".Equals(Convert.ToString(campaignResultsWS.Range["A1"].Value2)) & "Last Name".Equals(Convert.ToString(campaignResultsWS.Range["B1"].Value2)) & "Email Address".Equals(Convert.ToString(campaignResultsWS.Range["C1"].Value2))
                      & "Group".Equals(Convert.ToString(campaignResultsWS.Range["D1"].Value2)) & "Viewed Images / Opened Email".Equals(Convert.ToString(campaignResultsWS.Range["E1"].Value2)) & "Passed".Equals(Convert.ToString(campaignResultsWS.Range["F1"].Value2))
                      & "Failed".Equals(Convert.ToString(campaignResultsWS.Range["G1"].Value2)) & "Failed Date".Equals(Convert.ToString(campaignResultsWS.Range["H1"].Value2)) & "Campaign".Equals(Convert.ToString(campaignResultsWS.Range["I1"].Value2))
                      & "campaign type".Equals(Convert.ToString(campaignResultsWS.Range["J1"].Value2)) & "Payload".Equals(Convert.ToString(campaignResultsWS.Range["K1"].Value2)) & "Payload Type".Equals(Convert.ToString(campaignResultsWS.Range["L1"].Value2))
                      & "Group(s)".Equals(Convert.ToString(campaignResultsWS.Range["M1"].Value2)) & "Clicked Link".Equals(Convert.ToString(campaignResultsWS.Range["N1"].Value2)))
                    {
                        //delete columns: D, F, H - J, L - N
                        Excel.Range range = campaignResultsWS.Range["D1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["E1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["F1", "H1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["G1", "I1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["A1", "F1"];
                        range.EntireColumn.AutoFit();

                        for (int i = 2; i <= maxRow; i++)
                        {
                            if (campaignResultsWS.Range["E" + i].Value.Equals("Yes"))
                            {
                                failedEmails++;
                            }
                            if (campaignResultsWS.Range["D" + i].Value.Equals("Yes"))
                            {
                                openedEmails++;
                            }
                        }
                        Console.WriteLine("total emails: " + totalEmails);
                        Console.WriteLine("failed count: " + failedEmails);
                        Console.WriteLine("opened count: " + openedEmails);
                    } else
                    {
                        MessageBox.Show("Please use an UNEDITED phishing campaign file (.csv) that was downloaded from Insight or KnowBe4.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        showReportTab();
                        //exit method to return to the Main Class
                        return;
                    }

                    //-------------------------------------------------------- Specific to Phishing Campaigns ------------------------------------------------------------
                    setLoadingLabel("Starting Word");
                    Word.Application wordApp = new Word.Application();
                    //wordApp.Visible = true;
                    Word.Document reportDoc = wordApp.Documents.Open("C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Report Template - Phishing.docx", ReadOnly: false);

                    setLoadingLabel("Updating Content Control Fields");
                    reportDoc.ContentControls[1].Range.Text = txtClient.Text.ToString(); //Client's Name
                    reportDoc.ContentControls[12].Range.Text = txtPOC.Text; //Contact's Name
                    reportDoc.ContentControls[4].Range.Text = (totalEmails).ToString(); //Total Emails
                    reportDoc.ContentControls[5].Range.Text = (totalEmails - failedEmails).ToString(); //Passed Emails
                    reportDoc.ContentControls[6].Range.Text = failedEmails.ToString(); //Failed Emails
                    reportDoc.ContentControls[8].Range.Text = openedEmails.ToString(); //Opened Emails
                    reportDoc.ContentControls[10].Range.Text = dateTimePicker1.Value.ToShortDateString(); //Phishing Testing Email Start Date
                    if (failedEmails > 0)
                    {
                        reportDoc.ContentControls[7].DropdownListEntries[2].Select(); //an unsuccesful
                    }
                    else
                    {
                        reportDoc.ContentControls[7].DropdownListEntries[1].Select(); //a successful
                    }

                    /*
                    int jk = 1;
                    foreach (Word.Paragraph prg in reportDoc.Paragraphs)
                    {
                        Word.Style style = prg.get_Style() as Word.Style;
                        string styleName = style.NameLocal;
                        string text = prg.Range.Text;
                        Console.WriteLine("Prg " + jk + " Style name: " + styleName);
                        Console.WriteLine("Prg " + jk + " Text: " + text);
                        jk++;
                    }
                    Console.WriteLine("------------------------------------------------------------------------------------------------------");
                    */

                    setLoadingLabel("Updating Vishing Charts Data");
                    Word.Chart vishingChart = reportDoc.Shapes[3].Chart;
                    Excel.Workbook vishingChartWB = vishingChart.ChartData.Workbook;
                    Excel.Worksheet vishingChartWS = vishingChartWB.Worksheets[1];
                    vishingChartWS.Range["B2"].Value = (totalEmails - openedEmails); //Not Opened Emails
                    vishingChartWS.Range["B3"].Value = openedEmails; //Opened Emails

                    vishingChart = reportDoc.Shapes[4].Chart;
                    vishingChartWB = vishingChart.ChartData.Workbook;
                    vishingChartWS = vishingChartWB.Worksheets[1];
                    vishingChartWS.Range["B2"].Value = (totalEmails - failedEmails); //Passed Emails
                    vishingChartWS.Range["B3"].Value = failedEmails; //Failed Emails


                    setLoadingLabel("Pasting Phishing Email Engagement Table");
                    campaignResultsWS.Range["A1", "F" + maxRow].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range["A1", "F" + maxRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range["A1", "F" + maxRow].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range["A1", "F" + maxRow].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range["A1", "F" + maxRow].Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range["A1", "F" + maxRow].Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range["A1", "F" + maxRow].Copy();
                    reportDoc.Paragraphs[42].Range.Paste();
                    reportDoc.Tables[1].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;

                    setLoadingLabel("Save the Phishing Report");
                    wordApp.Visible = true;
                    int currentYear = DateTime.Now.Year;
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

                    setLoadingLabel("Exiting Word and Excel");
                    xlApp.DisplayAlerts = false;
                    campaignResultsWB.Close();
                    xlApp.Quit();
                    reportDoc.Close();
                    wordApp.Quit();
                    setLoadingLabel("Success!");
                    showReportTab();
                }
            }
            else if (radBoth.Checked) //--------------------------------------------------------------------------------------------------------------------------------------------
            {
                hideReportTab();
                setLoadingLabel("Open the Call List and Phishing Campaign Results");
                //open the file that contains the email campaign results 
                OpenFileDialog openCampaignResults = new OpenFileDialog() { Filter = "Comma Seperated Values|*.csv", ValidateNames = true, Title = "Pick the Email Campaign Results .csv file for " + txtClient.Text + "'s RSE." };
                DialogResult phishingResult = openCampaignResults.ShowDialog();
                //open the file that contains the phone call results 
                OpenFileDialog openCallList = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true, Title = "Pick the Call List .xlsx file that was made by the RSE Tool." };
                DialogResult vishingResult = openCallList.ShowDialog();

                if (phishingResult == DialogResult.OK & vishingResult == DialogResult.OK)
                {
                    setLoadingLabel("Starting Excel");
                    Excel.Application xlApp = new Excel.Application();
                    //xlApp.Visible = true;
                    contactpath = openCampaignResults.FileName;

                    //---------------------------------------------------------------- Email Calculations -------------------------------------------------------------------------------------

                    Excel.Workbook campaignResultsWB = xlApp.Workbooks.Open(contactpath, ReadOnly: false);
                    Excel.Worksheet campaignResultsWS = campaignResultsWB.Worksheets[1];

                    Excel.Range last = campaignResultsWS.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int maxRow = last.Row;

                    setLoadingLabel("Identifying the Phishing Platform");
                    int totalEmails = maxRow - 1;
                    int failedEmails = 0;
                    int openedEmails = 0;

                    if ("Email".Equals(Convert.ToString(campaignResultsWS.Range["A1"].Value2)) & "Clicked at".Equals(Convert.ToString(campaignResultsWS.Range["B1"].Value2)) & "Data entered at".Equals(Convert.ToString(campaignResultsWS.Range["C1"].Value2))
                        & "Attachment opened at".Equals(Convert.ToString(campaignResultsWS.Range["D1"].Value2)) & "Macro enabled at".Equals(Convert.ToString(campaignResultsWS.Range["E1"].Value2)) & "Opened at".Equals(Convert.ToString(campaignResultsWS.Range["F1"].Value2))
                        & "Delivered at".Equals(Convert.ToString(campaignResultsWS.Range["G1"].Value2)) & "Bounced at".Equals(Convert.ToString(campaignResultsWS.Range["H1"].Value2)) & "First Name".Equals(Convert.ToString(campaignResultsWS.Range["I1"].Value2))
                        & "Last Name".Equals(Convert.ToString(campaignResultsWS.Range["J1"].Value2)) & "Job Title".Equals(Convert.ToString(campaignResultsWS.Range["K1"].Value2)) & "Group".Equals(Convert.ToString(campaignResultsWS.Range["L1"].Value2))
                        & "Manager Name".Equals(Convert.ToString(campaignResultsWS.Range["M1"].Value2)) & "Manager Email".Equals(Convert.ToString(campaignResultsWS.Range["N1"].Value2)) & "Location".Equals(Convert.ToString(campaignResultsWS.Range["O1"].Value2))
                        & "Division".Equals(Convert.ToString(campaignResultsWS.Range["P1"].Value2)) & "Employee number".Equals(Convert.ToString(campaignResultsWS.Range["Q1"].Value2)) & "IP Address".Equals(Convert.ToString(campaignResultsWS.Range["R1"].Value2))
                        & "IP Location".Equals(Convert.ToString(campaignResultsWS.Range["S1"].Value2)) & "Browser".Equals(Convert.ToString(campaignResultsWS.Range["T1"].Value2)) & "Browser Version".Equals(Convert.ToString(campaignResultsWS.Range["U1"].Value2))
                        & "Operating System".Equals(Convert.ToString(campaignResultsWS.Range["V1"].Value2)) & "Email Template".Equals(Convert.ToString(campaignResultsWS.Range["W1"].Value2)))
                    {
                        //delete columns: C - E; G - H (D & E); K - V
                        Excel.Range range = campaignResultsWS.Range["C1", "E1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["D1", "E1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["F1", "Q1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["A1", "F1"];
                        range.EntireColumn.AutoFit();

                        for (int i = 2; i <= maxRow; i++)
                        {
                            if (Convert.ToString(campaignResultsWS.Range["B" + i].Value2) != null)
                            {
                                failedEmails++;
                            }
                            if (Convert.ToString(campaignResultsWS.Range["C" + i].Value2) != null)
                            {
                                openedEmails++;
                            }
                        }
                        Console.WriteLine("total emails: " + totalEmails);
                        Console.WriteLine("failed count: " + failedEmails);
                        Console.WriteLine("opened count: " + openedEmails);

                        //copy columns A1 - F[maxRow] and paste into the report doc 
                    }
               else if ("First Name".Equals(Convert.ToString(campaignResultsWS.Range["A1"].Value2)) & "Last Name".Equals(Convert.ToString(campaignResultsWS.Range["B1"].Value2)) & "Email Address".Equals(Convert.ToString(campaignResultsWS.Range["C1"].Value2))
                            & "Group".Equals(Convert.ToString(campaignResultsWS.Range["D1"].Value2)) & "Viewed Images / Opened Email".Equals(Convert.ToString(campaignResultsWS.Range["E1"].Value2)) & "Passed".Equals(Convert.ToString(campaignResultsWS.Range["F1"].Value2))
                            & "Failed".Equals(Convert.ToString(campaignResultsWS.Range["G1"].Value2)) & "Failed Date".Equals(Convert.ToString(campaignResultsWS.Range["H1"].Value2)) & "Campaign".Equals(Convert.ToString(campaignResultsWS.Range["I1"].Value2))
                            & "campaign type".Equals(Convert.ToString(campaignResultsWS.Range["J1"].Value2)) & "Payload".Equals(Convert.ToString(campaignResultsWS.Range["K1"].Value2)) & "Payload Type".Equals(Convert.ToString(campaignResultsWS.Range["L1"].Value2))
                            & "Group(s)".Equals(Convert.ToString(campaignResultsWS.Range["M1"].Value2)) & "Clicked Link".Equals(Convert.ToString(campaignResultsWS.Range["N1"].Value2)))
                    {
                        //delete columns: D, F, H - J, L - N
                        Excel.Range range = campaignResultsWS.Range["D1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["E1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["F1", "H1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["G1", "I1"];
                        range.EntireColumn.Delete(Missing.Value);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                        range = campaignResultsWS.Range["A1", "F1"];
                        range.EntireColumn.AutoFit();

                        for (int i = 2; i <= maxRow; i++)
                        {
                            if (campaignResultsWS.Range["E" + i].Value.Equals("Yes"))
                            {
                                failedEmails++;
                            }
                            if (campaignResultsWS.Range["D" + i].Value.Equals("Yes"))
                            {
                                openedEmails++;
                            }
                        }
                        Console.WriteLine("Phishing Results");
                        Console.WriteLine("total emails: " + totalEmails);
                        Console.WriteLine("failed count: " + failedEmails);
                        Console.WriteLine("opened count: " + openedEmails);
                    } else
                    {
                        MessageBox.Show("Please use an UNEDITED phishing campaign file (.csv) that was downloaded from Insight or KnowBe4.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        showReportTab();
                        //exit method to return to the Main Class
                        return;
                    }
                    string emailResultRange = "A1:F" + maxRow.ToString();
                    //--------------------------------------------------------------------- Phone Call Calculations -------------------------------------------------------------------------------------------------
                    setLoadingLabel("Verifying Vishing Call List File");
                    contactpath = openCallList.FileName;
                    dataTable = excelSheetToDataTable(contactpath, false);

                    //check that the Call List was made by the RSE Tool so that there are no errors later 
                    if (!dataTable.Rows[0][0].Equals("Calling As") & !dataTable.Rows[1][0].Equals("Phone # Displayed") & !dataTable.Rows[2][0].Equals("Name Drop")
                        & !dataTable.Rows[3][0].Equals("Engagements Needed") & !dataTable.Rows[4][0].Equals("Engagements per Day") & !dataTable.Rows[5][0].Equals("Current Engagements")
                        & !dataTable.Rows[6][0].Equals("Business Hours"))
                    {
                        MessageBox.Show("Please use a Call List that was created by the RSE Tool.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        showReportTab();
                        //exit method to return to the Main Class
                        return;
                    }

                    dataTable.Rows[0].Delete();
                    dataTable.Rows[1].Delete();
                    dataTable.Rows[2].Delete();
                    dataTable.Rows[3].Delete();
                    dataTable.Rows[4].Delete();
                    dataTable.Rows[5].Delete();
                    dataTable.Rows[6].Delete();
                    dataTable.Rows[7].Delete();
                    dataTable.AcceptChanges();

                    string path = "C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Vishing Notes Template.xlsx";
                    Excel.Workbook wb = xlApp.Workbooks.Open(path, ReadOnly: false);
                    Excel.Worksheet ws = wb.Worksheets[1];
                    Excel.Worksheet tempWS = wb.Worksheets[2];

                    int dtMaxRow = dataTable.Rows.Count;
                    int dtResultCol = 0;
                    Boolean hasExtension = false;
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

                    setLoadingLabel("Creating Vishing Notes File");
                    int j = 2;
                    maxRow = 0;
                    string tempDate = null;
                    List<string> dates = new List<string>();
                    string tempDescrip = null;
                    int voicemailCount = 0;
                    int passedCount = 0;
                    int failedCount = 0;
                    for (int i = 1; i < dtMaxRow; i++) //i = current DataTable row
                    {
                        tempDate = null;
                        dates.Clear();
                        tempDescrip = "temp";

                        if (!dataTable.Rows[i][dtResultCol].Equals(DBNull.Value))
                        {
                            if (dataTable.Rows[i][dtResultCol].Equals("PASSED"))
                            {
                                passedCount++;
                            }
                            if (dataTable.Rows[i][dtResultCol].Equals("FAILED"))
                            {
                                failedCount++;
                            }

                            ws.Range["A" + j.ToString()].Value = dataTable.Rows[i][dtResultCol]; //Final Result
                            ws.Range["B" + j.ToString()].Value = dataTable.Rows[i][0]; //Name
                            ws.Range["C" + j.ToString()].Value = dataTable.Rows[i][1]; //Phone 
                            if (hasExtension == true)
                            {
                                ws.Range["D" + j.ToString()].Value = dataTable.Rows[i][dtExtensionCol]; //Extension
                            }

                            for (int k = dataTable.Columns.Count - 1; k > dtResultCol; k--) //k = current DataTable Column to the right of Result Column
                            {
                                if (!dataTable.Rows[i][k].Equals(DBNull.Value))
                                {
                                    if (dataTable.Rows[i][k].Equals("Voicemail"))
                                    {
                                        voicemailCount++;
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

                            ws.Range["A" + (j + 1).ToString()].Value = "Dates:  " + String.Join(", ", dates); //Dates:
                            ws.Range["A" + (j + 2).ToString()].Value = "Description: " + tempDescrip; //Description
                            tempWS.Range["A1"].Value = "Description: " + tempDescrip; //Temp Description
                            double rowHeight = tempWS.Range["A1"].RowHeight;
                            ws.Range["A" + (j + 2).ToString()].RowHeight = rowHeight; //Description

                            maxRow = j + 2;
                            j = j + 4;
                        }
                    }

                    string vishingNotesRange = "A1:D" + maxRow.ToString();
                    Console.WriteLine("Vishing Results");
                    Console.WriteLine("Unanswered: " + voicemailCount);
                    Console.WriteLine("Compromised: " + failedCount);
                    Console.WriteLine("Uncompromised: " + passedCount);


                    //excelApp.ScreenUpdating = false
                    setLoadingLabel("Save the Vishing Notes File");
                    int currentYear = DateTime.Now.Year;
                    SaveFileDialog vishingNotesFileStream = new SaveFileDialog();
                    vishingNotesFileStream.Title = "Vishing Notes/Phone Engagment Detail Tabke File Save as";
                    vishingNotesFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Vishing Notes.xlsx";
                    vishingNotesFileStream.DefaultExt = ".xlsx";
                    vishingNotesFileStream.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                    DialogResult vishingNotesResult = vishingNotesFileStream.ShowDialog();
                    if (vishingNotesResult == DialogResult.OK)
                    {
                        string fileName = vishingNotesFileStream.FileName;
                        wb.SaveAs(fileName);
                    }


                    //------------------------------------------------------------------ Specific to Phishing and Vishing  -----------------------------------------------------------------------
                    setLoadingLabel("Starting Word");
                    Word.Application wordApp = new Word.Application();
                    //wordApp.Visible = true;
                    Word.Document reportDoc = wordApp.Documents.Open("C:\\Users\\bmartin\\Documents\\Tools\\Repos\\Trace-RSE-Tool-master\\csv test-6.28.18\\reports\\RSE Report Template - Phishing and Vishing.docx", ReadOnly: false);

                    setLoadingLabel("Updating the Content Control fields");
                    reportDoc.ContentControls[1].Range.Text = txtClient.Text.ToString(); //Client's Name
                    reportDoc.ContentControls[4].Range.Text = txtPOC.Text; //Contact's Name
                    reportDoc.ContentControls[6].Range.Text = (passedCount + failedCount + voicemailCount).ToString(); //Total Calls
                    reportDoc.ContentControls[7].Range.Text = passedCount.ToString(); //Uncompromised
                    reportDoc.ContentControls[8].Range.Text = failedCount.ToString(); //Compromised
                    reportDoc.ContentControls[9].Range.Text = voicemailCount.ToString(); //Unanswered
                    if (failedCount > 0) //Choose an item for Phone Calls
                    {
                        reportDoc.ContentControls[10].DropdownListEntries[2].Select(); //an unsuccesful
                    }
                    else
                    {
                        reportDoc.ContentControls[10].DropdownListEntries[1].Select(); //a successful
                    }
                    reportDoc.ContentControls[12].Range.Text = (totalEmails).ToString(); //Total Emails
                    reportDoc.ContentControls[13].Range.Text = (totalEmails - failedEmails).ToString(); //Passed Emails
                    reportDoc.ContentControls[14].Range.Text = failedEmails.ToString(); //Failed Emails
                    reportDoc.ContentControls[16].Range.Text = openedEmails.ToString(); //Opened Emails
                    reportDoc.ContentControls[18].Range.Text = dateTimePicker1.Value.ToShortDateString(); //Phishing Test Start Date
                    if (failedEmails > 0)
                    {
                        reportDoc.ContentControls[15].DropdownListEntries[2].Select(); //an unsuccesful
                    }
                    else
                    {
                        reportDoc.ContentControls[15].DropdownListEntries[1].Select(); //a successful
                    }


                    setLoadingLabel("Updating Vishing and Phishing Charts' data");
                    Word.Chart vishingChart = reportDoc.Shapes[5].Chart;
                    Excel.Workbook vishingChartWB = vishingChart.ChartData.Workbook;
                    Excel.Worksheet vishingChartWS = vishingChartWB.Worksheets[1];
                    vishingChartWS.Range["B2"].Value = passedCount; //passed calls
                    vishingChartWS.Range["B3"].Value = failedCount; //failed calls 
                    vishingChartWS.Range["B4"].Value = voicemailCount; //did not answer
                    System.Threading.Thread.Sleep(2000);
                    vishingChartWB.Close();

                    Word.Chart phishingOpenChart = reportDoc.Shapes[6].Chart;
                    Excel.Workbook phishingOpenChartWB = phishingOpenChart.ChartData.Workbook;
                    Excel.Worksheet phishingOpenChartWS = phishingOpenChartWB.Worksheets[1];
                    phishingOpenChartWS.Range["B2"].Value = (totalEmails - openedEmails); //Not Opened Emails
                    phishingOpenChartWS.Range["B3"].Value = openedEmails; //Opened Emails
                    System.Threading.Thread.Sleep(2000);
                    phishingOpenChartWB.Close();

                    Word.Chart phishingResultChart = reportDoc.Shapes[7].Chart;
                    Excel.Workbook phishingResultChartWB = phishingResultChart.ChartData.Workbook;
                    Excel.Worksheet phishingResultChartWS = phishingResultChartWB.Worksheets[1];
                    phishingResultChartWS.Range["B2"].Value = (totalEmails - failedEmails); //Passed Emails
                    phishingResultChartWS.Range["B3"].Value = failedEmails; //Failed Emails
                    System.Threading.Thread.Sleep(2000);
                    phishingResultChartWB.Close();


                    setLoadingLabel("Pasting Vishing Notes Summary into Report");
                    //paste Vishing Call Notes into Paragraph 72 //
                    ws.Range[vishingNotesRange].Copy();
                    System.Threading.Thread.Sleep(2000);
                    reportDoc.Paragraphs[72].Range.Paste();
                    xlApp.DisplayAlerts = false;
                    wb.Close();

                    setLoadingLabel("Pasting Email Engagement Table into Report");
                    //paste Email Results into Paragraph 70 //
                    campaignResultsWS.Range[emailResultRange].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range[emailResultRange].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range[emailResultRange].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range[emailResultRange].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range[emailResultRange].Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range[emailResultRange].Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                    campaignResultsWS.Range[emailResultRange].Copy();
                    System.Threading.Thread.Sleep(2000);
                    reportDoc.Paragraphs[70].Range.Paste();
                    campaignResultsWB.Close();
                    xlApp.Quit();

                    setLoadingLabel("Formatting Report for Pretty Printing");
                    for (int i = 1; i <= reportDoc.Tables.Count; i++)
                    {
                        //System.Threading.Thread.Sleep(1000);
                        reportDoc.Tables[i].Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    }


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
                    int emailTableLastRow = reportDoc.Tables[1].Rows.Count;
                    //if "Phone Engagement Details" HEADER is on the same page as the last row of the Email Engagement Detail TABLE then insert a page break
                    if (reportDoc.Paragraphs[phoneEngagementDetailParagraph].Range.Information[Word.WdInformation.wdActiveEndPageNumber] == reportDoc.Tables[1].Rows[emailTableLastRow].Range.Information[Word.WdInformation.wdActiveEndPageNumber])
                    {
                        reportDoc.Paragraphs.Add(reportDoc.Paragraphs[phoneEngagementDetailParagraph].Range);
                        reportDoc.Paragraphs[phoneEngagementDetailParagraph].Range.InsertBreak(Word.WdBreakType.wdPageBreak);
                        reportDoc.Paragraphs[phoneEngagementDetailParagraph].set_Style(reportDoc.Styles["Normal"]);
                    }

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

                    for (int i = 1; i <= reportDoc.Tables.Count; i++)
                    {
                        reportDoc.Tables[i].Rows[1].Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }

                    setLoadingLabel("Save the Phising and Vishing Report");
                    wordApp.Visible = true;
                    SaveFileDialog phishingAndVishingReportFileStream = new SaveFileDialog();
                    phishingAndVishingReportFileStream.FileName = txtClient.Text.ToString().Trim() + " RSE " + currentYear + " Phishing and Vishing Report.docx";
                    phishingAndVishingReportFileStream.DefaultExt = ".docx";
                    phishingAndVishingReportFileStream.Filter = "Word Document File (.docx)|*.docx";
                    DialogResult phishingAndVishingReportResult = phishingAndVishingReportFileStream.ShowDialog();
                    if (phishingAndVishingReportResult == DialogResult.OK)
                    {
                        string fileName = phishingAndVishingReportFileStream.FileName;
                        reportDoc.SaveAs(fileName);
                    }

                    setLoadingLabel("Exiting Word and Excel");
                    reportDoc.Close();
                    wordApp.Quit();
                    setLoadingLabel("Success!");
                    showReportTab();
                }
            }
        }

        private void hideReportTab()
        {
            label2.Visible = false;
            label5.Visible = false;
            label7.Location = new Point(133, 73);
            label6.Visible = false;
            txtClient.Visible = false;
            txtPOC.Visible = false;
            dateTimePicker1.Visible = false;
            radEmail.Visible = false;
            radPhone.Visible = false;
            radBoth.Visible = false;
            btnReportShell.Visible = false;
            //pictureBox1.Visible = true;

            //backgroundWorker1.RunWorkerAsync();
        }

        private void showReportTab()
        {
            txtClient.Clear();
            txtPOC.Clear();
            dateTimePicker1.Value = new DateTime(2019,1,1);
            label2.Visible = true;
            label5.Visible = true;
            label7.Text = "Point of Contact*";
            label7.Location = new Point(4, 73);
            //label7.Visible = true;
            label6.Visible = true;
            txtClient.Visible = true; 
            txtPOC.Visible = true;
            dateTimePicker1.Visible = true;
            radEmail.Visible = true;
            radPhone.Visible = true;
            radBoth.Visible = true;
            btnReportShell.Visible = true;
            pictureBox1.Visible = false;
        }

        private void setLoadingLabel(string text)
        {
            label7.Text = text;
            label7.Location = new Point((397 - label7.Width) / 2, 73);
        }
    }
}

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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;

namespace csv_test_6._28._18
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

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
                    contactpath = openFile.FileName;
                    fileType = "Word";
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
                    fileType = "Excel";
                    contactpath = openFile.FileName;                    
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

        private void btnPreview_Click(object sender, EventArgs e)
        {
            ToolTip toolPreview = new ToolTip();
            toolPreview.ShowAlways = false;
            toolPreview.SetToolTip(btnPreview, "Preview Extracted Data");
            Preview newpreview = new Preview();
            newpreview.ContactPath = contactpath;
            newpreview.FileType = fileType;
            newpreview.ShowDialog();
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            ToolTip toolPreview = new ToolTip();
            toolPreview.ShowAlways = false;
            toolPreview.SetToolTip(btnPreview, "Preview Extracted Data");
            Preview newpreview = new Preview();
            newpreview.ContactPath = contactpath;
            newpreview.FileType = fileType;
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

        private void btnReportShell_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    Read reading = new Read();
            //    reading.NameFile = contactpath;
            //    // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
            //    string[] copyHeader = reading.WordTableHeader();
            //    // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
            //    string[,] copyData = reading.WordDoc();
            //    // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
            //    int rowcount = copyData.GetUpperBound(0) + 1;
            //    int colcount = copyData.GetUpperBound(1) + 1;
            //    // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
            //    string[,] filterData = new string[rowcount, 5];
            //    string[,] reorderedData = new string[rowcount, 4];
            //    for (int i = 0; i < rowcount; i++)
            //    {
            //        for (int j = 0; j < colcount; j++)
            //        {
            //            // SET FIRST NAME VALUE
            //            if (copyHeader[j].Contains("first"))
            //            {
            //                filterData[i, 0] = copyData[i, j];
            //            }
            //            // SET LAST NAME VALUE
            //            else if (copyHeader[j].Contains("last"))
            //            {
            //                filterData[i, 1] = copyData[i, j];
            //            }
            //            // SET PHONE NUMBER VALUE
            //            else if (copyHeader[j].Contains("phone"))
            //            {
            //                filterData[i, 2] = copyData[i, j];
            //            }
            //            // SET EXTENSION VALUE
            //            else if (copyHeader[j].Contains("ext"))
            //            {
            //                filterData[i, 3] = copyData[i, j];
            //            }
            //            // SET EMAIL ADDRESS VALUE
            //            else if (copyHeader[j].Contains("email"))
            //            {
            //                filterData[i, 4] = copyData[i, j];
            //            }
            //        }
            //    }
            //    for (int a = 0; a < filterData.GetUpperBound(0) + 1; a++)
            //    {
            //        reorderedData[a, 0] = filterData[a, 0] + " " + filterData[a, 1];
            //        for (int b = 1; b <= 4; b++)
            //        {
            //            reorderedData[a, b] = filterData[a, b];
            //        }
            //    }
            //    string saveFile = string.Empty;
            //    string tempFile = @"reports\RSE Vishing Report.docx";
            //    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
            //    SaveFileDialog fileStream = new SaveFileDialog();
            //    fileStream.FileName = "RSE Vishing Report.docx";
            //    fileStream.DefaultExt = ".docx";
            //    fileStream.Filter = "Word Document File (*.docx)|*.docx";
            //    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
            //    DialogResult result = fileStream.ShowDialog();
            //    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
            //    if (result == DialogResult.OK)
            //    {
            //        saveFile = fileStream.FileName;
            //    }
            //    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
            //    Create makeFile = new Create();
            //    makeFile.ReportFile = tempFile;
            //    makeFile.PerfectArray = reorderedData;
            //    makeFile.SaveFile = saveFile;
            //    makeFile.MakeReport();
            //}

            //catch
            //{
            //    MessageBox.Show("Something went wrong");
            //}
        }

        //method that will create a new Excel Sheet that will be used when making calls to clients 
        private void btnCreateCallList_Click(object sender, EventArgs e)
        {

        }
    }
}

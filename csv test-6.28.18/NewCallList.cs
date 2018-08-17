using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelDataReader;

namespace csv_test_6._28._18
{
    public partial class NewCallList : Form
    {

        public DataTable dataTable { get; set; }
        public List<String> headers = new List<String>();
        public string fileName;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlWorkSheet;

        private int hasName;
        private int hasPhone;
        private int hasExtension;
        private int hasEmail;
        private int hasTitle;
        private int hasLocation;
        private int hasDepartment;

        private string companyName = string.Empty;
        private string nameDrop = string.Empty;
        private string engagementsNeeded = string.Empty;
        private string engagementsPerDay = string.Empty;
        private string businessHours = string.Empty;
        private string timeZone = string.Empty;
        private string callingAs = string.Empty;
        private string phoneNumberDisplayed = string.Empty;

        public NewCallList()
        {
            InitializeComponent();
            saveDataLabel.Visible = false;
            hasName = -1;
            hasPhone = -1;
            hasExtension = -1;
            hasEmail = -1;
            hasTitle = -1;
            hasLocation = -1;
            hasDepartment = -1;
        }

        private void NewCallList_Load(object sender, EventArgs e)
        {
            foreach (DataColumn column in dataTable.Columns)
            {
                headers.Add(column.ColumnName);
            }

            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i].ToLower().Contains("name"))
                {
                    hasName = i;
                }
                if (headers[i].ToLower().Contains("phone"))
                {
                    hasPhone = i;
                }
                if (headers[i].ToLower().Contains("extension") | headers[i].ToLower().Contains("ext"))
                {
                    hasExtension = i;
                }
                if (headers[i].ToLower().Contains("email") | headers[i].ToLower().Contains("e-mail"))
                {
                    hasEmail = i;
                }
                if (headers[i].ToLower().Contains("title"))
                {
                    hasTitle = i;
                }
                if (headers[i].ToLower().Contains("location") | headers[i].ToLower().Contains("branch") | headers[i].ToLower().Contains("office"))
                {
                    hasLocation = i;
                }
                if (headers[i].ToLower().Contains("department") | headers[i].ToLower().Contains("dept") | headers[i].ToLower().Contains("division"))
                {
                    hasDepartment = i;
                }
            }


            if (hasName == -1)
            {
                CkBxName.Enabled = false;
            }
            if (hasPhone == -1)
            {
                ckBxPhone.Enabled = false;
            }
            if (hasExtension == -1)
            {
                CkBxExtension.Enabled = false;
            }
            if (hasEmail == -1)
            {
                CkBxEmail.Enabled = false;
            }
            if (hasTitle == -1)
            {
                CkBxTitle.Enabled = false;
            }
            if (hasLocation == -1)
            {
                CkBxLocation.Enabled = false;
            }
            if (hasDepartment == -1)
            {
                CkBxDepartment.Enabled = false;
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtBxCompanyName.Text))
            {
                companyName = txtBxCompanyName.Text;
            }
            else
            {
                saveDataLabel.Visible = false;
                MessageBox.Show("Please Enter Company Name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtBxCompanyName.Clear();
                txtBxCompanyName.Focus();
                return;
            }
            int currentYear = DateTime.Now.Year;
            string saveDataText = "Saving Data to \"" + companyName.Trim() + " RSE " + currentYear + " Call List.xlsx\"";
            saveDataLabel.Text = saveDataText;
            saveDataLabel.Visible = true;
            nameDrop = txtBxNameDrop.Text;
            if (!string.IsNullOrWhiteSpace(txtBxEngagementsNeeded.Text))
            {
                engagementsNeeded = txtBxEngagementsNeeded.Text;
            }
            else
            {
                saveDataLabel.Visible = false;
                MessageBox.Show("Please Enter the Number of Engagements Needed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtBxEngagementsNeeded.Clear();
                txtBxEngagementsNeeded.Focus();
                return;
            }
            engagementsPerDay = txtBxEngagementsPerDay.Text;
            businessHours = txtBxHours.Text;
            timeZone = txtBxTimeZone.Text;
            if (!string.IsNullOrWhiteSpace(txtBxCallingAs.Text))
            {
                callingAs = txtBxCallingAs.Text;
            }
            else
            {
                saveDataLabel.Visible = false;
                MessageBox.Show("Please Enter the Company you are Calling As.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtBxCallingAs.Clear();
                txtBxCallingAs.Focus();
                return;
            }
            if (!string.IsNullOrWhiteSpace(txtBxNumberDisplayed.Text))
            {
                phoneNumberDisplayed = txtBxNumberDisplayed.Text;
            }
            else
            {
                saveDataLabel.Visible = false;
                MessageBox.Show("Please Enter the Phone Number you want Displayed when Making Calls.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtBxCallingAs.Clear();
                txtBxCallingAs.Focus();
                return;
            }
            saveDataLabel.Text = saveDataText + ".";


            //create new headers
            List<String[]> newHeaders = new List<String[]>();
            String[] temp;
            if (CkBxName.Checked)
            {
                temp = new String[] { "Name", headers[hasName] };
                newHeaders.Add(temp);
            }
            if (ckBxPhone.Checked)
            {
                temp = new String[] { "Phone", headers[hasPhone] };
                newHeaders.Add(temp);
            }
            if (CkBxExtension.Checked)
            {
                temp = new String[] { "Extension", headers[hasExtension] };
                newHeaders.Add(temp);
            }
            if (CkBxEmail.Checked)
            {
                temp = new String[] { "Email", headers[hasEmail] };
                newHeaders.Add(temp);
            }
            if (CkBxTitle.Checked)
            {
                temp = new String[] { "Title", headers[hasTitle] };
                newHeaders.Add(temp);
            }
            if (CkBxLocation.Checked)
            {
                temp = new String[] { "Location", headers[hasLocation] };
                newHeaders.Add(temp);
            }
            if (CkBxDepartment.Checked)
            {
                temp = new String[] { "Department", headers[hasDepartment] };
                newHeaders.Add(temp);
            }

            //create new DataTable 
            DataTable newDataTable = new DataTable();
            //add headers to each column in the new DataTable
            for (int i = 0; i < newHeaders.Count; i++)
            {
                newDataTable.Columns.Add(newHeaders[i][0]);
            }
            saveDataLabel.Text = saveDataText + "..";
            //iterate through every row in the old DataTable to be inputted into the new DataTable
            foreach (DataRow row in dataTable.Rows)
            {
                //create a new row in the new DataTable
                DataRow newRow = newDataTable.NewRow();
                //iterate through each column in the current row in the old DataTable
                foreach (DataColumn column in dataTable.Columns)
                {
                    //iterate through the new headers array because i need to know when the correct column is found 
                    for (int i = 0; i < newHeaders.Count; i++)
                    {
                        //if the old DataTable's current column name equals the new Headers column name then assign that cell value to the new row's cell
                        if (column.ColumnName.Equals(newHeaders[i][1]))
                        {
                            newRow[newHeaders[i][0]] = row[column.ColumnName];
                            break;
                        }
                    }
                }
                //add the new row to the new DataTable
                newDataTable.Rows.Add(newRow);
            }
            saveDataLabel.Text = saveDataText + "...";

            //create .xlsx file
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[1];
            saveDataLabel.Text = saveDataText + "";

            //headers
            xlWorkSheet.Cells[1, 1] = "Calling As";
            xlWorkSheet.Cells[2, 1] = "Phone # Displayed";
            xlWorkSheet.Cells[3, 1] = "Name Drop";
            xlWorkSheet.Cells[4, 1] = "Engagements Needed";
            xlWorkSheet.Cells[5, 1] = "Engagements per Day";
            xlWorkSheet.Cells[6, 1] = "Current Engagements";
            xlWorkSheet.Cells[7, 1] = "Business Hours";
            xlWorkSheet.Cells[1, 2] = callingAs;
            xlWorkSheet.Cells[2, 2] = phoneNumberDisplayed;
            xlWorkSheet.Cells[3, 2] = nameDrop;
            xlWorkSheet.Cells[4, 2] = engagementsNeeded;
            xlWorkSheet.Cells[5, 2] = engagementsPerDay;
            xlWorkSheet.Cells[6, 2] = "waiting for calls...";
            xlWorkSheet.Cells[7, 2] = businessHours + " " + timeZone;

            //employee call list
            for (int i = 0; i < newHeaders.Count; i++)
            {
                xlWorkSheet.Cells[9, i + 1] = newHeaders[i][0];
            }
            saveDataLabel.Text = saveDataText + ".";
            for (int i = 0; i < newDataTable.Rows.Count; i++)
            {
                for (int j = 0; j < newDataTable.Columns.Count; j++)
                {
                    xlWorkSheet.Cells[i + 10, j + 1] = newDataTable.Rows[i][j].ToString();
                }
            }

            //add Final Result column into worksheet which will determine if Employee Passed, Failed, or Didn't Answer the Phone.
            xlWorkSheet.Cells[9, newHeaders.Count + 1] = "Result";

            //reformat column size to display all the data cleanly
            xlWorkSheet.Columns.AutoFit();
            saveDataLabel.Text = saveDataText + "..";

            SaveFileDialog fileStream = new SaveFileDialog();
            fileStream.FileName = companyName.Trim() + " RSE " + currentYear + " Call List.xlsx";
            fileStream.DefaultExt = ".xlsx";
            fileStream.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            DialogResult result = fileStream.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileName = fileStream.FileName;
                xlWorkBook.SaveAs(fileName);
            }
            saveDataLabel.Text = saveDataText + "...";
            xlWorkBook.Close();
            xlApp.Quit();
            this.Close();
            saveDataLabel.Text = saveDataText;
            saveDataLabel.Visible = false;
        }
    }
}

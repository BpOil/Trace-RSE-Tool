using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;

namespace csv_test_6._28._18
{
    public partial class MakeCalls : Form
    {
        private string filePath;
        public int i = -1;
        private int lastRow;
        private bool displaySpoofcard;
        public bool failed;
        public bool continueProgram;
        public ChromeDriver driver;

        private string companyName;
        private string hours;
        private string callingAs;
        private string numberDisplayed;
        private string nameDrop;
        private string engagementsNeeded;
        private double currentEngagements;
        private double todaysEngagements;

        private DataTable dataTable;

        private bool hasName;
        private bool hasTitle;
        private bool hasExtension;
        private bool hasEmail;
        private bool hasLocation;
        private bool hasDepartment;
        private bool hasToday;

        public MakeCalls()
        {
            InitializeComponent();
            //disable save data label that takes up the whole window
            saveDataLabel.Visible = false;
            //show SpoofCard Web Page initially for the user to log in
            displaySpoofcard = true;
            //assume that create class method didn't failed
            failed = false;
            continueProgram = true;
            //assume that no column exist
            hasName = false;
            hasTitle = false;
            hasExtension = false;
            hasEmail = false;
            hasLocation = false;
            hasDepartment = false;
            hasToday = false;
            //assume no engagements have been made
            currentEngagements = 0;
            todaysEngagements = 0;

            //create new DataTable to get Excel Sheet data
            var dt = new DataTable();
            try
            { //open Call List Excel File
                OpenFileDialog openFile = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true };
                DialogResult result = openFile.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //store file path 
                    filePath = openFile.FileName;
                }
                var file = new FileInfo(filePath);

                //read Excel File and store it into a DataTable
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

                //don't use headers as column names
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                };

                var dataSet = reader.AsDataSet(conf);
                //stores data into DataTable here
                dt = dataSet.Tables[0];
                reader.Close();
                fs.Close();
            }
            catch
            {
                MessageBox.Show("The file could not be loaded", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //set failed to true so Main Class doesn't display the Make Calls form
                failed = true;
                //exit method to return to the Main Class
                return;
            }

            //check that the Call List was made by the RSE Tool so that there are no errors later 
            if (!dt.Rows[0][0].Equals("Calling As") & !dt.Rows[1][0].Equals("Phone # Displayed") & !dt.Rows[2][0].Equals("Name Drop")
                & !dt.Rows[3][0].Equals("Engagements Needed") & !dt.Rows[4][0].Equals("Engagements per Day") & !dt.Rows[5][0].Equals("Current Engagements")
                 & !dt.Rows[6][0].Equals("Business Hours"))
            {
                MessageBox.Show("Please use a Call List that was created by the RSE Tool.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //set failed to true so Main Class doesn't display the Make Calls form
                failed = true;
                //exit method to return to the Main Class
                return;
            }

            //enable big label that takes up the whole window and ask the user to start the Chrome Browser so they can log into SpoofCard
            saveDataLabel.Text = "Start the Chrome Driver.";
            btnContinueProgram.Text = "Press Me!";
            saveDataLabel.Visible = true;


            //get the company name from the file path
            string[] splitter = filePath.Split('\\');
            splitter = splitter[splitter.Length - 1].Split(new string[] { "RSE" }, StringSplitOptions.None);
            companyName = splitter[0].Trim();
            //store the business hours
            hours = dt.Rows[6][1].ToString();
            //store the company that the trace employee is possing as
            callingAs = dt.Rows[0][1].ToString();
            //store the number that the trace employee is spoofing as
            numberDisplayed = dt.Rows[1][1].ToString();
            //store the name drop 
            nameDrop = dt.Rows[2][1].ToString();
            //store the number of engagements needed for project
            engagementsNeeded = dt.Rows[3][1].ToString();
            //see if current engagements if not zero
            if (!dt.Rows[5][1].Equals("waiting for calls..."))
            {
                currentEngagements = (double)dt.Rows[5][1];
            }

            //insert the above variables into the windows form
            txtBxCompanyName.Text = companyName;
            txtBxHours.Text = hours;
            txtBxCallingAs.Text = callingAs;
            txtBxNumberDisplayed.Text = numberDisplayed;
            txtBxNameDrop.Text = nameDrop;
            txtBxEngagementsNeeded.Text = engagementsNeeded;
            txtBxCurrentEngagements.Text = currentEngagements.ToString();


            //make new DataTable that contains just the employee information 
            dataTable = new DataTable();
            dataTable = dt;
            //delete the static data from the Data Table (e.g. calling as company, the phone # displayed, the name drop, the business hours, the engagements needed, etc.)
            dataTable.Rows[0].Delete();
            dataTable.Rows[1].Delete();
            dataTable.Rows[2].Delete();
            dataTable.Rows[3].Delete();
            dataTable.Rows[4].Delete();
            dataTable.Rows[5].Delete();
            dataTable.Rows[6].Delete();
            dataTable.Rows[7].Delete();
            dataTable.AcceptChanges();


            //get the headers for the employee info
            List<string> headers = new List<string>();
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                string header = dataTable.Rows[0][j].ToString();
                //if first row current cell contains AM or PM then manipulate the date to just include the date and no timestamp 
                if (header.Contains("AM") | header.Contains("PM"))
                {
                    splitter = header.Split(' ');
                    header = splitter[0];
                    dataTable.Rows[0][j] = header;
                }

                headers.Add(header);
            }
            dataTable.AcceptChanges();
            //determine what column actually exist so the program knows what data to display on the windows form
            for (int j = 0; j < headers.Count; j++)
            {
                if (headers.Contains("Name"))
                {
                    hasName = true;
                }
                if (headers.Contains("Title"))
                {
                    hasTitle = true;
                }
                if (headers.Contains("Extension"))
                {
                    hasExtension = true;
                }
                if (headers.Contains("Email"))
                {
                    hasEmail = true;
                }
                if (headers.Contains("Location"))
                {
                    hasLocation = true;
                }
                if (headers.Contains("Department"))
                {
                    hasDepartment = true;
                }
                //check if the table contains today's date so the program knows if it should add the column later
                if (headers.Contains(DateTime.Now.ToShortDateString()))
                {
                    hasToday = true;
                    //if today's date is found in the Employee's DataTable headers then calculate the number of engagements found that day
                    todaysEngagements = getTodaysEngagements(dataTable);
                    txtBxTodayEngagements.Text = todaysEngagements.ToString();
                }
            }
            //make the first row of the Data Table the Column Headers
            int index = 0;
            foreach (DataColumn column in dataTable.Columns)
            {
                column.ColumnName = dataTable.Rows[0][index].ToString();
                index++;
            }
            //if today's date is not included in the Excel table then add that column
            if (hasToday == false)
            {
                dataTable.Columns.Add(DateTime.Now.ToShortDateString(), typeof(string));
                txtBxTodayEngagements.Text = "0";
            }
            //delete the first row now that it's the headers 
            dataTable.Rows[0].Delete();
            dataTable.AcceptChanges();
            //get the number of rows in the new DataTable that contains just the employee info
            lastRow = dataTable.Rows.Count;
        }

        public void insertEmployeeInfo()
        {
            //if the current row of the Employee DataTable is the last row then tell the user that it is going back to the top of the DataTable and then do so
            if (i == lastRow - 1)
            {
                i = 0;
                MessageBox.Show("Reached bottom of the Employee List. Going back to the top of the Employee List.", "Recycling Employees");
            }
            else //if the last row is not active then increment i so that the next row is selected and input into the window form text boxes
            {
                i++;
            }
            if (dataTable.Rows[i]["Result"].Equals("PASSED") | dataTable.Rows[i]["Result"].Equals("FAILED") | dataTable.Rows[i][DateTime.Now.ToShortDateString()].Equals("Voicemail") | string.IsNullOrWhiteSpace(dataTable.Rows[i]["Name"].ToString()))
            { //if current row doesn't have a name, or has already had a pass/fail engagement with the user or has called them today then skip this row and go to the next row.
                insertEmployeeInfo();
            }
            else
            { //if employee has NO ANSWER at the moment or hasn't been called then go to main page of SpoofCard

                driver.Navigate().GoToUrl("https://www.spoofcard.com/account");
                driver.SwitchTo().DefaultContent();
                //set the phone number that will be displayed when calling client
                driver.FindElementByXPath("//*[@id='display_address']").SendKeys(numberDisplayed);
                //set the client's phone number 
                driver.FindElementByXPath("//*[@id='call_destination_address']").SendKeys(dataTable.Rows[i]["Phone"].ToString());
                //click make new call button
                driver.FindElementByXPath("//*[@id='create_call']/fieldset[4]/button").Click();
                //wait for SpoofCard Pin WebPage to load
                System.Threading.Thread.Sleep(2000);
                //get the SpoofCard Pin
                string spoofcardPin = "####";
                try
                {
                    spoofcardPin = driver.FindElement(By.XPath("//*[@id='myaccount-step2']/div/div[1]")).Text.Substring(36);
                } catch (Exception)
                {
                    displaySpoofcard = false;
                    btnSpoofCard.PerformClick();
                }
                //insert the employee info into the window's form if that data is found in the employee DataTable
                if (hasName == true)
                {
                    txtBxEmployeeName.Text = dataTable.Rows[i]["Name"].ToString();
                }
                if (hasTitle == true)
                {
                    txtBxEmployeeTitle.Text = dataTable.Rows[i]["Title"].ToString();
                }
                if (hasExtension == true)
                {
                    txtBxExtension.Text = dataTable.Rows[i]["Extension"].ToString();
                }
                if (hasEmail == true)
                {
                    txtBxEmail.Text = dataTable.Rows[i]["Email"].ToString();
                }
                if (hasLocation == true)
                {
                    txtBxEmployeeLocation.Text = dataTable.Rows[i]["Location"].ToString();
                }
                if (hasDepartment == true)
                {
                    txtBxEmployeeDepartment.Text = dataTable.Rows[i]["Department"].ToString();
                }
                //insert the SpoofCard Pin into the window's form
                spoofcardPinLabel.Text = spoofcardPin;
            }
            //clear converstation textbox every time this method is called
            txtBxConversation.Clear();
        }

        private void btnSkip_Click(object sender, EventArgs e)
        {
            //if skip button is clicked then just move on the next employee
            insertEmployeeInfo();
        }

        private void btnVoicemail_Click(object sender, EventArgs e)
        {
            //if Voicemail button is clicked then increment totalEngagements and today's Engagements by 1 third
            currentEngagements = currentEngagements + 0.333333333333333333333333333333333333333333;
            todaysEngagements = todaysEngagements + 0.333333333333333333333333333333333333333333;
            //update window's form info 
            txtBxCurrentEngagements.Text = currentEngagements.ToString();
            txtBxTodayEngagements.Text = todaysEngagements.ToString();
            //assign Employee Result Cell as "NO ANSWER"
            dataTable.Rows[i]["Result"] = "NO ANSWER";
            //assign Employee's today's cell value as "Voicemail"
            dataTable.Rows[i][DateTime.Now.ToShortDateString()] = "Voicemail";
            //move on to the next employee row
            insertEmployeeInfo();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if "Preview Data" button is clicked then display the current state of the Employee DataTable in it's own window
            Preview preview = new Preview();
            preview.dataTable = dataTable;
            preview.makeCalls = this;
            preview.ShowDialog();
        }

        private void btnPassed_Click(object sender, EventArgs e)
        {
            //if conversation text box is empty then tell the user the enter some info about the conversation and end the method
            if (string.IsNullOrWhiteSpace(txtBxConversation.Text))
            {
                MessageBox.Show("Please give a description about the conversation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //increment the totalEngagements and today's engagements by 1
            currentEngagements = currentEngagements + 1;
            todaysEngagements = todaysEngagements + 1;
            //update window's form info
            txtBxCurrentEngagements.Text = currentEngagements.ToString();
            txtBxTodayEngagements.Text = todaysEngagements.ToString();
            //set the Employee's Final result as "PASSED"
            dataTable.Rows[i]["Result"] = "PASSED";
            //set the Employee's today's cell as the text found in the conversation info textbox
            dataTable.Rows[i][DateTime.Now.ToShortDateString()] = txtBxConversation.Text;
            //move on to the next employee row
            insertEmployeeInfo();
        }

        private void btnFailed_Click(object sender, EventArgs e)
        {
            //if conversation text box is empty then tell the user the enter some info about the conversation and end the method
            if (string.IsNullOrWhiteSpace(txtBxConversation.Text))
            {
                MessageBox.Show("Please give a description about the conversation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //increment the totalEngagements and today's engagements by 1
            currentEngagements = currentEngagements + 1;
            todaysEngagements = todaysEngagements + 1;
            //update the window's form info
            txtBxCurrentEngagements.Text = currentEngagements.ToString();
            txtBxTodayEngagements.Text = todaysEngagements.ToString();
            //set the Employee's final result cell as "FAILED"
            dataTable.Rows[i]["Result"] = "FAILED";
            //set the Employee's today's cell as the text found in the conversation info textbox
            dataTable.Rows[i][DateTime.Now.ToShortDateString()] = txtBxConversation.Text;
            //move onto the next employee row
            insertEmployeeInfo();
        }

        private void btnSaveData_Click(object sender, EventArgs e)
        { //if Save Data button is clicked then do the following
            //get the file name from the file path
            string[] fileName = filePath.Split('\\');
            string labelText = "Saving Data to \"" + fileName[fileName.Length - 1] + "\"";
            //set the saveDataLabel text to "Saving Data to [filename].xlsx". This is the big label that takes up the whole window
            saveDataLabel.Text = labelText;
            //display the saveDatalabel 
            saveDataLabel.Visible = true;

            //start a new Excel Application
            Excel.Application xlApp = new Excel.Application();
            //append a period to the end of the saveDataLabel to show the user that code is being executed
            saveDataLabel.Text = labelText + ".";
            //open excel workbook in the background
            Excel.Workbook wb = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet ws = wb.Worksheets[1];
            //append a period to the end of the saveDataLabel to show the user that code is being executed
            saveDataLabel.Text = labelText + "..";

            //update Current Engagement cell
            if (currentEngagements != 0)
            {
                ws.Cells[6, 2] = currentEngagements;
            }


            //get the Employee DataTable's Headers
            List<string> dtHeaders = new List<string>();
            foreach (DataColumn col in dataTable.Columns)
            {
                dtHeaders.Add(col.ColumnName);
                Console.WriteLine("Data Table Headers: " + col.ColumnName);
            }
            Console.WriteLine("");


            var cellValue = "";
            //add missing Headers to Excel Sheet
            for (int k = dataTable.Columns.IndexOf("Result"); k < dataTable.Columns.Count; k++)
            {
                try
                {
                    //get current Excel Sheet's header name
                    cellValue = (string)(ws.Cells[9, k + 1] as Excel.Range).Value;
                }
                catch (Exception)
                {
                    //if program returns error then the header is probably a date so convert that to a string and get just the date and no timestamp
                    DateTime temp = (ws.Cells[9, k + 1] as Excel.Range).Value;
                    cellValue = temp.ToShortDateString();
                }
                if (string.IsNullOrEmpty(cellValue) || !cellValue.Equals(dtHeaders[k]))
                { //if current cell in Excel header row is empty then add the Employee's DataTable's current header cell value
                    ws.Cells[9, k + 1] = dtHeaders[k];
                }
            }

            //add missing table data to the Excel Sheet
            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                for (int k = dataTable.Columns.IndexOf("Result"); k < dataTable.Columns.Count; k++)
                {
                    //get Excels current row's cell value
                    cellValue = (string)(ws.Cells[j + 10, k + 1] as Excel.Range).Value;
                    if (string.IsNullOrEmpty(dataTable.Rows[j][k].ToString()) || string.IsNullOrEmpty(cellValue) || !cellValue.Equals(dataTable.Rows[j][k]))
                    { //if current excel row's cell value is null or doesn't equal the Employee's dataTable value then update the current Excel row's cell
                        ws.Cells[j + 10, k + 1] = dataTable.Rows[j][k];
                    }
                }
            }
            Console.WriteLine("Done!");
            //append a period to the end of the saveDataLabel to show the user that code is being executed
            saveDataLabel.Text = labelText + "...";

            //save the workbook
            wb.Save();
            //close the workbook
            wb.Close();
            //end the Excel process that is running in the background. sometimes this doesn't work...
            xlApp.Quit();
            //let the user know that the workbook was updated and saved
            saveDataLabel.Text = "Done!";
            System.Threading.Thread.Sleep(1000);
            //disable the saveDataLabel to display the employee info window form again
            saveDataLabel.Visible = false;
        }

        private double getTodaysEngagements(DataTable dt)
        { //if method calculates the number of engagements that are recorded for today's date 
            int columnIndex = 0;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                //find the column that contains today's date
                if (dt.Rows[0][i].Equals(DateTime.Now.ToShortDateString()))
                {
                    //store that column index that contains today's date
                    columnIndex = i;
                }
            }

            //assume there are no engagements for today
            double todaysEngagements = 0;
            //iterate through all the cells in the today's column
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                //if the cell is empty then move on to the next cell in the column
                if (string.IsNullOrEmpty(dt.Rows[i][columnIndex].ToString()) | string.IsNullOrWhiteSpace(dt.Rows[i][columnIndex].ToString()))
                {
                    continue;
                }
                //if the cell contains Voicemail then add 1 third to today's engagement
                else if (dt.Rows[i][columnIndex].ToString().Equals("Voicemail"))
                {
                    todaysEngagements = todaysEngagements + 0.333333333333333333333333333333333333333333;
                }
                //if cell isn't empty and doesn't contain Voicemail then add 1 to today's engagement because it will contain the conversation that the user had with a client 
                else
                {
                    todaysEngagements = todaysEngagements + 1;
                }
            }
            //return the number of engagements that happened today
            return todaysEngagements;
        }

        private void btnContinueProgram_Click(object sender, EventArgs e)
        { //continueProgram is assigned true initially so that Chrome Browser is only started once  
            if (continueProgram == true)
            {
                //let the user know that the Chrome Browser is starting 
                saveDataLabel.Text = "Starting Chrome Browser. Please Wait...";
                //make sure that the Window's form UI is updated
                saveDataLabel.Invalidate();
                saveDataLabel.Update();
                saveDataLabel.Refresh();
                //disable the continueProgram button so user can't click it before logging in 
                btnContinueProgram.Visible = false;
                //force window's form to update
                Application.DoEvents();
                //create service object to specify that I don't want the chrome command prompt box to display
                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                //create Chrome Options object to specify where I want the Chrome Browser to be located
                var options = new ChromeOptions();
                options.AddArgument("--window-position=800,10");
                //start the Chrome Browser
                driver = new ChromeDriver(service, options);
                //edit the Chrome Browser to implicately wait when I use the method .Until()
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                //go to the SpoofCard Login Web Page
                driver.Navigate().GoToUrl("https://www.spoofcard.com/login");

                //tell the user to log into SpoofCard
                saveDataLabel.Text = "Please log into SpoofCard.";
                //change button to say "I Logged In" so user knows to click it after logging in
                btnContinueProgram.Text = "I Logged In";
                //enable that button that the user will click after logging in to SpoofCard
                btnContinueProgram.Visible = true;
                //set continueProgram to false so next time the user clicks "I Logged In" button it doesn't create a new Chrome Browser 
                //and disables the big label that takes up the entire windows form and inserts the employee info into the window's form
                continueProgram = false;
            }
            else
            {
                //after user logs into SpoofCard. Hide the Chrome Browser 
                btnSpoofCard_Click(sender, e);
                //disable the big label that takes up the entire window's form
                saveDataLabel.Visible = false;
                //disable the continue program button so no one can click it randomly
                btnContinueProgram.Visible = false;

                //insert the first employee row info into the window's form
                insertEmployeeInfo();
            }
        }

        private void btnSpoofCard_Click(object sender, EventArgs e)
        { //when this button is clicked, either hide or display the Chrome Browser that contains the SpoofCard Web Page
            if (displaySpoofcard == true)
            {
                //when displaySpoofCard is true then I want to hide the Chrome Browser and change the button text to "Display SpoofCard"
                //so next time the button is clicked the program will display the Chrome Browser
                driver.Manage().Window.Position = new Point(-35000, -35000);
                //change displaySpoofCard so next time the SpoofCard button is click the other condition is executed
                displaySpoofcard = false;
                btnSpoofCard.Text = "Display SpoofCard";
            }
            else
            {
                //when displaySpoofCard is false then I want to show the Chrome Browser and change the button text to "Hide SpoofCard"
                //so next time the button is clicked the program will hide the Chrome Browser
                driver.Manage().Window.Position = new Point(750, 10);
                //change displaySpoofCard so next time the SpoofCard button is click the other condition is executed
                displaySpoofcard = true;
                btnSpoofCard.Text = "Hide SpoofCard";
            }
        }
    }
}

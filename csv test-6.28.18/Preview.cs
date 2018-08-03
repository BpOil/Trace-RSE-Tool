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
    public partial class Preview : Form
    {
        public Preview()
        {
            InitializeComponent();
        }

        public string ContactPath { get; set; }

        public string FileType { get; set; }

        public DataTable dataTable;

        private void Preview_Load(object sender, EventArgs e)
        {
            Read displayData = new Read();
            displayData.NameFile = ContactPath;
            if (FileType.Equals("Word"))
            {
                string[] copyHeader = displayData.TableHeader();
                string[,] displayArray = displayData.WordDoc();
                dataTable = new DataTable();
                for (int i = 0; i < copyHeader.Length; i++)
                {
                    dataTable.Columns.Add(copyHeader[i], typeof(String));
                }


                for (int i = 0; i < (displayArray.Length / copyHeader.Length); i++)
                {
                    DataRow row = dataTable.NewRow();
                    for (int j = 0; j < copyHeader.Length; j++)
                    {
                        row[copyHeader[j]] = displayArray[i, j];
                    }
                    dataTable.Rows.Add(row);
                }


            }
            else if (FileType.Equals("Excel"))
            {
                DataSet result;
                var file = new FileInfo(ContactPath);
                IExcelDataReader reader;
                FileStream fs = File.Open(ContactPath, FileMode.Open, FileAccess.Read);
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
                        UseHeaderRow = true
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

                dataTable = new DataTable();
                List<String> headers = new List<String>();
                if (firstNameColumn != -1 & lastNameColumn != -1)
                {
                    dataTable.Columns.Add("Name", typeof(String));
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (!column.ColumnName.Equals(firstNameColumnName) & !column.ColumnName.Equals(lastNameColumnName))
                        {
                            dataTable.Columns.Add(column.ColumnName, typeof(String));
                        }
                    }


                    foreach (DataRow dr in dt.Rows)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Name"] = dr[firstNameColumnName].ToString() + " " + dr[lastNameColumnName];
                        foreach (DataColumn column in dt.Columns)
                        {
                            if (!column.ColumnName.Equals(firstNameColumnName) & !column.ColumnName.Equals(lastNameColumnName))
                            {
                                row[column.ColumnName] = dr[column.ColumnName];
                            }
                        }
                        dataTable.Rows.Add(row);
                    }

                    foreach (DataColumn column in dataTable.Columns)
                    {
                        headers.Add(column.ColumnName);
                    }
                }
                else
                {
                    dataTable = dt;
                    headers = initialHeaders;
                }

            }
            dataGridView.DataSource = dataTable;
        } //end of Preview Class
    }
}

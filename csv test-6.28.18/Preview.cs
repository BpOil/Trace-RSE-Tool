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

        public DataTable dataTable { get; set; }

        public MakeCalls makeCalls { get; set; }

        private void Preview_Load(object sender, EventArgs e)
        {
            dataGridView.DataSource = dataTable;
        } 

        private void dataGridView_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView.Visible = false;
            makeCalls.i = e.RowIndex - 1;
            makeCalls.insertEmployeeInfo();
            this.Close();
        }
    } //end of Preview Class
}

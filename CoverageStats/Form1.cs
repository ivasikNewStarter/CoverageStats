using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace CoverageStats
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void buttonImportFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.Multiselect = false;
            if (browse.ShowDialog() == DialogResult.OK)
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + browse.FileName + @";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0"""; //C:\Users\u0139221\Desktop\file.xlsx.
                conn.Open();

                DataTable a = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                OleDbCommand command = new OleDbCommand
               (
                    "SELECT "+txtColumnName.Text+"  FROM[" + txtSheetName.Text+"$]", conn
                );

                DataTable dsnew = new DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);

                adapter.Fill(dsnew);

                dataGridView1.DataSource = dsnew;

                StringBuilder result = new StringBuilder();

                /* select count(*)from oa.organizations o
                 where o.org_tier_id = 1 and (o.org_nda_id in ())*/

                foreach (DataRow row in dsnew.Rows)
                {
                    result.Append(row[0]);
                    result.Append("\n");
                }
                string coverageList = Splitter.splitEndLineToInClauseOracle(result.ToString(), "o.org_nda_id",true);
                string SQL = string.Format("select count(*)from oa.organizations o " +
                 " where o.org_tier_id = 1 and {0} ", coverageList);

                txtResult.Text = SQL; //result.ToString();
            }
        }

 

        private void buttonCopyClipboard_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtResult.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.Multiselect = true;
            DialogResult result = browse.ShowDialog();

        }


    }
    }


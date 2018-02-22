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
using CEMAutomationLib.Common;

namespace CoverageStats
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        public void buttonImportFile_Click(object sender, EventArgs e)
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
                 " where o.org_tier_id = 1 and ({0}) ", coverageList);

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

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void txtSheetName_TextChanged(object sender, EventArgs e)
        {

        }

        public StringBuilder loadExcelColumn(string columnName, string sheetName)
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
                    "SELECT " + columnName + "  FROM[" + sheetName + "$]", conn
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

                return result;
                //string coverageList = Splitter.splitEndLineToInClauseOracle(result.ToString(), "o.orgid", true);
                //string SQL = string.Format("select ORGID,ORG_DATA_PROVIDER from OA.ORGANIZATIONS o left join ife.entityids e on e.entityid = o.orgid " +
                // " where orgid  in ({0}) ", coverageList);

               // txtResult.Text = SQL; //result.ToString();
            }
            return new StringBuilder();
        }


        public void runLEI_Click(object sender, EventArgs e)
        {
            StringBuilder listRecords = loadExcelColumn(txtColumnName.Text, txtSheetName.Text);
            StringBuilder result = new StringBuilder();
            string coverageList = Splitter.splitEndLineToInClauseOracle(listRecords.ToString(), "ife.entityid", true);
            //string SQL = string.Format("select ORGID,ORG_DATA_PROVIDER from OA.ORGANIZATIONS o left join ife.entityids e on e.entityid = o.orgid " +
            // " where {0} ", coverageList);
            string SQL = string.Format("select * from (select ife.entityid, ORG_DATA_PROVIDER, ad.USERname, ife.AUDITDATETIME " +
            " from ife.audittrail ife left join OA.ORGANIZATIONS oa on ife.entityid = oa.orgid " +
            " left join admin.users ad on ife.analystuserid = ad.userid where attributeid = 499 and newvalue like '%ID_TYPE>38</ID_TYPE%' " +
             " and  {0}  order by ife.AUDITDATETIME DESC)" +
            "where rownum = 1;" , coverageList);

            txtResult.Text = SQL; //result.ToString();
        }

        private void runCA_Click(object sender, EventArgs e)
        {

            StringBuilder listRecords = loadExcelColumn(txtColumnName.Text, txtSheetName.Text);
            StringBuilder result = new StringBuilder();
            foreach (string record in listRecords.ToString().Split('\n'))
            {
                result.Append("'Deals No: ");
                result.Append(record);
                result.Append("',");
            }         
            string dealsList = result.ToString(0, result.Length - 1);
            string SQL = string.Format("select req.REQUEST_ID, cc.COUNTRYdesc, req.ENTITY_ID, req.REQUEST_DATE, status.TASK_STATUS, req.ADDITIONAL_FIELD_VALUE_2, req.ANALYST_COMMENT " +
                "from REQUESTTOOL.REQUEST_QUEUE req left join TASKTOOL.TASK_STATUS_TYPE_CODES status on req.status_id = status.TASK_STATUS_ID " +
                "join REQUESTTOOL.request_detail reqd on reqd.request_id = req.request_id join domains.countrycodes cc on reqd.ATTRIBUTE_VALUE = cc.COUNTRYCODE and reqd.attribute_id = 259 " +
                "where ADDITIONAL_FIELD_VALUE_2 in ({0}) and req.STATUS_ID in (1,2,3,4) and req.ADDITIONAL_FIELD_Name_1 = 'TARGET' AND cc.COUNTRYCODE not in ('ID','CA','BD', 'GB', 'AU', 'US', 'KR', 'SG', 'CN', 'HK', 'KY', 'BVI', 'BS', 'JP', 'IN', 'TW', 'VN', 'PH', 'MY','TH','NZ', 'IE', 'LK', 'ZA' )", dealsList);

            txtResult.Text = SQL; //result.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataBase db = new DataBase(DataBase.DATABASE_ORACLE);
            bool testConn = db.testConnection();
            DataSet result = db.getDataSet(txtResult.Text.ToString(), "Export");
            ExcelManager excel = new ExcelManager();
            excel.GenerateWorkbook(result);
            excel.SaveIn("C:\\Users\\u0139221\\Desktop\\ImportSQl.xlsx");
            excel.CleanAndClose();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
    }


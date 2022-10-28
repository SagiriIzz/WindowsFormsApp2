using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection = null;
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;
        private bool newRowAdding = false;
     
        public Form1()
        {
            InitializeComponent();
        }
        private void LoadData()
        {
            try {
                sqlDataAdapter = new SqlDataAdapter("select *, 'Delete' as [Delete]  from ludi", sqlConnection);
                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);
                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();
                dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet, "ludi" ) ;
                dataGridView1.DataSource = dataSet.Tables["ludi"];
                for(int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[9, i] = linkCell;

                }
                    
            }
            catch(Exception ex) {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReloadData()
        {
            try
            {
                dataSet.Tables["ludi"].Clear();
                sqlDataAdapter.Fill(dataSet, "ludi");
                dataGridView1.DataSource = dataSet.Tables["ludi"];
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[9, i] = linkCell;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""C:\Users\Sagiri Izumi\source\repos\WindowsFormsApp2\WindowsFormsApp2\Database1.mdf"";Integrated Security=True") ;
       sqlConnection.Open();
            LoadData();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 9)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строчку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;
                            dataGridView1.Rows.RemoveAt(rowIndex);
                            dataSet.Tables["ludi"].Rows[rowIndex].Delete();
                            sqlDataAdapter.Update(dataSet, "ludi");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView1.Rows.Count - 2;
                        DataRow row = dataSet.Tables["ludi"].NewRow();

                        row["fa"] = dataGridView1.Rows[rowIndex].Cells["fa"].Value;
                        row["im"] = dataGridView1.Rows[rowIndex].Cells["im"].Value;
                        row["otch"] = dataGridView1.Rows[rowIndex].Cells["otch"].Value;
                        row["organiz"] = dataGridView1.Rows[rowIndex].Cells["organiz"].Value;
                        row["podraz"] = dataGridView1.Rows[rowIndex].Cells["podraz"].Value;
                        row["dol"] = dataGridView1.Rows[rowIndex].Cells["dol"].Value;
                        row["hours"] = dataGridView1.Rows[rowIndex].Cells["hours"].Value;
                        row["day"] = dataGridView1.Rows[rowIndex].Cells["day"].Value;

                        dataSet.Tables["ludi"].Rows.Add(row);
                        dataSet.Tables["ludi"].Rows.RemoveAt(dataSet.Tables["ludi"].Rows.Count -1);
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                        dataGridView1.Rows[e.RowIndex].Cells[9].Value = "Delete";
                        sqlDataAdapter.Update(dataSet, "ludi");
                        newRowAdding = false;
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet.Tables["ludi"].Rows[r]["fa"] = dataGridView1.Rows[r].Cells["fa"].Value;
                        dataSet.Tables["ludi"].Rows[r]["im"] = dataGridView1.Rows[r].Cells["im"].Value;
                        dataSet.Tables["ludi"].Rows[r]["otch"] = dataGridView1.Rows[r].Cells["otch"].Value;
                        dataSet.Tables["ludi"].Rows[r]["organiz"] = dataGridView1.Rows[r].Cells["organiz"].Value;
                        dataSet.Tables["ludi"].Rows[r]["podraz"] = dataGridView1.Rows[r].Cells["podraz"].Value;
                        dataSet.Tables["ludi"].Rows[r]["dol"] = dataGridView1.Rows[r].Cells["dol"].Value;
                        dataSet.Tables["ludi"].Rows[r]["hours"] = dataGridView1.Rows[r].Cells["hours"].Value;
                        dataSet.Tables["ludi"].Rows[r]["day"] = dataGridView1.Rows[r].Cells["day"].Value;
                        sqlDataAdapter.Update(dataSet, "ludi");
                        dataGridView1.Rows[e.RowIndex].Cells[9].Value = "Delete";
                    }
                    ReloadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding == false) {
                    newRowAdding = true;
                    int lastRow = dataGridView1.Rows.Count - 2;

                    DataGridViewRow row = dataGridView1.Rows[lastRow];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[9, lastRow] = linkCell;
                    row.Cells["Delete"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding == false) { 
                int rowIndex = dataGridView1.SelectedCells[0].RowIndex;
                    DataGridViewRow editingRow = dataGridView1.Rows[rowIndex];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[9, rowIndex] = linkCell;
                    editingRow.Cells["Delete"].Value = "Update";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void SaveTabie()
        {
            string path = System.IO.Directory.GetCurrentDirectory() + @"\" + "Save1.xlsx";
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelapp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet;

            for (int i = 1; i < dataGridView1.RowCount + 1; i++)
            {
                for (int j = 1; j < dataGridView1.ColumnCount + 1; j++) {
                    worksheet.Rows[i].Columns[j] = dataGridView1.Rows[i - 1].Cells[j - 1].Value;
                }
            }
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            SaveTabie();
        }
       

}
}

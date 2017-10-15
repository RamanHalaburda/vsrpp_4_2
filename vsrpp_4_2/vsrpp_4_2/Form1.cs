using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace vsrpp_4_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DateTimePicker dtp;
        String[] headerList = {"ФИО","Должность","Дата приёма","Стаж"};

        private void Form1_Load(object sender, EventArgs e)
        {
            dgv.ColumnCount = 4;
            dgv.RowCount = 1;
            dgv.RowHeadersVisible = false;
            dgv.Width = 403;
            dgv.Columns[0].Width = dgv.Columns[1].Width = dgv.Columns[2].Width = dgv.Columns[3].Width = 100;

            for(int i = 0; i < dgv.ColumnCount; ++i)
            {
                dgv.Columns[i].HeaderCell.Value = headerList[i];            
            }

            dtp = new DateTimePicker();
            dtp.Format = DateTimePickerFormat.Short;
            dtp.Visible = false;
            dtp.Width = 100;
            dgv.Controls.Add(dtp);

            dtp.ValueChanged += this.dtp_ValueChanged;
        }

        private void ExportToExcel()
        {
            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Работники";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                
                //Loop through each row and read value from each column. 
                for (int i = 0; i < dgv.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgv.Columns[j].HeaderText;
                        }
                        else
                        {
                            if (dgv.Rows[i].Cells[j].Value != DBNull.Value)
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = dgv.Rows[i].Cells[j].Value.ToString();
                            }
                            else
                            {
                                break;
                            }
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        private void dgv_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            try
            {
                if(dgv.Focused && dgv.CurrentCell.ColumnIndex == 2)
                {
                    dtp.Location = dgv.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false).Location;
                    dtp.Visible = true;
                    /*
                    if(dgv.CurrentCell.Value != DBNull.Value)
                    {
                        dtp.Value = (DateTime) dgv.CurrentCell.Value;
                    }
                    else
                    {
                        dtp.Value = DateTime.Today;
                    }
                    */
                    dtp.Value = DateTime.Today;
                }
                else
                {
                    dtp.Visible = false;
                }
            }
            catch(Exception ex)
            { MessageBox.Show( ex.ToString()); }
        }

        private void dgv_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if(dgv.Focused && dgv.CurrentCell.ColumnIndex == 2)
                {
                    dgv.CurrentCell.Value = dtp.Value.Date;
                }
            }
            catch(Exception ex)
            { MessageBox.Show( ex.ToString()); }
        }

        private void dtp_ValueChanged(object sender, EventArgs e)
        {
            dgv.CurrentCell.Value = dtp.Value.Date;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dgv.RowCount++;
        }
    }
}

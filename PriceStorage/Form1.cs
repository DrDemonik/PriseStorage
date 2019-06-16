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
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace PriceStorage
{
    public partial class Form1 : Form
    {
        System.Data.DataTable tableGryz;
        System.Data.DataTable tableTarif;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    tableGryz = ConvertExcelToDataTable(openFileDialog1.FileName, "Груз");
                    tableTarif = ConvertExcelToDataTable(openFileDialog1.FileName, "Тариф");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            try
            {
                
                if (tableGryz != null && tableTarif != null)
                {
                    var dataTableOut = tableGryz.Clone();
                    dataTableOut.Columns.Add("Начало периода", typeof(System.DateTime));
                    dataTableOut.Columns.Add("Окончание периода", typeof(System.DateTime));
                    dataTableOut.Columns.Add("Кол-во дней хранения", typeof(System.Int32));
                    dataTableOut.Columns.Add("Ставка", typeof(System.Int32));
                    dataTableOut.Columns.Add("Сумма", typeof(System.Double));

                    foreach (var drGryz in tableGryz.Select())
                    {
                        var startDate = dateTimePicker1.Value;
                        var endDate = dateTimePicker2.Value;
                        endDate = new DateTime(endDate.Year, endDate.Month, endDate.Day, 23, 59, 59);

                        var minDate = drGryz[1] != System.DBNull.Value ? Convert.ToDateTime(drGryz[1]) : new DateTime(1,12,1);
                        if (startDate < minDate)
                        {
                            startDate = minDate;
                        }

                        var maxDate = drGryz[2] != System.DBNull.Value ? Convert.ToDateTime(drGryz[2]) : new DateTime(9999, 12, 1);
                        if (endDate > maxDate)
                        {
                            endDate = maxDate;
                        }
                        if (endDate < startDate)
                        {
                            endDate = startDate;
                        }

                        var deltatime = endDate - startDate;
                        var countDay = Math.Floor(deltatime.TotalDays);
                        double sum=0;
                        Int32 stavka = 0;
                        foreach (var drTarif in tableTarif.Select())
                        {
                            var maxDay = drTarif[2] != System.DBNull.Value ? (double)drTarif[2] : countDay;
                            if (countDay > (double)drTarif[1] && countDay >= maxDay)
                            {
                                var s = (maxDay - (double)drTarif[1]) * (double)drTarif[3];
                                sum += s;
                                stavka = Convert.ToInt32(drTarif[3]);
                            }
                            else if (countDay >= (double)drTarif[1] && countDay < maxDay)
                            {
                                var s = (countDay - ((double)drTarif[1]-1)) * (double)drTarif[3];
                                sum += s;
                                stavka = Convert.ToInt32(drTarif[3]);
                            }
                        }

                        var desRow = dataTableOut.NewRow();
                        desRow.ItemArray = drGryz.ItemArray.Clone() as object[];
                        desRow["Начало периода"] = startDate;
                        desRow["Окончание периода"] = endDate;
                        desRow["Кол-во дней хранения"] = (Int32)countDay;
                        desRow["Ставка"] = stavka;
                        desRow["Сумма"] = sum;
                        dataTableOut.Rows.Add(desRow);

                    }
                    dataGridView1.DataSource = dataTableOut;
                }
                else
                {
                    throw new Exception("Данные в таблицах отсутвуют. Загрузите данные из файла!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }   

        public static System.Data.DataTable ConvertExcelToDataTable(string fileName, string mySheetName="")
        {
            System.Data.DataTable dtResult = null;
            int totalSheet = 0;  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         && dataRow["TABLE_NAME"].ToString().Contains(mySheetName)
                                         select dataRow)?.CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; 
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > dateTimePicker2.Value)
            {
                dateTimePicker2.Value = dateTimePicker1.Value;
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker2.Value < dateTimePicker1.Value)
            {
                dateTimePicker1.Value = dateTimePicker2.Value;
            }
        }
    }
}


using System;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ExcelDosyaOkuma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

      

        public DataTable ToDataTable(ExcelApp.Range range, int rows, int cols)
        {
            DataTable table = new DataTable();

            for (int i = 1; i <= rows; i++)
            {
                if (i == 1)
                { 
                    for (int j = 1; j <= cols; j++)
                    {
                       
                        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                            table.Columns.Add(range.Cells[i, j].Value2.ToString());
                        else
                            table.Columns.Add(j.ToString() + ".Sütun");
                    }
                    continue;
                }

                var yeniSatir = table.NewRow();
                for (int j = 1; j <= cols; j++)
                {
                   
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        yeniSatir[j - 1] = range.Cells[i, j].Value2.ToString();
                    else
                        yeniSatir[j - 1] = String.Empty;
                }
                table.Rows.Add(yeniSatir);

            }
            return table;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();

            excelApp.Visible = true;
            object Missing = Type.Missing;
            
            ExcelApp.Workbook workbook = excelApp.Workbooks.Add(Missing);
            ExcelApp.Worksheet sheet1 = (ExcelApp.Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j< dataGridView1.Columns.Count; j++)
            {
                ExcelApp.Range myRange = (ExcelApp.Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j< dataGridView1.Columns.Count; j++)
                {

                    ExcelApp.Range myRange = (ExcelApp.Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();


                }
            }
            
        }

        private void DosyaSec_Click_1(object sender, EventArgs e)
        {
            string Yol;
            string DosyaAdi;
            DataTable dt;
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası | *.xls; *.xlsx; *.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                Yol = file.FileName;

                DosyaAdi = file.SafeFileName;

                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                {
                    MessageBox.Show("Excel yüklü değil.");
                    return;
                }

                ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(Yol);

                ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];

                ExcelApp.Range excelRange = excelSheet.UsedRange;

                int satirSayisi = excelRange.Rows.Count;
                int sutunSayisi = excelRange.Columns.Count;
                dt = ToDataTable(excelRange, satirSayisi, sutunSayisi);

                dataGridView1.DataSource = dt;
                dataGridView1.Refresh();

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            }
            else
            {
                MessageBox.Show("Dosya Seçilemedi.");
            }
        }
    }
}

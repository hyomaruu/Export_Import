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
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;
using System.Data.SqlClient;

namespace what
{


    public partial class Form1 : Form
    {
        public string idterpilih;
        SqlConnection kon = new SqlConnection(what.Properties.Resources.kont.ToString());
        SqlCommand cmd;
        public Form1()
        {
            InitializeComponent();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            for(int i=0;i<dataGridView1.Rows.Count-1;i++)
            {
                kon.Open();
                string id, nis, nama, kelas;
                id = dataGridView1.Rows[i].Cells[0].Value.ToString();
                nis = dataGridView1.Rows[i].Cells[1].Value.ToString();
                nama = dataGridView1.Rows[i].Cells[2].Value.ToString();
                kelas = dataGridView1.Rows[i].Cells[3].Value.ToString();
                SqlCommand com = new SqlCommand("insert into siswa(nis,nama,kelas) values('" + nis + "','" + nama + "','" + kelas+"')", kon);
                MessageBox.Show("Data" + id + "Done!");

                kon.Close();
            }          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            kon.Open();
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XLS files(*.xls)|*.xls";
            sfd.FilterIndex = 2;

            if(sfd.ShowDialog()== DialogResult.OK)
            {
                SqlDataAdapter da = new SqlDataAdapter("select * from siswa", kon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                ExcelFile ef = new ExcelFile();
                ExcelWorksheet ew = ef.Worksheets.Add("Sheet1");
                ew.Cells[0, 0].Value = "DataTable insert example:";

                ew.InsertDataTable(dt, new InsertDataTableOptions()
                {
                    ColumnHeaders = true
                });
                ef.Save(sfd.FileName);
                
            }
            kon.Close();
        }

        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "XLS files(*.xls)|*.xls";
            ofd.FilterIndex = 3;

            if(ofd.ShowDialog()== DialogResult.OK)
            {
                dataGridView1.Columns.Clear();
                ExcelFile ef = ExcelFile.Load(ofd.FileName);
                DataGridViewConverter.ExportToDataGridView(ef.Worksheets.ActiveWorksheet, this.dataGridView1, new ExportToDataGridViewOptions() { ColumnHeaders = true });
            }

        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kur4
{
    public partial class Form8 : Form
    {
        public Form8()
        {
            InitializeComponent();
        }
        SqlConnection connect = new SqlConnection(
"Data Source=MS-SQL\\SQLEXPRESS;Initial Catalog=Kursa4;Integrated Security=True");

        private void Form8_Load(object sender, EventArgs e)
        {
            this.bankiTableAdapter.Fill(this.kursa4DataSet15.Banki);
            dataGridView1.DataSource = this.kursa4DataSet15.Banki;
            this.sotrudnikiTableAdapter.Fill(this.kursa4DataSet17.Sotrudniki);
            dataGridView4.DataSource = this.kursa4DataSet17.Sotrudniki;
            this.otdeliTableAdapter.Fill(this.kursa4DataSet16.Otdeli);
            dataGridView2.DataSource = this.kursa4DataSet16.Otdeli;
            this.typeVkladTableAdapter.Fill(this.kursa4DataSet18.TypeVklad);
            dataGridView3.DataSource = this.kursa4DataSet18.TypeVklad;
            this.vkladiTableAdapter.Fill(this.kursa4DataSet20.Vkladi);
            this.vkladchikiTableAdapter.Fill(this.kursa4DataSet19.Vkladchiki);
            dataGridView5.DataSource = this.kursa4DataSet19.Vkladchiki;
            dataGridView6.DataSource = this.kursa4DataSet20.Vkladi;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < dataGridView6.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView6.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView6.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string phrase = textBox4.Text;
            string[] words = phrase.Split(' ');
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladchiki WHERE Фамилия =" + "'" + words[0] + "'" + "AND Имя =" + "'" + words[1] + "'" + "AND Отчество =" + "'" + words[2] + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView5.DataSource = ds.Tables[0];
        }

        private void button13_Click(object sender, EventArgs e)
        {
            this.vkladiTableAdapter.Fill(this.kursa4DataSet20.Vkladi);
            this.vkladchikiTableAdapter.Fill(this.kursa4DataSet19.Vkladchiki);
            dataGridView5.DataSource = this.kursa4DataSet19.Vkladchiki;
            dataGridView6.DataSource = this.kursa4DataSet20.Vkladi;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView5.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView5.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE НомерСотрудника =" + "'" + Convert.ToInt64(comboBox4) + "'", connect);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
            }
            catch
            {

            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE КодВкладчика =" + "'" + comboBox6.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE КодВклада =" + "'" + comboBox7.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladchiki WHERE КодБанка =" + "'" + comboBox8.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (textBox5.Text.Equals(""))
            {
                MessageBox.Show("Заполните поле");
            }
            else
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Banki WHERE ТелефонБанка =" + "'" + textBox5.Text + "'", connect);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (textBox5.Text.Equals(""))
            {
                MessageBox.Show("Заполните поле");
            }
            else
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Banki WHERE Наименование =" + "'" + textBox5.Text + "'", connect);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.bankiTableAdapter.Fill(this.kursa4DataSet15.Banki);
            dataGridView1.DataSource = this.kursa4DataSet15.Banki;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);

            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Banki WHERE Председатель =" + "'" + comboBox2.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
            {
                MessageBox.Show("Заполните поле");
            }
            else
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Otdeli WHERE НомерОтдела =" + "'" + Convert.ToInt64(textBox1.Text) + "'", connect);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
            {
                MessageBox.Show("Заполните поле");
            }
            else
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Sotrudniki WHERE Должность =" + "'" + textBox1.Text + "'", connect);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView4.DataSource = ds.Tables[0];
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string phrase = textBox1.Text;
            string[] words = phrase.Split(' ');
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Sotrudniki WHERE Фамилия =" + "'" + words[0] + "'" + "AND Имя =" + "'" + words[1] + "'" + "AND Отчество =" + "'" + words[2] + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.sotrudnikiTableAdapter.Fill(this.kursa4DataSet17.Sotrudniki);
            dataGridView4.DataSource = this.kursa4DataSet17.Sotrudniki;
            this.otdeliTableAdapter.Fill(this.kursa4DataSet16.Otdeli);
            dataGridView2.DataSource = this.kursa4DataSet16.Otdeli;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button24_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView4.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView4.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Sotrudniki WHERE Должность =" + "'" + comboBox1.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Sotrudniki WHERE НомерОтдела =" + "'" + Convert.ToInt64(comboBox3.Text) + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox3.Text.Equals(""))
                {
                    MessageBox.Show("Заполните поле");
                }
                else
                {
                    SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE КодВклада =" + "'" + textBox3.Text + "'", connect);
                    DataSet ds = new DataSet();
                    dataAdapter.Fill(ds);
                    dataGridView3.DataSource = ds.Tables[0];
                }
            }
            catch
            {
                MessageBox.Show("Проверьте правильность данных");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox3.Text.Equals(""))
            {
                MessageBox.Show("Заполните поле");
            }
            else
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE НазваниеВклада =" + "'" + textBox3.Text + "'", connect);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView3.DataSource = ds.Tables[0];
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.typeVkladTableAdapter.Fill(this.kursa4DataSet18.TypeVklad);
            dataGridView3.DataSource = this.kursa4DataSet18.TypeVklad;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView3.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView3.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            int a;
            if (comboBox5.Text == "Пролонгируемый")
            {
                a = 1;
            }
            else
            {
                a = 0;
            }

            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE Пролонгируемый =" + "'" + Convert.ToByte(a) + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE Процент =" + "'" + Convert.ToInt64(comboBox9.Text) + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }
    }
}

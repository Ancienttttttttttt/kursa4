using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace K4
{
    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
        }
        
        SqlConnection connect = new SqlConnection(
"Data Source=DESKTOP-VOAQ9R6;Initial Catalog=Kursa4;Integrated Security=True");
        private void Form7_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet13.Vkladchiki". При необходимости она может быть перемещена или удалена.
            this.vkladchikiTableAdapter.Fill(this.kursa4DataSet13.Vkladchiki);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet12.Vkladi". При необходимости она может быть перемещена или удалена.
            this.vkladiTableAdapter.Fill(this.kursa4DataSet12.Vkladi);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet11.TypeVklad". При необходимости она может быть перемещена или удалена.
            this.typeVkladTableAdapter.Fill(this.kursa4DataSet11.TypeVklad);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet10.Sotrudniki". При необходимости она может быть перемещена или удалена.
            this.sotrudnikiTableAdapter.Fill(this.kursa4DataSet10.Sotrudniki);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet9.Otdeli". При необходимости она может быть перемещена или удалена.
            this.otdeliTableAdapter.Fill(this.kursa4DataSet9.Otdeli);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet8.Banki". При необходимости она может быть перемещена или удалена.
            this.bankiTableAdapter.Fill(this.kursa4DataSet8.Banki);

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

        private void button3_Click(object sender, EventArgs e)
        {
            string phrase = textBox1.Text;
            string[] words = phrase.Split(' ');
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Sotrudniki WHERE Фамилия =" + "'" + words[0] + "'" + "AND Имя =" + "'" + words[1] + "'" + "AND Отчество =" + "'" + words[2] + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
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
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE Процент =" + "'" + comboBox9.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string phrase = textBox4.Text;
            string[] words = phrase.Split(' ');
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladchiki WHERE Фамилия =" + "'" + words[0] + "'" + "AND Имя =" + "'" + words[1] + "'" + "AND Отчество =" + "'" + words[2] + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void button25_Click(object sender, EventArgs e)
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

        private void button23_Click(object sender, EventArgs e)
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

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE НомерСотрудника =" + "'" + Convert.ToInt64(comboBox4) + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladchiki WHERE КодВкладчика =" + "'" + comboBox6.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE КодВклада =" + "'" + comboBox7.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE КодВклада =" + "'" + comboBox7.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
        }
    }
}

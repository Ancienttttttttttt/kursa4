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

namespace Kur4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection connect = new SqlConnection(
"Data Source=MS-SQL\\SQLEXPRESS;Initial Catalog=Kursa4;Integrated Security=True");


        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet14.Sotrudniki". При необходимости она может быть перемещена или удалена.
            this.sotrudnikiTableAdapter2.Fill(this.kursa4DataSet14.Sotrudniki);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet6.Sotrudniki". При необходимости она может быть перемещена или удалена.
            this.sotrudnikiTableAdapter1.Fill(this.kursa4DataSet6.Sotrudniki);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet5.Vkladi". При необходимости она может быть перемещена или удалена.
            this.vkladiTableAdapter.Fill(this.kursa4DataSet5.Vkladi);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet4.Vkladchiki". При необходимости она может быть перемещена или удалена.
            this.vkladchikiTableAdapter.Fill(this.kursa4DataSet4.Vkladchiki);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet3.TypeVklad". При необходимости она может быть перемещена или удалена.
            this.typeVkladTableAdapter.Fill(this.kursa4DataSet3.TypeVklad);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet2.Sotrudniki". При необходимости она может быть перемещена или удалена.
            this.sotrudnikiTableAdapter.Fill(this.kursa4DataSet2.Sotrudniki);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet1.Otdeli". При необходимости она может быть перемещена или удалена.
            this.otdeliTableAdapter.Fill(this.kursa4DataSet1.Otdeli);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursa4DataSet.Banki". При необходимости она может быть перемещена или удалена.
            this.bankiTableAdapter.Fill(this.kursa4DataSet.Banki);

            connect.Open();
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

        

        private void button6_Click(object sender, EventArgs e)
        {

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
            this.bankiTableAdapter.Fill(this.kursa4DataSet.Banki);
            dataGridView1.DataSource = this.kursa4DataSet.Banki;
        }

        private void button14_Click(object sender, EventArgs e)
        {

            Form3 frm = new Form3();
            frm.Show();
        }



        private void button5_Click_1(object sender, EventArgs e)
        {
            this.sotrudnikiTableAdapter.Fill(this.kursa4DataSet2.Sotrudniki);
            dataGridView4.DataSource = this.kursa4DataSet2.Sotrudniki;
            this.otdeliTableAdapter.Fill(this.kursa4DataSet1.Otdeli);
            dataGridView2.DataSource = this.kursa4DataSet1.Otdeli;
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

        private void button15_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string phrase = textBox1.Text;
            string[] words = phrase.Split(' ');
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Sotrudniki WHERE Фамилия =" + "'" + words[0] + "'"+"AND Имя =" + "'" + words[1] +"'"+"AND Отчество =" + "'" + words[2] +"'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.typeVkladTableAdapter.Fill(this.kursa4DataSet3.TypeVklad);
            dataGridView3.DataSource = this.kursa4DataSet3.TypeVklad;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            this.vkladiTableAdapter.Fill(this.kursa4DataSet5.Vkladi);
            this.vkladchikiTableAdapter.Fill(this.kursa4DataSet4.Vkladchiki);
            dataGridView4.DataSource = this.kursa4DataSet4.Vkladchiki;

        }

        private void button16_Click(object sender, EventArgs e)
        {
            otdeliTableAdapter.Update(kursa4DataSet1);
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

        private void button18_Click(object sender, EventArgs e)
        {
            typeVkladTableAdapter.Update(kursa4DataSet3);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            vkladiTableAdapter.Update(kursa4DataSet5);
            vkladchikiTableAdapter.Update(kursa4DataSet4);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            bankiTableAdapter.Update(kursa4DataSet);
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

        private void button6_Click_1(object sender, EventArgs e)
        {
            string phrase = textBox4.Text;
            string[] words = phrase.Split(' ');
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladchiki WHERE Фамилия =" + "'" + words[0] + "'" + "AND Имя =" + "'" + words[1] + "'" + "AND Отчество =" + "'" + words[2] + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView5.DataSource = ds.Tables[0];
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
                dataGridView4.Rows[i].Visible = dataGridView4[1, i].Value.ToString() == textBox1.Text;

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

        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {


        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Vkladi WHERE НомерСотрудника =" + "'" + Convert.ToInt64(comboBox4) + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView6.DataSource = ds.Tables[0];
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From Banki WHERE Председатель =" + "'" + comboBox2.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
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
                a=0;
            }

            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE Пролонгируемый =" + "'" + Convert.ToByte(a) + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
        }

        private void button26_Click(object sender, EventArgs e)
        {
            Form7 fm7 = new Form7();
            fm7.ShowDialog();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            Form6 fm6 = new Form6();
            fm6.ShowDialog();
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

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

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("Select * From TypeVklad WHERE Процент =" + "'" + comboBox9.Text + "'", connect);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }
    }
}

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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        int i = 0;
        int b = 0;
        SqlConnection connect = new SqlConnection(
"Data Source=DESKTOP-VOAQ9R6;Initial Catalog=Kursa4;Integrated Security=True");
        private void button2_Click(object sender, EventArgs e)
        {
            connect.Open();
            try
            {
                SqlCommand command1 = new SqlCommand("", connect);
                string query1 = "INSERT INTO Sotrudniki (Фамилия,Имя,Отчество,Должность,Телефон,НомерОтдела,Логин,Пароль)";
                query1 += " VALUES (@Фамилия, @Имя, @Отчество, @Должность,@Телефон, @НомерОтдела, @Логин, @Пароль)";
                SqlCommand mycommand1 = new SqlCommand(query1, connect);
                mycommand1.Parameters.AddWithValue("@Фамилия", textBox6.Text);
                mycommand1.Parameters.AddWithValue("@Имя", textBox5.Text);
                mycommand1.Parameters.AddWithValue("@Отчество", textBox3.Text);
                mycommand1.Parameters.AddWithValue("@Должность", textBox2.Text);
                mycommand1.Parameters.AddWithValue("@Телефон", textBox1.Text);
                mycommand1.Parameters.AddWithValue("@НомерОтдела", Convert.ToInt64(textBox10.Text));
                mycommand1.Parameters.AddWithValue("@Логин", textBox8.Text);
                mycommand1.Parameters.AddWithValue("@Пароль", textBox9.Text);
                mycommand1.ExecuteNonQuery();
                b++;
                textBox7.Text = "Добавлен " + b + " сотрудник";
                connect.Close();
            }
            catch
            {
                connect.Close();
                MessageBox.Show("Проверьте правильность данных");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            connect.Open();
            try
            {
                SqlCommand command = new SqlCommand("", connect);
                string query = "INSERT INTO Otdeli (НазваниеОтдела)";
                query += " VALUES (@НазваниеОтдела)";
                SqlCommand myCommand = new SqlCommand(query, connect);
                myCommand.Parameters.AddWithValue("@НазваниеОтдела", textBox4.Text);
                myCommand.ExecuteNonQuery();
                i++;
                textBox7.Text = "Добавлен " + i + " отдел";
                connect.Close();
            }
            catch
            {
                connect.Close();
                MessageBox.Show("Проверьте правильность данных");
            }
        }
    }
}

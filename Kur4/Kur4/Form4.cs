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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        SqlConnection connect = new SqlConnection("Data Source=MS-SQL\\SQLEXPRESS;Initial Catalog=Kursa4;Integrated Security=True");

        private void button2_Click(object sender, EventArgs e)
        {
            connect.Open();
            bool success=false;

            
                try
                {
                const string comand = "Select * From Sotrudniki WHERE Логин=@Логин AND Пароль=@Пароль AND Должность=@Должность";
                SqlCommand check = new SqlCommand(comand, connect);
                check.Parameters.AddWithValue("@Логин", textBox1.Text);
                check.Parameters.AddWithValue("@Пароль", textBox2.Text);
                check.Parameters.AddWithValue("@Должность", comboBox1.Text);
                using (var dataReader = check.ExecuteReader())
                    {
                        success = dataReader.Read();
                    }
                }
                finally
                {
                    connect.Close();
                }

                if (success)
                {
                if(comboBox1.Text == "Администратор")
                {
                    Form1 form1 = new Form1();
                    form1.Show();
                }
                else
                {
                    Form8 form8 = new Form8();
                    form8.Show();
                }
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль или должность");
                }
            
      
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

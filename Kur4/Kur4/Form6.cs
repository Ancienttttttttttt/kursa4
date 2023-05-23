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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Kur4
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }
        int i = 0;
        SqlConnection connect = new SqlConnection(
"Data Source=MS-SQL\\SQLEXPRESS;Initial Catalog=Kursa4;Integrated Security=True");

        private void button2_Click(object sender, EventArgs e)
        {
            connect.Open();
            try
            {
                SqlCommand command1 = new SqlCommand("", connect);
                string query1 = "INSERT INTO Vkladi (КодВкладчика, КодВклада, ДатаОткрытияВклада, ДатаЗакрытияВклада, НомерСотрудника)";
                query1 += " VALUES (@КодВкладчика, @КодВклада, @ДатаОткрытияВклада, @ДатаЗакрытияВклада, @НомерСотрудника)";
                SqlCommand mycommand1 = new SqlCommand(query1, connect);
                mycommand1.Parameters.AddWithValue("@ДатаЗакрытияВклада", textBox4.Text);
                mycommand1.Parameters.AddWithValue("@КодВкладчика", textBox5.Text);
                mycommand1.Parameters.AddWithValue("@КодВклада", textBox3.Text);
                mycommand1.Parameters.AddWithValue("@ДатаОткрытияВклада", textBox1.Text);
                mycommand1.Parameters.AddWithValue("@НомерСотрудника", textBox8.Text);
                i++;
                MessageBox.Show("Добавлен " + i + " вклад");
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


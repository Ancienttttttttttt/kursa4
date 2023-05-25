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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        int i = 0;
        SqlConnection connect = new SqlConnection(
"Data Source=DESKTOP-VOAQ9R6;Initial Catalog=Kursa4;Integrated Security=True");

        private void button2_Click(object sender, EventArgs e)
        {
            connect.Open();
            try
            {
                SqlCommand command1 = new SqlCommand("", connect);
                string query1 = "INSERT INTO Vkladi (КодВкладчика, КодВклада, ДатаОткрытияВклада, ДатаЗакрытияВклада, НомерСотрудника)";
                query1 += " VALUES (@КодВкладчика, @КодВклада, @ДатаОткрытияВклада, @ДатаЗакрытияВклада, @НомерСотрудника)";
                SqlCommand mycommand1 = new SqlCommand(query1, connect);
                mycommand1.Parameters.AddWithValue("@КодВкладчика", Convert.ToInt64(textBox5.Text));
                mycommand1.Parameters.AddWithValue("@КодВклада", Convert.ToInt64(textBox3.Text));
                mycommand1.Parameters.AddWithValue("@ДатаОткрытияВклада", Convert.ToDateTime(textBox1.Text));
                mycommand1.Parameters.AddWithValue("@ДатаЗакрытияВклада", Convert.ToDateTime(textBox4.Text));
                mycommand1.Parameters.AddWithValue("@НомерСотрудника", Convert.ToInt64(textBox8.Text));
                mycommand1.ExecuteNonQuery();
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

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
    public partial class Form8 : Form
    {
        public Form8()
        {
            InitializeComponent();
        }
        int i = 0;
        SqlConnection connect = new SqlConnection(
"Data Source=DESKTOP-VOAQ9R6;Initial Catalog=Kursa4;Integrated Security=True");
        private void button2_Click(object sender, EventArgs e)
        {
            connect.Open();

            if(checkBox1.Checked == true)
            {
                checkBox1.Enabled = true;
            }
            else
            {
                checkBox1.Enabled = false;
            }
            if (checkBox2.Checked == true)
            {
                checkBox2.Enabled = true;
            }
            else
            {
                checkBox2.Enabled = false;
            }
            string query1 = "INSERT INTO TypeVklad (НазваниеВклада, СрокВклада, МинимальнаяСумма, Процент, Пролонгируемый, Пополнение)";
            query1 += " VALUES (@НазваниеВклада, @СрокВклада, @МинимальнаяСумма, @Процент, @Пролонгируемый, @Пополнение)";
            SqlCommand mycommand1 = new SqlCommand(query1, connect);
            mycommand1.Parameters.AddWithValue("@НазваниеВклада", textBox3.Text);
            mycommand1.Parameters.AddWithValue("@СрокВклада", textBox1.Text);
            mycommand1.Parameters.AddWithValue("@МинимальнаяСумма", textBox8.Text);
            mycommand1.Parameters.AddWithValue("@Процент", textBox4.Text);
            mycommand1.Parameters.AddWithValue("@Пролонгируемый", Convert.ToBoolean(checkBox1.Checked));
            mycommand1.Parameters.AddWithValue("@Пополнение", Convert.ToBoolean(checkBox2.Checked));
            mycommand1.ExecuteNonQuery();
            i++;
            MessageBox.Show("Добавлен " + i + " тип вклада");
            connect.Close();
        }
    }
}

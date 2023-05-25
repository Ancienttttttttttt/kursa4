using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace K4
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        int i = 0;
        SqlConnection connect = new SqlConnection(
"Data Source=DESKTOP-VOAQ9R6;Initial Catalog=Kursa4;Integrated Security=True");
        private void Form2_Load(object sender, EventArgs e)
        {
            connect.Open();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "")
            {

                connect.Open();
                SqlCommand command = new SqlCommand("", connect);
                string query = "INSERT INTO Banki (Наименование,ТелефонБанка,АдресБанка,Председатель)";
                query += " VALUES (@Наименование,@ТелефонБанка,@АдресБанка,@Председатель)";
                SqlCommand myCommand = new SqlCommand(query, connect);
                myCommand.Parameters.AddWithValue("@Наименование", textBox1.Text);
                myCommand.Parameters.AddWithValue("@ТелефонБанка", textBox2.Text);
                myCommand.Parameters.AddWithValue("@АдресБанка", textBox3.Text);
                myCommand.Parameters.AddWithValue("@Председатель", textBox4.Text);
                myCommand.ExecuteNonQuery();
                i++;
                MessageBox.Show("Добавлен " + i + " банк");
                connect.Close();
            }
            else
            {
                MessageBox.Show("Введите данные!");
            }
        }
    }
}

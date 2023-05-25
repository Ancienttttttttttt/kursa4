﻿using System;
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
    public partial class Form4 : Form
    {
        public Form4()
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
                string query1 = "INSERT INTO Vkladchiki (Фамилия, Имя, Отчество, ДатаРождения, Адрес, Телефон, КодБанка)";
                query1 += " VALUES (@Фамилия, @Имя, @Отчество, @ДатаРождения, @Адрес, @Телефон, @КодБанка)";
                SqlCommand mycommand1 = new SqlCommand(query1, connect);
                mycommand1.Parameters.AddWithValue("@Фамилия", textBox5.Text);
                mycommand1.Parameters.AddWithValue("@Имя", textBox3.Text);
                mycommand1.Parameters.AddWithValue("@Отчество", textBox2.Text);
                mycommand1.Parameters.AddWithValue("@Адрес", textBox4.Text);
                mycommand1.Parameters.AddWithValue("@ДатаРождения", Convert.ToDateTime(textBox1.Text));
                mycommand1.Parameters.AddWithValue("@Телефон", textBox8.Text);
                mycommand1.Parameters.AddWithValue("@КодБанка", Convert.ToInt64(textBox7.Text));
                mycommand1.ExecuteNonQuery();
                i++;
                MessageBox.Show("Добавлен " + i + " вкладчик");
                connect.Close();
            }
            catch
            {
                connect.Close();
                MessageBox.Show("Проверьте правильность данных");
            }
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            textBox1.Text = monthCalendar1.SelectionRange.Start.ToShortDateString();
        }
    }
}

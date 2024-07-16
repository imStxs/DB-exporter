using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MySql.Data;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;
using System.Threading;


namespace Экспорт
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string conStr = "server=127.0.0.1;user=root;" + "database=123;password=root;sslMode=none;";

        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\\Users\maxme\OneDrive\Рабочий стол\БД\Таблица за ";
            DirectoryInfo di = Directory.CreateDirectory(path);
            Directory.Move(@"C:\\Users\maxme\OneDrive\Рабочий стол\БД\Таблица за ", @"C:\\Users\maxme\OneDrive\Рабочий стол\БД\Таблица за " + di.CreationTime.ToShortDateString());
        }
     
        private void button2_Click_1(object sender, EventArgs e)
        {
            int selectedIndex = comboBox1.SelectedIndex;
            string selecteVal = (string)comboBox1.SelectedItem;
            string conStr = "server=127.0.0.1;user=root;" + "database=123;password=root;sslMode=none;";
            using (MySqlConnection con = new MySqlConnection(conStr))
            {
                string script = "SELECT * FROM " + selecteVal + ";";

                MySqlConnection connection = new MySqlConnection(conStr);
                connection.Open();

                MySqlDataAdapter mySql_dataAdapter = new MySqlDataAdapter(script, connection);
                System.Data.DataTable table = new System.Data.DataTable();

                mySql_dataAdapter.Fill(table);
                dataGridView1.DataSource = table;
                connection.Close();
            }
        }
       
        private void button4_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy-MM-dd";
            string date = dateTimePicker1.Text; 
            /*dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "yyyy-MM-dd";
            string date2 = dateTimePicker2.Text;*/

            using (MySqlConnection con = new MySqlConnection(conStr))
            {
                string script = "SELECT * FROM "+ comboBox1.SelectedItem+ " WHERE( birthday LIKE '" + date +"')";

                MySqlConnection connection = new MySqlConnection(conStr);
                connection.Open();

                MySqlDataAdapter mySql_dataAdapter = new MySqlDataAdapter(script, connection);
                System.Data.DataTable table = new System.Data.DataTable();

                mySql_dataAdapter.Fill(table);
                dataGridView1.DataSource = table;
                connection.Close();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            string conStr = "server=127.0.0.1;user=root;" + "database=123;password=root;sslMode=none;";
            using (MySqlConnection con = new MySqlConnection(conStr))
            {
                string script = "SELECT * FROM user;";

                MySqlConnection connection = new MySqlConnection(conStr);
                connection.Open();

                MySqlDataAdapter mySql_dataAdapter = new MySqlDataAdapter(script, connection);
                System.Data.DataTable table = new System.Data.DataTable();

                mySql_dataAdapter.Fill(table);
                dataGridView1.DataSource = table;
                
                
                    string str = " SELECT table_name FROM information_schema.tables WHERE table_schema = '123' ";
                    MySqlCommand command = new MySqlCommand(str, connection);
                    MySqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        comboBox1.Items.Add(reader.GetString("TABLE_NAME"));
                    }
                connection.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }



            //string path = @"C:\\Users\maxme\OneDrive\Рабочий стол\БД\Таблица за ";
            //DirectoryInfo di = Directory.CreateDirectory(path);
            //Directory.Move(@"C:\\Users\maxme\OneDrive\Рабочий стол\БД\Таблица за ", @"C:\\Users\maxme\OneDrive\Рабочий стол\БД\Таблица за " + di.CreationTime.ToShortDateString());
            //string a = di.CreationTime.ToShortDateString();
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
            //ExcelWorkBook.SaveAs(@"C:\\Users\\maxme\\OneDrive\\Рабочий стол\\БД\\Таблица за " + a + "\\табличка.xlsx"); //формат Excel 2007
            //ExcelWorkBook.Close(false); //false - закрыть рабочую книгу не сохраняя изменения
            //ExcelApp.Quit(); //закрываем приложение Excel
            //MessageBox.Show("Файл сохранён!", "Сохранение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}

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
using System.Xml;
using System.Xml.Linq;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Excel;

namespace internship
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.Text = "internship";
            string[] Stroka = { "Имя Объекта", "Время старта", "Время остановки", "Схема проверки", "Интервал измерения"};
            //добавляем ее в таблицу
            addGridParam(Stroka, dataGridView1);

            //указываем путь
            DirectoryInfo di = new DirectoryInfo("..\\..\\..\\#data/card/PKE");

            //заходим в папку 
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                foreach (DirectoryInfo dir1 in dir.GetDirectories())
                {
                    foreach (var file in dir1.GetFiles())
                    {
                        IEnumerable<Checks> items = LoadFiles(file.FullName);
                        foreach (Checks item in items)
                        {
                            //выводим данные о всех объектах
                            string[] strokaEshe = { item.nameObject, item.transfTime(item.TimeStart), item.transfTime(item.TimeStop), item.active_cxema, item.transf_averaging_interval_time(item.averaging_interval_time) };
                            addGridParam(strokaEshe, dataGridView1);
                        }
                    }
                }
            }

        }

        //Загрузка и обработка файлов
        static IEnumerable<Checks> LoadFiles(string path)
        {
            XDocument xdoc = XDocument.Load(path);
            var items = from xe in xdoc.Element("RM3_ПКЭ").Elements("Param_Check_PKE")
                        select new Checks
                        {
                            TimeStart = xe.Attribute("TimeStart").Value,
                            TimeStop = xe.Attribute("TimeStop").Value,
                            nameObject = xe.Attribute("nameObject").Value,
                            averaging_interval_time = xe.Attribute("averaging_interval_time").Value,
                            averaging_interval = xe.Attribute("averaging_interval").Value,
                            active_cxema = xe.Attribute("active_cxema").Value,
                        };
            return items;
        }

        //функция заполнения таблицы
        /*
             N - массив строк
             Grid - сетка в которой будем отображать данные
        */
        public void addGridParam(string[] N, DataGridView Grid)
        {
            //пока столбцов не будет достаточное количество добавляем их
            while (N.Length > Grid.ColumnCount)
            {
                //если колонок нехватает добавляем их пока их будет хватать
                Grid.Columns.Add("", "");
            }
            //заполняем строку
            Grid.Rows.Add(N);
        }

        class Checks
        {
            public string TimeStart { get; set; }
            public string TimeStop { get; set; }
            public string nameObject { get; set; }
            public string averaging_interval_time { get; set; }
            public string averaging_interval { get; set; }
            public string active_cxema { get; set; }
            //приведение даты в нормальный вид
            public string transfTime(string date)
            {
                DateTime date1 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                DateTime transfData = date1.AddMilliseconds(Convert.ToUInt64(date));                
                string date_str = transfData.ToString("dd.MM.yyyy HH.mm");
                return date_str;
            }
            //приведение интервал между измерениями в нормальный вид
            public string transf_averaging_interval_time(string date)
            {
                TimeSpan TS = new TimeSpan(0,0, 0, 0,Convert.ToInt32(date));
                return TS.ToString();   
            }
        }
        //переход к форме 1 по двойному нажатию на клавишу
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string nameObject = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            string TimeStart = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            Form1 f1 = new Form1(nameObject, TimeStart);
            f1.Show();
        }

        private void button1_Click(object sender, EventArgs e)
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
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем excel.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
    }
}

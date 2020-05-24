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


namespace internship
{
    public partial class Form1 : Form
    {
        string nameObject, TimeStart;
        int i = 0, j = 0, k = 0;
        public Form1(string nameObject, string TimeStart)
        {
            InitializeComponent();
            this.nameObject = nameObject;
            this.TimeStart = TimeStart;            
        }        

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "internship";
            DirectoryInfo di = new DirectoryInfo("..\\..\\..\\#data/card/PKE/" + nameObject + "/" + TimeStart);
            foreach (var file in di.GetFiles())
            {
                IEnumerable<Check1> items1 = LoadFiles1(file.FullName);
                IEnumerable<Check2> items2 = LoadFiles2(file.FullName);
                IEnumerable<Check3> items3 = LoadFiles3(file.FullName);

                //сортировки
                List<Check1> list1 = items1.ToList();
                list1 = list1.OrderBy(b => b.TimeTek).ToList();
                List<Check2> list2 = items2.ToList();
                list2 = list2.OrderBy(b => b.TimeTek).ToList();
                List<Check3> list3 = items3.ToList();
                list3 = list3.OrderBy(b => b.TimeTek).ToList();

                foreach (Check1 itemo in list1)
                {
                    if (i==0)
                    {
                        string[] Stroka = { "Дата / время", "UA", "IA", "PA", "QA", "SA", "Freq", "sigmaUy" };
                        //добавляем ее в таблицу
                        addGridParam(Stroka, dataGridView1);
                        i++;
                    }
                    string[] strokaEshe = { TimeStart, itemo.UA, itemo.IA, itemo.PA, itemo.QA, itemo.SA, itemo.Freq, itemo.sigmaUy };
                    addGridParam(strokaEshe, dataGridView1);
                }

                foreach (Check2 item in list2)
                {
                    if (j == 0)
                    {
                        string[] Stroka = { "Дата / время", "UAB", "UBC", "UCA", "IAB", "IBC", "ICA", "IA", "IB", "IC", "PO", "PP", "QO", "QP", "SO", "SP", "UO", "UP", "IO", "IP", "KO", "Freq", "sigmaUy", "sigmaUyAB", "sigmaUyBC", "sigmaUyCA" };
                        //добавляем ее в таблицу
                        addGridParam(Stroka, dataGridView1);
                        j++;
                    }
                    string[] strokaEshe = { TimeStart , item.UAB, item.UBC, item.UCA, item.IAB, item.IBC, item.ICA, item.IA, item.IB, item.IC, item.PO, item.PP, item.QO, item.QP, item.SO, item.SP, item.UO, item.UP, item.IO, item.IP, item.KO, item.Freq, item.sigmaUy, item.sigmaUyAB, item.sigmaUyBC, item.sigmaUyCA };
                    addGridParam(strokaEshe, dataGridView1);
                }

                foreach (Check3 itemt in list3)
                {
                    if (k == 0)
                    {
                        string[] Stroka = { "Дата / время", "UAB", "UBC", "UCA", "IA", "IB", "IC", "UA", "UB", "UC", "PO", "PP", "PH", "QO", "QP", "QH", "SO", "SP", "SH", "UO", "UP", "UH", "IO", "IP", "IH", "KO", "KH", "Freq", "sigmaUy", "sigmaUyA", "sigmaUyB", "sigmaUyC", };
                        //добавляем ее в таблицу
                        addGridParam(Stroka, dataGridView1);
                        k++;
                    }
                    string[] strokaEshe = { TimeStart, itemt.UAB, itemt.UBC, itemt.UCA, itemt.IA, itemt.IB, itemt.IC, itemt.UA, itemt.UB, itemt.UC, itemt.PO, itemt.PP, itemt.PH, itemt.QO, itemt.QP, itemt.QH, itemt.SO, itemt.SP, itemt.SH, itemt.UO, itemt.UP, itemt.UH, itemt.IO, itemt.IP, itemt.IH, itemt.KO, itemt.KH, itemt.Freq, itemt.sigmaUy, itemt.sigmaUyA, itemt.sigmaUyB, itemt.sigmaUyC };
                    addGridParam(strokaEshe, dataGridView1);
                }
            }
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

        //загрузка и обработка файлов для проверки 1
        static IEnumerable<Check1> LoadFiles1(string path)
        {
            XDocument xdoc = XDocument.Load(path);
            var items = from xe in xdoc.Element("RM3_ПКЭ").Elements("Result_Check_PKE")
                        where xe.Attribute("pke_cxema").Value == "1"
                        select new Check1
                        {
                            pke_cxema = xe.Attribute("pke_cxema").Value,
                            TimeTek = xe.Attribute("TimeTek").Value,
                            IA = xe.Attribute("IA").Value,
                            UA = xe.Attribute("UA").Value,
                            PA = xe.Attribute("PA").Value,
                            QA = xe.Attribute("QA").Value,
                            SA = xe.Attribute("SA").Value,
                            Freq = xe.Attribute("Freq").Value,
                            sigmaUy = xe.Attribute("sigmaUy").Value
                        };
            return items;
        }
        class Check1
        {
            public string pke_cxema { get; set; }
            public string TimeTek { get; set; }
            public string Freq { get; set; }
            public string sigmaUy { get; set; }
            public string UA { get; set; }
            public string PA { get; set; }
            public string QA { get; set; }
            public string SA { get; set; }
            public string UB { get; set; }
            public string IA { get; set; }
        }

        //загрузка и обработка файлов для проверки 2 
        static IEnumerable<Check2> LoadFiles2(string path)
        {
            XDocument xdoc = XDocument.Load(path);
            var items = from xe in xdoc.Element("RM3_ПКЭ").Elements("Result_Check_PKE")
                        where xe.Attribute("pke_cxema").Value == "2"
                        select new Check2
                        {
                            pke_cxema = xe.Attribute("pke_cxema").Value,
                            UAB = xe.Attribute("UAB").Value,
                            UBC = xe.Attribute("UBC").Value,
                            UCA = xe.Attribute("UCA").Value,
                            IAB = xe.Attribute("IAB").Value,
                            IBC = xe.Attribute("IBC").Value,
                            ICA = xe.Attribute("ICA").Value,
                            IA = xe.Attribute("IA").Value,
                            IB = xe.Attribute("IB").Value,
                            IC = xe.Attribute("IC").Value,
                            PO = xe.Attribute("PO").Value,
                            PP = xe.Attribute("PP").Value,
                            QO = xe.Attribute("QO").Value,
                            QP = xe.Attribute("QP").Value,
                            SO = xe.Attribute("SO").Value,
                            SP = xe.Attribute("SP").Value,
                            UO = xe.Attribute("UO").Value,
                            UP = xe.Attribute("UP").Value,
                            IO = xe.Attribute("IO").Value,
                            IP = xe.Attribute("IP").Value,
                            KO = xe.Attribute("KO").Value,
                            Freq = xe.Attribute("Freq").Value,
                            sigmaUy = xe.Attribute("sigmaUy").Value,
                            sigmaUyAB = xe.Attribute("sigmaUyAB").Value,
                            sigmaUyBC = xe.Attribute("sigmaUyBC").Value,
                            sigmaUyCA = xe.Attribute("sigmaUyCA").Value,
                            TimeTek = xe.Attribute("TimeTek").Value,
                        };
            return items;
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

        class Check2
        {
            public string pke_cxema { get; set; }
            public string UAB { get; set; }         
            public string UBC { get; set; }
            public string UCA { get; set; }
            public string IAB { get; set; }
            public string IBC { get; set; }
            public string ICA { get; set; }
            public string IA { get; set; }
            public string IB { get; set; }
            public string IC { get; set; }
            public string PO { get; set; }
            public string PP { get; set; }
            public string QO { get; set; }
            public string QP { get; set; }
            public string SO { get; set; }
            public string SP { get; set; }
            public string UO { get; set; }
            public string UP { get; set; }
            public string IO { get; set; }
            public string IP { get; set; }
            public string KO { get; set; }
            public string Freq { get; set; }
            public string sigmaUy { get; set; }
            public string sigmaUyAB { get; set; }
            public string sigmaUyBC { get; set; }
            public string sigmaUyCA { get; set; }
            public string TimeTek { get; set; }
        }

        //загрузка и обработка файлов для проверки 3
        static IEnumerable<Check3> LoadFiles3(string path)
        {
            XDocument xdoc = XDocument.Load(path);
            var items = from xe in xdoc.Element("RM3_ПКЭ").Elements("Result_Check_PKE")
                        where xe.Attribute("pke_cxema").Value == "3"
                        select new Check3
                        {
                            pke_cxema = xe.Attribute("pke_cxema").Value,
                            UAB = xe.Attribute("UAB").Value,
                            UBC = xe.Attribute("UBC").Value,
                            UCA = xe.Attribute("UCA").Value,                         
                            IA = xe.Attribute("IA").Value,
                            IB = xe.Attribute("IB").Value,
                            IC = xe.Attribute("IC").Value,
                            PO = xe.Attribute("PO").Value,
                            PP = xe.Attribute("PP").Value,
                            QO = xe.Attribute("QO").Value,
                            QP = xe.Attribute("QP").Value,
                            SO = xe.Attribute("SO").Value,
                            SP = xe.Attribute("SP").Value,
                            UO = xe.Attribute("UO").Value,
                            UP = xe.Attribute("UP").Value,
                            IO = xe.Attribute("IO").Value,
                            IP = xe.Attribute("IP").Value,
                            KO = xe.Attribute("KO").Value,
                            Freq = xe.Attribute("Freq").Value,
                            sigmaUy = xe.Attribute("sigmaUy").Value,                            
                            TimeTek = xe.Attribute("TimeTek").Value,
                            UA = xe.Attribute("UA").Value,
                            UB = xe.Attribute("UB").Value,
                            UC = xe.Attribute("UC").Value,
                            PH = xe.Attribute("PH").Value,
                            QH = xe.Attribute("QH").Value,
                            SH = xe.Attribute("SH").Value,
                            UH = xe.Attribute("UH").Value,
                            IH = xe.Attribute("IH").Value,
                            KH = xe.Attribute("KH").Value,
                            sigmaUyA = xe.Attribute("sigmaUyA").Value,
                            sigmaUyB = xe.Attribute("sigmaUyB").Value,
                            sigmaUyC = xe.Attribute("sigmaUyC").Value,
                        };
            return items;
        }

        class Check3
        {
            public string pke_cxema { get; set; }
            public string UAB { get; set; }
            public string UBC { get; set; }
            public string UCA { get; set; }           
            public string IA { get; set; }
            public string IB { get; set; }
            public string IC { get; set; }
            public string PO { get; set; }
            public string PP { get; set; }
            public string QO { get; set; }
            public string QP { get; set; }
            public string SO { get; set; }
            public string SP { get; set; }
            public string UO { get; set; }
            public string UP { get; set; }
            public string IO { get; set; }
            public string IP { get; set; }
            public string KO { get; set; }
            public string Freq { get; set; }
            public string sigmaUy { get; set; }           
            public string TimeTek { get; set; }
            public string UA { get; set; }           
            public string UB { get; set; }
            public string PH { get; set; }
            public string QH { get; set; }
            public string SH { get; set; }
            public string UC { get; set; }
            public string UH { get; set; }
            public string IH { get; set; }
            public string KH { get; set; }
            public string sigmaUyA { get; set; }
            public string sigmaUyB { get; set; }
            public string sigmaUyC { get; set; }
        }
    }
}

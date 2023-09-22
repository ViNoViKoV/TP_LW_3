using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ExcelDataReader;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace TP_LW_3_c1_n72
{
    public partial class ExchangeRateForm : Form
    {
        //Объявление имени файла
        private string filename = string.Empty;
        //здесь объвление места под файл ексель
        private DataTableCollection dataTable = null;
        public ExchangeRateForm()
        {
            InitializeComponent();
        }
        //Метод проверки на выбор файла excel
        private bool IsExcelFile(string filePath)
        {
            string extension = Path.GetExtension(filePath);
            return extension.Equals(".xls") || extension.Equals(".xlsx") || extension.Equals(".xlsm");
        }

        // Метод для чтения данных из файла Excel
        private DataTable OpenExcelFile(string path)
        {
            // Открытие потока для чтения выбранного файла Excel
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            // Создание объекта для чтения данных из файла Excel
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            // Чтение данных из файла Excel и преобразование их в объект DataSet
            DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = false
                }
            });


            DataTable table = dataSet.Tables[0];
            dataGridView1.DataSource = table;
            return table;
        }

        private void DrawCharts(DataTable table)
        {
            // Добавление данных
            for (int i = 0; i < table.Rows.Count; i++)
            {
                // Данные находятся во втором и третьем столбцах таблицы
                double x = i; // Простое приращение для координаты X для равномерности графика
                double y1 = Convert.ToDouble(table.Rows[i][1]);
                double y2 = Convert.ToDouble(table.Rows[i][2]);
                //графики получают данные
                chart1.Series[0].Points.AddXY(x, y1);
                chart1.Series[1].Points.AddXY(x, y2);

                // Подпись дат из графика
                chart1.Series[0].Points[i].AxisLabel = (table.Rows[i][0]).ToString().Substring(0, 10);
            }
        }

        private void FindLiftDrop(DataTable table)
        {
            double maxLiftD = 0;
            double maxDropD = 0;
            string maxLiftDateD = "-";
            string maxDropDateD = "-";
            double maxLiftY = 0;
            double maxDropY = 0;
            string maxLiftDateY = "-";
            string maxDropDateY = "-";
            //Ищем для доллара
            for (int i = 1; i < table.Rows.Count; i++)
            {
                double currDev = Convert.ToDouble(table.Rows[i][1]) - Convert.ToDouble(table.Rows[i-1][1]);
                if (currDev < 0 && currDev < maxDropD)
                {
                    maxDropD = currDev;
                    maxDropDateD = Convert.ToString(table.Rows[i][0]);

                }
                if (currDev >= 0 && currDev > maxLiftD)
                {
                    maxLiftD = currDev;
                    maxLiftDateD = Convert.ToString(table.Rows[i][0]);
                }
            }
            //Юань
            for (int i = 1; i < table.Rows.Count; i++)
            {
                double currDev = Convert.ToDouble(table.Rows[i][2]) - Convert.ToDouble(table.Rows[i-1][2]);
                if (currDev < 0 && currDev < maxDropY)
                {
                    maxDropY = currDev;
                    maxDropDateY = Convert.ToString(table.Rows[i][0]);

                }
                if (currDev >= 0 && currDev > maxLiftY)
                {
                    maxLiftY = currDev;
                    maxLiftDateY = Convert.ToString(table.Rows[i][0]);
                }
            }
            //вывод в форму
            textBox1.Text = Math.Round(maxLiftD, 2).ToString() + " руб.";
            textBox2.Text = maxLiftDateD.Substring(0, 10); ;
            textBox3.Text = Math.Round(maxLiftY, 2).ToString() + " руб.";
            textBox4.Text = maxLiftDateY.Substring(0, 10); ;
            textBox5.Text = Math.Round(maxDropD, 2).ToString() + " руб.";
            textBox6.Text = maxDropDateD.Substring(0,10);
            textBox7.Text = Math.Round(maxDropY, 2).ToString() + " руб.";
            textBox8.Text = maxDropDateY.Substring(0, 10); ;
        }

        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Выберите файл excel";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                if (IsExcelFile(filePath))
                {
                    DataTable table = OpenExcelFile(filePath);
                    DrawCharts(table);
                    FindLiftDrop(table);
                }
                else
                {
                    MessageBox.Show("Выбран неверный файл.");
                }
            }
        }
    }
}

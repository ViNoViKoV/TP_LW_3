using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using ExcelDataReader;

namespace TP_LW_3_c1_n72
{
    public partial class TemperatureForm : Form
    {
        //Объявление имени файла
        private string filename = string.Empty;
        //здесь объвление места под файл ексель
        private DataTableCollection dataTable = null;
        public TemperatureForm()
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
                double y3 = Convert.ToDouble(table.Rows[i][3]);
                //графики получают данные
                chart1.Series[0].Points.AddXY(x, y1);
                chart1.Series[1].Points.AddXY(x, y2);
                chart1.Series[2].Points.AddXY(x, y3);
                // Подпись дат из графика
                chart1.Series[0].Points[i].AxisLabel = (table.Rows[i][0]).ToString().Substring(0, 10);
            }
        }

        private void FindLiftDrop(DataTable table)
        {
            double maxDeviation = 0;
            string Date = "";
            //Максимальный перепад средней температуры
            for (int i = 1; i < table.Rows.Count; i++)
            {
                double currDev = Convert.ToDouble(table.Rows[i][2]) - Convert.ToDouble(table.Rows[i - 1][2]);
                if (currDev > Math.Abs(maxDeviation))
                {
                    maxDeviation = currDev;
                    Date = Convert.ToString(table.Rows[i][0]);
                }
            }
            //вывод в форму
            textBox1.Text = Math.Round(maxDeviation, 0).ToString();
            textBox2.Text = Date.Substring(0, 10); ;
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
                    //PlotGraph(table);
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

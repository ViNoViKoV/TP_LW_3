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
                }
                else
                {
                    MessageBox.Show("Выбран неверный файл.");
                }
            }
        }
    }
}

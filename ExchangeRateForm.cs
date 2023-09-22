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
            Chart chart = new Chart();
            //явное преобразование типа
            chart.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
        }

    }
    }
}

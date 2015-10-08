using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using OfficeConvert;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void tryConvert(Converter converter, String inputFile, String outputFile)
        {
            try
            {
                converter.Convert(inputFile, outputFile);
            }
            catch (ConvertException err)
            {
                MessageBox.Show(err.Message + "\n" + err.StackTrace);
            }
        }

        private void btnWordClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String inputFile = op.FileName;
            String outputFile = String.Concat(inputFile, ".pdf");
            Converter converter = new WordConverter();
            tryConvert(converter, inputFile, outputFile);
        }

        private void btnExcelClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String inputFile = op.FileName;
            String outputFile = String.Concat(inputFile, ".pdf");
            Converter converter = new ExcelConverter();
            tryConvert(converter, inputFile, outputFile);
        }

        private void btnPptClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String inputFile = op.FileName;
            String outputFile = String.Concat(inputFile, ".pdf");
            Converter converter = new PowerPointConverter();
            tryConvert(converter, inputFile, outputFile);
        }
    }
}

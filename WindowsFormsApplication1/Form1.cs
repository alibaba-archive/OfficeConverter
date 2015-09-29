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

        private void btnWordClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String inputFile = op.FileName;
            String outputFile = String.Concat(inputFile, ".pdf");
            Converter converter = new WordConverter();
            converter.Convert(inputFile, outputFile);
        }

        private void btnExcelClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String inputFile = op.FileName;
            String outputFile = String.Concat(inputFile, ".pdf");
            Converter converter = new ExcelConverter();
            converter.Convert(inputFile, outputFile);
        }

        private void btnPptClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String inputFile = op.FileName;
            String outputFile = String.Concat(inputFile, ".pdf");
            Converter converter = new PowerPointConverter();
            converter.Convert(inputFile, outputFile);
        }
    }
}

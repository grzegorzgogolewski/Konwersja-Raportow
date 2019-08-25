using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OfficeOpenXml;

namespace KonwersjaRaportow
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            Text = @"Konwersja raportów";
            buttonStart.Text = @"Start";
            buttonSelect.Text = "Wybierz plik";
        }

        private void ButtonSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = false,
                Filter = "Pliki CSV (*.csv)|*.csv|Wszystie pliki (*.*)|*.*"
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBoxFileName.Text = dlg.FileName;
            }
        }

        private void ButtonStart_Click(object sender, EventArgs e)
        {
            ExcelTextFormat format = new ExcelTextFormat
            {
                Delimiter = '\t',
                Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                {
                    DateTimeFormat = {ShortDatePattern = "yyyy-mm-dd"}
                },
                Encoding = new UTF8Encoding()
            };

            //read the CSV file from disk
            FileInfo file = new FileInfo(textBoxFileName.Text);

            //create a new Excel package
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create a WorkSheet
                ExcelWorksheet worksheetAll = excelPackage.Workbook.Worksheets.Add("wszystkie");

                //load the CSV data into cell A1
                worksheetAll.Cells["A1"].LoadFromText(file, format);

                int columnsCount = worksheetAll.Dimension.Columns;

                List<string> numeryKOntroli = new List<string>();

                int rowsCounter = 1;

                do
                {
                    numeryKOntroli.Add(worksheetAll.Cells["A" + rowsCounter].Value.ToString());
                    rowsCounter++;

                } while (worksheetAll.Cells["A" + rowsCounter].Value != null);

                numeryKOntroli = numeryKOntroli.Distinct().ToList(); // unikalna lista kontroli

                foreach (string numerKontroli in numeryKOntroli)
                {
                    ExcelWorksheet arkusz = excelPackage.Workbook.Worksheets.Add(numerKontroli);

                    int destRow = 1;
                    rowsCounter = 1;
                    do
                    {
                        if (worksheetAll.Cells[rowsCounter, 1].Value.ToString() == numerKontroli)
                        {
                            worksheetAll.Cells[rowsCounter, 1, rowsCounter, columnsCount].Copy(arkusz.Cells[destRow++, 1]);
                        }

                        rowsCounter++;

                    } while (worksheetAll.Cells["A" + rowsCounter].Value != null);

                    arkusz.Cells.AutoFitColumns();
                }


                excelPackage.SaveAs(new FileInfo(Path.Combine(Path.GetDirectoryName(textBoxFileName.Text) ?? throw new InvalidOperationException(), Path.GetFileNameWithoutExtension(textBoxFileName.Text) + ".xlsx")));
            }

            Process.Start(Path.Combine(Path.GetDirectoryName(textBoxFileName.Text) ?? throw new InvalidOperationException(), Path.GetFileNameWithoutExtension(textBoxFileName.Text) + ".xlsx"));

            MessageBox.Show("Koniec", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}

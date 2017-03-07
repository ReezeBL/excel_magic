using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Windows.Forms;
using Meister_Hämmerlein.Core;
using Microsoft.Office.Interop.Excel;

namespace Meister_Hämmerlein.WindowForms
{
    public partial class MainWindow : Form
    {
        private readonly ExcelBookManager manager = new ExcelBookManager();

        public MainWindow()
        {
            InitializeComponent();

            dataGrid.DataError += OnDataError;
            dataGrid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
            typeof(Control).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance)
               .SetValue(dataGrid, true, null);
            //dataGrid.DataSource = new BindingList<DataEntity>();
        }

        private static void OnDataError(object sender, DataGridViewDataErrorEventArgs dataGridViewDataErrorEventArgs)
        {
            MessageBox.Show(dataGridViewDataErrorEventArgs.ToString());
        }

        private BindingList<DataEntity> GenerateDataTable(Workbook workbook)
        {
            var worksheet = (Worksheet)workbook.Worksheets[1];
            var range = worksheet.UsedRange;
            var rows = range.Rows.Count;
            var entites = new List<DataEntity>();

            progressBar.Show();
            progressBar.Maximum = rows;

            for (var i = 9; i <= rows; i++)
            {
                try
                {
                    Console.WriteLine(i);
                    var entity = new DataEntity
                    {
                        Date = DateTime.Parse((string) range[i, 1].Value2),
                        Document = (string) range[i, 2].Value2,
                        Analytics = (string) range[i, 4].Value2,
                        DebetRu = (decimal)(range[i, 7].Value2 as double? ?? 0),
                        CreditRu = (decimal)(range[i, 10].Value2 as double? ?? 0),
                    };

                    entites.Add(entity);
                    progressBar.Value = i;
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            progressBar.Hide();
            return new BindingList<DataEntity>(entites);
        }

        #region Events

        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        private void loadMenuButton_Click(object sender, EventArgs e)
        {
            var res = openFileDialog.ShowDialog();
            if (res == DialogResult.OK)
            {
                try
                {
                    var workbook = manager.OpenWorkbook(openFileDialog.FileName);
                    Console.WriteLine($"Succesfully opened {openFileDialog.FileName}");
                    dataGrid.RowHeadersVisible = false;
                    (dataGrid.DataSource as IDisposable)?.Dispose();
                    dataGrid.DataSource = GenerateDataTable(workbook);
                    dataGrid.RowHeadersVisible = true;
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                }
            }
        }

        private void saveMenuButton_Click(object sender, EventArgs e)
        {
            var res = saveFileDialog.ShowDialog();
            if (res == DialogResult.OK)
            {
                var data = dataGrid.DataSource as BindingList<DataEntity>;
                if (data == null)
                {
                    Console.WriteLine("Ooops");
                    return;
                }

                var workbook = manager.CreateWorkbook();
                var worksheet = (Worksheet) workbook.Worksheets[1];

                worksheet.Cells[1, 1] = "Период";
                worksheet.Cells[1, 2] = "Документ";
                worksheet.Cells[1, 3] = "Аналитика Дт";
                worksheet.Cells[1, 4] = "Дебет, рубли";
                worksheet.Cells[1, 5] = "Дебет, USD";
                worksheet.Cells[1, 6] = "Кредит, рубли";
                worksheet.Cells[1, 7] = "Кредит, USD";
                worksheet.Cells[1, 8] = "Тип";

                progressBar.Show();
                progressBar.Maximum = data.Count;

                for (var i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i+2, 1] = data[i].Date;
                    worksheet.Cells[i+2, 2] = data[i].Document;
                    worksheet.Cells[i+2, 3] = data[i].Analytics;
                    worksheet.Cells[i+2, 4] = data[i].DebetRu;
                    worksheet.Cells[i+2, 5] = data[i].DebetUsd;
                    worksheet.Cells[i+2, 6] = data[i].CreditRu;
                    worksheet.Cells[i+2, 7] = data[i].CreditUsd;
                    worksheet.Cells[i+2, 8] = data[i].Type;

                    progressBar.Value = i;
                }

                
                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();

                progressBar.Hide();
            }
        }

        private void saveFileDialog_FileOk(object sender, CancelEventArgs e)
        {

        }

        #endregion

        
    }
}

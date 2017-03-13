using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Meister_Hämmerlein.Core;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Meister_Hämmerlein.WindowForms
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();

            dataGrid.DataError += OnDataError;
            dataGrid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
            typeof(Control).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance)
               .SetValue(dataGrid, true, null);
            //dataGrid.DataSource = new BindingList<DataEntity>();
        }

        private BindingList<DataEntity> GenerateDataTable(ISheet workSheet)
        {
            var rows = workSheet.LastRowNum;
            var entites = new List<DataEntity>();

            progressBar.Show();
            progressBar.Maximum = rows;

            for (var i = 98; i <= rows; i++)
            {
                var row = workSheet.GetRow(i);
                if(row == null) continue;

                try
                {
                    var entity = new DataEntity
                    {
                        Date = DateTime.Parse(row.GetCell(0).StringCellValue),
                        Document = row.GetCell(1).StringCellValue,
                        Analytics = row.GetCell(3).StringCellValue,
                    };

                    var debetCell = row.GetCell(6)?.CellType;
                    var creditCell = row.GetCell(9)?.CellType;

                    if (debetCell != null && debetCell.Value == CellType.Numeric)
                        entity.DebetRu = (decimal) row.GetCell(6).NumericCellValue;
                    if (creditCell != null && creditCell.Value == CellType.Numeric)
                        entity.CreditRu = (decimal) row.GetCell(9).NumericCellValue;

                    entites.Add(entity);
                    progressBar.Value = i;
                }
                catch (FormatException)
                {
                    //ignored
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Format exception in {i} : {e}");
                }
            }

            progressBar.Hide();
            return new BindingList<DataEntity>(entites);
        }

        private void CreateHeader(ISheet workSheet)
        {
            var row = workSheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Период");
            row.CreateCell(1).SetCellValue("Документ");
            row.CreateCell(2).SetCellValue("Аналитика Дт");
            row.CreateCell(3).SetCellValue("Дебет, рубли");
            row.CreateCell(4).SetCellValue("Дебет, USD");
            row.CreateCell(5).SetCellValue("Кредит, рубли");
            row.CreateCell(6).SetCellValue("Кредит, USD");
            row.CreateCell(7).SetCellValue("Тип");
        }

        private void FillData(ISheet workSheet, int startRow, IList<DataEntity> data)
        {
            progressBar.Value = 0;
            progressBar.Maximum = data.Count;
            progressBar.Show();

            var newDataFormat = workSheet.Workbook.CreateDataFormat();
            var style = workSheet.Workbook.CreateCellStyle();

            style.DataFormat = newDataFormat.GetFormat("dd.MM.yyyy");

            for (int i = 0; i < data.Count; i++)
            {
                var row = workSheet.CreateRow(startRow + i);

                var cell = row.CreateCell(0);
                cell.SetCellValue(data[i].Date);
                cell.CellStyle = style;

                row.CreateCell(1).SetCellValue(data[i].Document);
                row.CreateCell(2).SetCellValue(data[i].Analytics);
                row.CreateCell(3).SetCellValue((double) data[i].DebetRu);
                row.CreateCell(4).SetCellValue((double) data[i].DebetUsd);
                row.CreateCell(5).SetCellValue((double) data[i].CreditRu);
                row.CreateCell(6).SetCellValue((double) data[i].CreditUsd);
                row.CreateCell(7).SetCellValue(data[i].Type);

                progressBar.Value = i;
            }

            progressBar.Hide();
        }

        private static void OnDataError(object sender, DataGridViewDataErrorEventArgs dataGridViewDataErrorEventArgs)
        {
            MessageBox.Show(dataGridViewDataErrorEventArgs.ToString());
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
                    HSSFWorkbook workbook;
                    using (var file = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new HSSFWorkbook(file);
                    }

                    var sheet = workbook.GetSheetAt(0);
                    var list = GenerateDataTable(sheet);
                    dataGrid.DataSource = list;
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
                var data = dataGrid.DataSource as IList<DataEntity>;
                if (data == null)
                {
                    Console.WriteLine("Data in table is strange..");
                    return;
                }

                var workbook = new HSSFWorkbook();
                var sheet = workbook.CreateSheet();
                CreateHeader(sheet);
                FillData(sheet, 1, data);

                using (var file = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
            }
        }

        private void saveFileDialog_FileOk(object sender, CancelEventArgs e)
        {

        }

        #endregion

        
    }
}

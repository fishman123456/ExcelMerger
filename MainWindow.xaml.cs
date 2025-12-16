using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

using ClosedXML.Excel;

namespace ExcelMerger
{
    public partial class MainWindow : Window
    {
        private string _file1Path;
        private string _file2Path;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseFile1_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel-файлы (*.xlsx)|*.xlsx",
                Title = "Выберите первый файл Excel"
            };

            if (dialog.ShowDialog() == true)
            {
                _file1Path = dialog.FileName;
                TxtFile1.Text = _file1Path;
                UpdateStatus($"Загружен файл 1: {Path.GetFileName(_file1Path)}");
            }
        }

        private void BrowseFile2_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel-файлы (*.xlsx)|*.xlsx",
                Title = "Выберите второй файл Excel"
            };

            if (dialog.ShowDialog() == true)
            {
                _file2Path = dialog.FileName;
                TxtFile2.Text = _file2Path;
                UpdateStatus($"Загружен файл 2: {Path.GetFileName(_file2Path)}");
            }
        }

        private void MergeFiles_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_file1Path) || string.IsNullOrEmpty(_file2Path))
            {
                MessageBox.Show("Пожалуйста, выберите оба файла.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                ProgressBar.Visibility = Visibility.Visible;
                ProgressBar.IsIndeterminate = true;

                // Читаем оба файла
                var workbook1 = new XLWorkbook(_file1Path);
                var workbook2 = new XLWorkbook(_file2Path);

                var ws1 = workbook1.Worksheet(1);
                var ws2 = workbook2.Worksheet(1);

                // Словарь для быстрого поиска: ключ — значение из 1‑го столбца файла 2
                var dict = new Dictionary<string, IXLRow>();
                foreach (var row in ws2.RowsUsed())
                {
                    var key = row.Cell(1).Value.ToString().Trim();
                    if (!string.IsNullOrEmpty(key))
                    {
                        dict[key] = row;
                    }
                }

                // Новый файл для результата
                var resultWorkbook = new XLWorkbook();
                var resultWs = resultWorkbook.Worksheets.Add("Объединённые данные");

                int rowNum = 1;

                // Копируем заголовки из файла 1
                var headerRow = ws1.Row(1);
                headerRow.CopyTo(resultWs.Row(rowNum));
                rowNum++;

                // Проходим по строкам файла 1 (начиная со 2‑й)
                foreach (var srcRow in ws1.RowsUsed().Skip(1))
                {
                    var key = srcRow.Cell(1).Value.ToString().Trim();

                    if (dict.ContainsKey(key))
                    {
                        // Копируем строку из файла 1
                        srcRow.CopyTo(resultWs.Row(rowNum));

                        // Добавляем данные из совпадающей строки файла 2 (начиная со 2‑го столбца)
                        var matchRow = dict[key];
                        int colNum = ws1.LastColumnUsed().ColumnNumber() + 1;
                        foreach (var cell in matchRow.Cells())
                        {
                            if (cell.Address.ColumnNumber > 1) // пропускаем 1‑й столбец
                            {
                                resultWs.Cell(rowNum, colNum).Value = cell.Value;
                                colNum++;
                            }
                        }

                        rowNum++;
                    }
                }

                // Сохраняем результат
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel-файлы (*.xlsx)|*.xlsx",
                    FileName = "Объединённый_файл.xlsx",
                    Title = "Сохранить объединённый файл"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    resultWorkbook.SaveAs(saveDialog.FileName);
                    UpdateStatus($"Объединение завершено! Файл сохранён: {saveDialog.FileName}");
                }

                resultWorkbook.Dispose();
                workbook1.Dispose();
                workbook2.Dispose();

                ProgressBar.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                ProgressBar.Visibility = Visibility.Collapsed;
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus("Ошибка: " + ex.Message);
            }
        }

        private void UpdateStatus(string message)
        {
            TxtStatus.Text += $"{DateTime.Now:HH:mm:ss} > {message}\r\n";
        }
    }
}

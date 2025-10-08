using Microsoft.Win32;
using OfficeOpenXml;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Navigation;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;

namespace Excel_Diff_Remake
{
    public partial class MainWindow : Window
    {
        private string filePath1;
        private string filePath2;

        const int MAX_CELLS = 5000000;

        public MainWindow()
        {
            InitializeComponent();

            ExcelPackage.License.SetNonCommercialPersonal("Jonas Thaun");

            progressBar1.Visibility = Visibility.Collapsed;
        }

        private void File1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Title = "Select an Excel file";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (ofd.ShowDialog() == true) 
            {
                ResetTable();
                filePath1 = ofd.FileName; // path

                file1Label.Content = Path.GetFileName(filePath1); // only name

                progressBar1.Visibility = Visibility.Visible;
                progressBar1.Value = 0;
            }
        }

        private void File2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Title = "Select an Excel file";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (ofd.ShowDialog() == true)
            {
                ResetTable();
                filePath2 = ofd.FileName; // path

                file2Label.Content = Path.GetFileName(filePath2); // only name

                progressBar1.Visibility = Visibility.Visible;
                progressBar1.Value = 0;
            }
        }

        private async void ShowDifference_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(filePath1) || string.IsNullOrWhiteSpace(filePath2))
            {
                MessageBox.Show("Select two .xlsx files", "Missing Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                int maxRowFinal;

                using (ExcelPackage packageA = new ExcelPackage(new FileInfo(filePath1)))
                using (ExcelPackage packageB = new ExcelPackage(new FileInfo(filePath2)))
                {
                    maxRowFinal = Math.Max(packageA.Workbook.Worksheets[0].Dimension?.End?.Row ?? 0,
                                           packageB.Workbook.Worksheets[0].Dimension?.End?.Row ?? 0);
                }

                compareEverythingButton.IsEnabled = false;
                compareDifferenceButton.IsEnabled = false;
                file1Button.IsEnabled = false;
                file2Button.IsEnabled = false;

                int totalSteps = maxRowFinal * 2;
                progressBar1.Maximum = totalSteps;
                progressBar1.Value = 0;

                var progressHandler = new Progress<int>(rowsLoaded =>
                {
                    progressBar1.Value = rowsLoaded;
                });

                List<(string cell, string valueA, string valueB)> diffs = await Task.Run(() =>
                {
                    return CompareFilesLogic(progressHandler);
                });

                if (diffs.Count > MAX_CELLS)
                {
                    MessageBox.Show($"Limit: The files have more than {MAX_CELLS} cells.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    compareEverythingButton.IsEnabled = true;
                    compareDifferenceButton.IsEnabled = true;
                    file1Button.IsEnabled = true;
                    file2Button.IsEnabled = true;
                    return;
                }

                int uiStartValue = maxRowFinal;
                int finalMaximum = uiStartValue + diffs.Count; 

                progressBar1.Maximum = finalMaximum;

                dataGridMain.Items.Clear(); 
                dataGridMain.Columns.Clear();

                dataGridMain.Columns.Add(new DataGridTextColumn() { Header = "Cell (with difference)", Binding = new Binding("Cell")});
                dataGridMain.Columns.Add(new DataGridTextColumn() { Header = "File 1", Binding = new Binding("Value1")});
                dataGridMain.Columns.Add(new DataGridTextColumn() { Header = "File 2", Binding = new Binding("Value2")});

                int i = 0;
                foreach (var diff in diffs)
                {
                    dataGridMain.Items.Add(new { Cell = diff.cell, Value1 = diff.valueA, Value2 = diff.valueB });

                    i++;
                    progressBar1.Value = uiStartValue + i;

                    await Task.Yield();
                }

                progressBar1.Value = progressBar1.Maximum;

                compareEverythingButton.IsEnabled = true;
                compareDifferenceButton.IsEnabled = true;
                file1Button.IsEnabled = true;
                file2Button.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel files:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
            finally
            {
                compareEverythingButton.IsEnabled = true;
                compareDifferenceButton.IsEnabled = true;
                file1Button.IsEnabled = true;
                file2Button.IsEnabled = true;
            }
        }

        private List<(string, string, string)> CompareFilesLogic(IProgress<int> progress)
        {
            List<(string, string, string)> differences = new List<(string, string, string)>();
            using (ExcelPackage packageA = new ExcelPackage(new FileInfo(filePath1)))
            using (ExcelPackage packageB = new ExcelPackage(new FileInfo(filePath2)))
            {
                ExcelWorksheet wsA = packageA.Workbook.Worksheets[0];
                ExcelWorksheet wsB = packageB.Workbook.Worksheets[0];

                int maxRow = Math.Max(wsA.Dimension.End.Row, wsB.Dimension.End.Row);
                int maxCol = Math.Max(wsA.Dimension.End.Column, wsB.Dimension.End.Column);

                int rowsProcessed = 0;

                for (int row = 1; row <= maxRow; row++)
                {
                    for (int col = 1; col <= maxCol; col++)
                    {
                        string valueA = wsA.Cells[row, col].Text == "" ? "-" : wsA.Cells[row, col].Text;
                        string valueB = wsB.Cells[row, col].Text == "" ? "-" : wsB.Cells[row, col].Text;

                        if (!valueA.Equals(valueB))
                        {
                            string cell = row.ToString() + " - " + ColumnNumberToLetter(col);
                            differences.Add((cell, valueA, valueB));
                        }
                    }

                    rowsProcessed++;
                    progress.Report(rowsProcessed);
                }
            }
            return differences;
        }

        private async void CompareEverything_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(filePath1) || string.IsNullOrWhiteSpace(filePath2))
            {
                MessageBox.Show("Select two .xlsx files", "Missing Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                int maxRowFinal;

                compareEverythingButton.IsEnabled = false;
                compareDifferenceButton.IsEnabled = false;
                file1Button.IsEnabled = false;
                file2Button.IsEnabled = false;

                using (ExcelPackage packageA = new ExcelPackage(new FileInfo(filePath1)))
                using (ExcelPackage packageB = new ExcelPackage(new FileInfo(filePath2)))
                {
                    ExcelWorksheet wsA = packageA.Workbook.Worksheets[0];
                    ExcelWorksheet wsB = packageB.Workbook.Worksheets[0];

                    int maxRowTemp = wsA.Dimension?.End?.Row ?? 0;
                    int maxColTemp = wsA.Dimension?.End?.Column ?? 0;
                    int maxRowTempB = wsB.Dimension?.End?.Row ?? 0;
                    int maxColTempB = wsB.Dimension?.End?.Column ?? 0;

                    maxRowFinal = Math.Max(maxRowTemp, maxRowTempB);
                    int maxColFinal = Math.Max(maxColTemp, maxColTempB);

                    long totalCells = (long)maxRowFinal * maxColFinal;

                    int totalSteps = maxRowFinal * 2;
                    progressBar1.Maximum = totalSteps;
                    progressBar1.Value = 0;

                    if (totalCells > MAX_CELLS)
                    {
                        double limitInMillions = Math.Round((double)MAX_CELLS / 1000000, 1);

                        MessageBox.Show($"Limit: The files have more than {MAX_CELLS} cells.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);

                        compareEverythingButton.IsEnabled = true;
                        compareDifferenceButton.IsEnabled = true;
                        file1Button.IsEnabled = true;
                        file2Button.IsEnabled = true;
                        return;
                    }
                }

                var progressHandler = new Progress<int>(rowsLoaded =>
                {
                    progressBar1.Value = rowsLoaded;
                });

                List<List<string>> result = await Task.Run(() =>
                {
                    return ShowEverythingLogic(progressHandler);
                });

                int uiStartValue = maxRowFinal;
                progressBar1.Value = uiStartValue;

                dataGridMain.Items.Clear();
                dataGridMain.Columns.Clear();

                dataGridMain.Columns.Add(new DataGridTextColumn()
                {
                    Header = "",
                    Binding = new Binding("[RowNum]"),
                    IsReadOnly = true,
                    ElementStyle = (Style)this.FindResource("BoldCellStyle")
                });

                int maxCol = result.Count > 0 ? result[0].Count : 0;

                for (int col = 1; col <= maxCol; col++)
                {
                    string columnKey = $"Column{col}";

                    DataGridTextColumn newCol = new DataGridTextColumn();
                    newCol.Header = ColumnNumberToLetter(col);

                    newCol.Binding = new Binding($"[{columnKey}]");

                    newCol.HeaderStyle = (Style)this.FindResource("BoldHeaderStyle");

                    dataGridMain.Columns.Add(newCol);
                }

                for (int row = 0; row < result.Count; row++)
                {
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    rowData["RowNum"] = (row + 1).ToString();

                    for (int col = 0; col < result[row].Count; col++)
                    {
                        rowData[$"Column{col + 1}"] = result[row][col];
                    }

                    dataGridMain.Items.Add(rowData);

                    progressBar1.Value = uiStartValue + (row + 1);

                    if (row % 50 == 0)
                    {
                        await Task.Yield();
                    }
                }

                compareEverythingButton.IsEnabled = true;
                compareDifferenceButton.IsEnabled = true;
                file1Button.IsEnabled = true;
                file2Button.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel files:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
            finally
            {
                compareEverythingButton.IsEnabled = true;
                compareDifferenceButton.IsEnabled = true;
                file1Button.IsEnabled = true;
                file2Button.IsEnabled = true;
            }
        }

        private List<List<string>> ShowEverythingLogic(IProgress<int> progress)
        {
            List<List<string>> result = new List<List<string>>();
            using (ExcelPackage packageA = new ExcelPackage(new FileInfo(filePath1)))
            using (ExcelPackage packageB = new ExcelPackage(new FileInfo(filePath2)))
            {
                ExcelWorksheet wsA = packageA.Workbook.Worksheets[0];
                ExcelWorksheet wsB = packageB.Workbook.Worksheets[0];

                int maxRow = Math.Max(wsA.Dimension.End.Row, wsB.Dimension.End.Row);
                int maxCol = Math.Max(wsA.Dimension.End.Column, wsB.Dimension.End.Column);

                int totalRowsToProcess = maxRow;
                int rowsProcessed = 0;

                for (int row = 1; row <= maxRow; row++)
                {
                    List<string> rowValues = new List<string>();
                    for (int col = 1; col <= maxCol; col++)
                    {
                        string valueA = wsA.Cells[row, col].Text == "" ? "-" : wsA.Cells[row, col].Text;
                        string valueB = wsB.Cells[row, col].Text == "" ? "-" : wsB.Cells[row, col].Text;

                        string cellText = (valueA.Equals(valueB)) ? valueA : valueA + "\r\n" + valueB;
                        rowValues.Add(cellText);
                    }
                    result.Add(rowValues);

                    rowsProcessed++;
                    progress.Report(rowsProcessed);
                }
            }
            return result;
        }

        private string ColumnNumberToLetter(int col)
        {
            string letter = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                letter = (char)(rem + 'A') + letter;
                col = (col - 1) / 26;
            }
            return letter;
        }

        private void ResetTable()
        {
            dataGridMain.Items.Clear();
            dataGridMain.Columns.Clear();

            progressBar1.Visibility = Visibility.Collapsed;
        }
    }

    
}
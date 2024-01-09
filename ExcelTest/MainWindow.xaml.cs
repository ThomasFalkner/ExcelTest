using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;


namespace ExcelTest
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Helpers

        private void ReleaseCOMObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        #endregion

        private void CreateExcel(string message, int count)
        {
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = true;
                dynamic workbooks = excelApp.Workbooks;
                dynamic workbook = workbooks.Add(Type.Missing);
                dynamic worksheet = workbook.Worksheets[1];
                for (int i = 1; i <= count; i++)
                {
                    worksheet.Cells[i, 1] = message;
                }
                ReleaseCOMObject(worksheet);
                ReleaseCOMObject(workbook);
                ReleaseCOMObject(excelApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            string message = textBoxMessage.Text ?? "Hallo Welt!";           
            int count = Convert.ToInt32(textBoxCount.Text);
            if (count < 1 || count > 1000)
            {
                MessageBox.Show("Bitte geben Sie eine Zahl zwischen 1 und 100 ein", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            CreateExcel(message, count);
        }

    }
}

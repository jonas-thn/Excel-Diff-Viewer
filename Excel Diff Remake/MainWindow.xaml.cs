using Microsoft.Win32;
using OfficeOpenXml;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace Excel_Diff_Remake
{
    public partial class MainWindow : Window
    {
        private string filePath1;
        private string filePath2;

        public MainWindow()
        {
            InitializeComponent();

            ExcelPackage.License.SetNonCommercialPersonal("Jonas Thaun");

            progressBar1.Visibility = Visibility.Collapsed;
        }

        private void File1_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void File2_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ShowDifference_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CompareEverything_Click(object sender, RoutedEventArgs e)
        {

        }
    }

    
}
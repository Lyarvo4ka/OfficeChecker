using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using winForms = System.Windows.Forms;

namespace OfficeChecker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> wordExtensionList_ ;
        private List<string> excelExtensionList_;


        public MainWindow()
        {
            InitializeComponent();
            wordExtensionList_ = new List<string> { };
            OnWordClicked(docCheckBox, null);
            excelExtensionList_ = new List<string> { };
            OnExcelClicked(xlsCheckBox, null);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {


            winForms.FolderBrowserDialog dirBrowse = new winForms.FolderBrowserDialog
            {
                RootFolder = Environment.SpecialFolder.MyComputer
            };
            winForms.DialogResult dirResult = dirBrowse.ShowDialog();

            if (dirResult == winForms.DialogResult.OK)
                TextEditFolder.Text = dirBrowse.SelectedPath;

        }
        private void UpdateExtensionList(ref CheckBox check_box , List<string> extList)
        {
            if (check_box != null)
                if (check_box.Content != null)
                {
                    if (check_box.IsChecked == true)
                        extList.Add(check_box.Content.ToString());
                }
        }

       string GetStringFromExtList(List<string> extList , string delimiter = " ")
        {
            string fullName = "";
            foreach (var theName in extList)
            {
                fullName += theName + delimiter;
            }
            return fullName;
        }

        private void OnWordClicked(object sender, RoutedEventArgs e)
        {
            wordExtensionList_.Clear();
            UpdateExtensionList(ref docCheckBox, wordExtensionList_);
            UpdateExtensionList(ref docxCheckBox, wordExtensionList_);
            UpdateExtensionList(ref rtfCheckBox, wordExtensionList_);

            string headerNewName = GetStringFromExtList(wordExtensionList_);
            if (String.IsNullOrEmpty(headerNewName))
            {
                WordExpaner.Header = "no Word";
            }
            else
                WordExpaner.Header = headerNewName;
        }
        private void OnExcelClicked(object sender, RoutedEventArgs e)
        {
            excelExtensionList_.Clear();
            UpdateExtensionList(ref xlsCheckBox, excelExtensionList_);
            UpdateExtensionList(ref xlsxCheckBox, excelExtensionList_);

            string headerNewName = GetStringFromExtList(excelExtensionList_);
            if (String.IsNullOrEmpty(headerNewName))
            {
                ExcelExapander.Header = "no Excel";
            }
            else
                ExcelExapander.Header = headerNewName;

        }

    }
}

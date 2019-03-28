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
        private string docExt_ ;
        private string docxExt_;
        private string rtfExt_;

        public MainWindow()
        {
            InitializeComponent();
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
        private void onDocChecked(object sender, RoutedEventArgs e)
        {
            List<string> wordExtList = new List<string>();
            if (docCheckBox != null)
            if (docCheckBox.Content != null)
            {
                if (docCheckBox.IsChecked == true)
                    wordExtList.Add(docCheckBox.Content.ToString());
            }
            if (docxCheckBox != null)
            if (docxCheckBox.Content != null)
            {
                if (docxCheckBox.IsChecked == true)
                    wordExtList.Add(docxCheckBox.Content.ToString());
            }
            if (rtfCheckBox != null)
            if (rtfCheckBox.Content != null)
            {
                if (rtfCheckBox.IsChecked == true)
                    wordExtList.Add(rtfCheckBox.Content.ToString());
            }

            string headerNewName = "";
            foreach (var theName in wordExtList)
            {
                headerNewName += theName;
            }
            WordExpaner.Header = headerNewName;

            //WordExpaner.Header = docCheckBox.Content.ToString();

            //docCheckBox.isC
            //var checkBox = sender as CheckBox;
            ////var checkBox = ((CheckBox)sender);
            //if (checkBox.Content!=null)
            //if (checkBox.IsChecked != null && checkBox.IsChecked.Value)
            //{
            //    WordExpaner.Header = checkBox.Content.ToString();
            //}
        }
        private void onDocxChecked(object sender, RoutedEventArgs e)
        {

        }
        private void onRtfChecked(object sender, RoutedEventArgs e)
        {

        }
    }
}

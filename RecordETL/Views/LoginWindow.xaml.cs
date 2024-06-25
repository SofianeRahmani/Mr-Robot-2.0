using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace RecordETL.Views
{
    public partial class LoginWindow 
    {
        private List<(string sheetName, string[] headers)> SheetNames;
        ViewModels.ExcelViewModel ViewModel => (ViewModels.ExcelViewModel)DataContext;

        public LoginWindow()
        {
            InitializeComponent();

            DataContext = new ViewModels.ExcelViewModel();

            DropArea.AllowDrop = true;
            DropArea.Drop += DropArea_Drop;
            DropArea.DragEnter += DropArea_DragEnter;
        }


        private void DropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files[0].EndsWith(".xlsx"))
                {
                    e.Effects = DragDropEffects.Copy;
                }
            }
        }

        private void DropArea_Drop(object sender, DragEventArgs e)
        {
            string[]? files = (string[])e.Data.GetData(DataFormats.FileDrop);

            ViewModel.ExcelPath = files[0];
        }

        private void DropArea_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                ViewModel.ExcelPath = dlg.FileName;
            }
        }

        private void LoadExcelData()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var fileInfo = new FileInfo(ViewModel.ExcelPath);
                using var package = new ExcelPackage(fileInfo);
                SheetNames = new List<(string sheetName, string[] headers)>();
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var headers = new List<string>();
                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        headers.Add(firstRowCell.Text);
                    }

                    SheetNames.Add((sheetName: worksheet.Name,
                        headers: headers.Where(x => !string.IsNullOrEmpty(x)).ToArray()));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                ViewModel.OutputPath = dialog.FileName;
            }
        }

        private void RadioMembre_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienMembres.Visibility = Visibility.Visible;
            TextBoxForRadioMembres.Visibility = Visibility.Visible;
        }

        private void RadioEmplois_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienEmplois.Visibility = Visibility.Visible;
            TextBoxForRadioEmplois.Visibility = Visibility.Visible;
        }

        private void RadioEmployeurs_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienEmployeur.Visibility = Visibility.Visible;
            TextBoxForRadioEmployeur.Visibility = Visibility.Visible;
        }

        private void RadioFonctions_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienFonctions.Visibility = Visibility.Visible;
            TextBoxForRadioFonctions.Visibility = Visibility.Visible;
        }

        private void RadioSecteurs_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienSecteurs.Visibility = Visibility.Visible;
            TextBoxForRadioSecteurs.Visibility = Visibility.Visible;
        }

        private void RadioEvenement_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienEvenements.Visibility = Visibility.Visible;
            TextBoxForRadioEvenements.Visibility = Visibility.Visible;
        }

        private void RadioTransactions_Checked(object sender, RoutedEventArgs e)
        {
            TextBlockCombienTransactions.Visibility = Visibility.Visible;
            TextBoxForRadioTransactions.Visibility = Visibility.Visible;
        }


        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            LoadExcelData();
            Window tabsWindow = new Window
            {
                Title = "Configuations",
                Content = new TabControl()
            };

            AddTabIfChecked(RadioMembre2, TextBoxForRadioMembres, (TabControl)tabsWindow.Content, "Membre");
            AddTabIfChecked(RadioEmplois2, TextBoxForRadioEmplois, (TabControl)tabsWindow.Content, "Emplois");
            AddTabIfChecked(RadioEmployeurs2, TextBoxForRadioEmployeur, (TabControl)tabsWindow.Content, "Employeurs");
            AddTabIfChecked(RadioFonctions2, TextBoxForRadioFonctions, (TabControl)tabsWindow.Content, "Fonctions");
            AddTabIfChecked(RadioSecteurs2, TextBoxForRadioSecteurs, (TabControl)tabsWindow.Content, "Secteurs");
            AddTabIfChecked(RadioEvenement2, TextBoxForRadioEvenements, (TabControl)tabsWindow.Content, "Événement");
            AddTabIfChecked(RadioTransactions2, TextBoxForRadioTransactions, (TabControl)tabsWindow.Content,
                "Transactions");

            tabsWindow.Show();
        }

        
        private void AddTabIfChecked(CheckBox checkBox, TextBox textBox, TabControl tabControl, string tabName)
        {
            if (checkBox.IsChecked != true || !int.TryParse(textBox.Text, out int numTabs))
                return;

            for (int i = 0; i < numTabs; i++)
            {
                TabItem tabItem = new TabItem
                {
                    Header = $"{tabName} {i + 1}"
                };


                switch (tabName)
                {
                    case "Membre":
                        tabItem.Content = CreateMembreTabContent();
                        break;

                    default:
                        break;
                }

                tabControl.Items.Add(tabItem);
            }
        }

        private ScrollViewer CreateMembreTabContent()
        {
            ScrollViewer scrollViewer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel stackPanel = new StackPanel { Margin = new Thickness(10) };

            stackPanel.Children.Add(new Label { Content = "Member Number" });
            ComboBox cbMemberNumbereuille = new ComboBox { Name = "feuille", Margin = new Thickness(0, 0, 0, 10) };
            stackPanel.Children.Add(cbMemberNumbereuille);
            ComboBox cbMemberNumberColonne = new ComboBox { Name = "Colonne", Margin = new Thickness(0, 0, 0, 10) };
            stackPanel.Children.Add(cbMemberNumberColonne);
            


            foreach (var sheetInfo in SheetNames)
            {
                cbMemberNumbereuille.Items.Add(sheetInfo.sheetName);
            }

            cbMemberNumbereuille.SelectionChanged += (sender, args) =>
            {
                ComboBox cb = sender as ComboBox;
                string selectedSheet = cb.SelectedItem?.ToString();
                cbMemberNumberColonne.Items.Clear();

                if (!string.IsNullOrEmpty(selectedSheet))
                {
                    var sheetInfo = SheetNames.FirstOrDefault(s => s.sheetName == selectedSheet);
                    if (sheetInfo.headers != null)
                    {
                        foreach (var header in sheetInfo.headers)
                        {
                            cbMemberNumberColonne.Items.Add(header);
                        }
                    }
                }
            };

            scrollViewer.Content = stackPanel;
            return scrollViewer;
        }
    }
}
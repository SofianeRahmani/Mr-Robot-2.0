﻿using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows;
using System.Windows.Controls;

namespace RecordETL.Views
{
    public partial class ExcelView : UserControl
    {
        ViewModels.ExcelViewModel ViewModel => (ViewModels.ExcelViewModel)DataContext;
        public ExcelView()
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
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            // Process the Excel file, e.g., read its content.

            ViewModel.ExcelPath = files[0];
        }

        private void DropArea_Click(object sender, RoutedEventArgs e)
        {
            // write the code to open file dialog and select only excel files
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                ViewModel.ExcelPath = dlg.FileName;
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
    }
}

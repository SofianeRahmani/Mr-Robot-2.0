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
    }
}

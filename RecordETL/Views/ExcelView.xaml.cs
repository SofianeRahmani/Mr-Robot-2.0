using System.Windows.Controls;

namespace RecordETL.Views
{
    public partial class ExcelView : UserControl
    {
        public ExcelView()
        {
            InitializeComponent();

            DataContext = new ViewModels.ExcelViewModel();
        }
    }
}

using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace RecordETL.Views
{
    public partial class DetailTabsWindow : Window
    {
        public DetailTabsWindow(Dictionary<string, List<string>> sheetColumns)
        {
            InitializeComponent();
            PopulateTabs(sheetColumns);
        }

        private void PopulateTabs(Dictionary<string, List<string>> sheetColumns)
        {
            foreach (var sheet in sheetColumns)
            {
                var tabItem = new TabItem { Header = sheet.Key };
                var listBox = new ListBox();

                foreach (var column in sheet.Value)
                {
                    listBox.Items.Add(column);
                }

                tabItem.Content = new ScrollViewer { Content = listBox };
                tabControl.Items.Add(tabItem);
            }
        }
    }
}
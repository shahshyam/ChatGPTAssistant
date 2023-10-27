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

namespace AssistantForWord.UI.Views
{
    /// <summary>
    /// Interaction logic for CustomPanelView.xaml
    /// </summary>
    public partial class CustomPanelView : UserControl
    {
        public CustomPanelView()
        {
            InitializeComponent();
            this.Loaded += OnCustomPanelViewLoaded;
            Globals.ThisAddIn.OnRefreshSidebarPanel += OnRefreshSidebarPanel;
        }

        private void OnRefreshSidebarPanel()
        {
            RefreshValue();
        }

        private void OnCustomPanelViewLoaded(object sender, RoutedEventArgs e)
        {
            RefreshValue();
        }
        internal void RefreshValue()
        {
            var model = this.DataContext as ViewModels.CustomPanelViewModel;
            if (model != null)
            {
                model.LoadExistingData();
            }
        }
    }
}

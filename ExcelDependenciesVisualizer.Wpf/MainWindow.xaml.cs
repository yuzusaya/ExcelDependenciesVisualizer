using ExcelDependenciesVisualizer.Wpf.ViewModels;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelDependenciesVisualizer.Wpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        public MainWindow()
        {
            InitializeComponent();
            DataContext = _viewModel = new MainViewModel();
            _viewModel.SearchExecuted += SearchExecuted;
        }

        private void SearchExecuted(object? sender, EventArgs e)
        {
            searchButton.Focus();
        }

        private void CellTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                _viewModel.SearchReferencedByCellsCommand?.Execute(null);
            }
        }
    }
}
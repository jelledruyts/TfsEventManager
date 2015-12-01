using Microsoft.TeamFoundation.Framework.Client;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace TfsEventManager
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                this.ViewModel = new MainWindowViewModel();
            }
            catch (FileNotFoundException exc)
            {
                ShowException(exc);
                Application.Current.Shutdown();
            }
        }

        internal static void ShowException(Exception exc)
        {
            MessageBox.Show(exc.Message, "An error occurred", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public MainWindowViewModel ViewModel
        {
            get { return (MainWindowViewModel)this.DataContext; }
            set { this.DataContext = value; }
        }

        private void eventSubscriptionsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.ViewModel.SelectedEventSubscriptions = this.eventSubscriptionsDataGrid.SelectedItems.Cast<Subscription>().ToArray();
        }
    }
}
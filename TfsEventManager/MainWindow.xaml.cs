using Microsoft.TeamFoundation.Framework.Client;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
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
                if (exc.Message.Contains("Version=10.0.0.0"))
                {
                    var entryAssemblyName = Assembly.GetEntryAssembly().GetName();
                    var message = "This application was compiled against the Team Explorer 2010 assemblies, which you don't seem to have installed." + Environment.NewLine + Environment.NewLine;
                    message += string.Format(CultureInfo.CurrentCulture, "You can redirect the assembly versions by uncommenting the appropriate lines in the \"{0}.exe.config\" file, depending on the version of Team Explorer you do have installed.", entryAssemblyName.Name);
                    MessageBox.Show(message, "Team Explorer Version Mismatch", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    ShowException(exc);
                }
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
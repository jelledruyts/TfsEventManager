using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using TfsEventManager.Infrastructure;

namespace TfsEventManager
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        #region Fields

        private string statusText;
        private string statusDetail;
        private Brush statusColorForeground;
        private Brush statusColorBackground;
        private NotificationJobLogLevel? notificationJobLogLevel;
        private bool hideSucceededEventJobHistoryEntries;
        private IList<TeamFoundationJobHistoryEntry> allEventJobHistoryEntries;

        #endregion

        #region Properties

        public RelayCommand GetAllEventSubscriptionsCommand { get; private set; }
        public RelayCommand DeleteSelectedEventSubscriptionsCommand { get; private set; }
        public RelayCommand GetEventJobHistoryCommand { get; private set; }
        public RelayCommand GetNotificationJobLogLevelCommand { get; private set; }
        public RelayCommand SetNotificationJobLogLevelCommand { get; private set; }
        public IList<string> AvailableTeamProjectCollectionUrls { get; private set; }
        public string TeamProjectCollectionUrl { get; set; }
        public ObservableCollection<Subscription> EventSubscriptions { get; private set; }
        public IList<Subscription> SelectedEventSubscriptions { get; set; }
        public ObservableCollection<TeamFoundationJobHistoryEntry> EventJobHistoryEntries { get; private set; }
        public bool HideSucceededEventJobHistoryEntries { get { return this.hideSucceededEventJobHistoryEntries; } set { this.hideSucceededEventJobHistoryEntries = value; FilterEventJobHistoryEntries(); } }
        public NotificationJobLogLevel? NotificationJobLogLevel { get { return this.notificationJobLogLevel; } set { this.notificationJobLogLevel = value; OnPropertyChanged("NotificationJobLogLevel"); } }
        public string StatusText { get { return this.statusText; } set { this.statusText = value; OnPropertyChanged("StatusText"); } }
        public string StatusDetail { get { return this.statusDetail; } set { this.statusDetail = value; OnPropertyChanged("StatusDetail"); } }
        public Brush StatusColorForeground { get { return this.statusColorForeground; } set { this.statusColorForeground = value; OnPropertyChanged("StatusColorForeground"); } }
        public Brush StatusColorBackground { get { return this.statusColorBackground; } set { this.statusColorBackground = value; OnPropertyChanged("StatusColorBackground"); } }

        #endregion

        #region Constructors

        public MainWindowViewModel()
        {
            this.GetAllEventSubscriptionsCommand = new RelayCommand(GetAllEventSubscriptions, CanGetAllEventSubscriptions);
            this.DeleteSelectedEventSubscriptionsCommand = new RelayCommand(DeleteSelectedEventSubscriptions, CanDeleteSelectedEventSubscriptions);
            this.GetEventJobHistoryCommand = new RelayCommand(GetEventJobHistory, CanGetEventJobHistory);
            this.GetNotificationJobLogLevelCommand = new RelayCommand(GetNotificationJobLogLevel, CanGetNotificationJobLogLevel);
            this.SetNotificationJobLogLevelCommand = new RelayCommand(SetNotificationJobLogLevel, CanSetNotificationJobLogLevel);
            this.AvailableTeamProjectCollectionUrls = RegisteredTfsConnections.GetProjectCollections().Select(c => c.Uri.ToString()).OrderBy(u => u).ToArray();
            this.TeamProjectCollectionUrl = this.AvailableTeamProjectCollectionUrls.FirstOrDefault();
            this.EventSubscriptions = new ObservableCollection<Subscription>();
            this.EventJobHistoryEntries = new ObservableCollection<TeamFoundationJobHistoryEntry>();
            SetStatus(string.Format(CultureInfo.CurrentCulture, "Discovered {0} registered Team Project Collection(s).", this.AvailableTeamProjectCollectionUrls.Count));
        }

        #endregion

        #region GetAllEventSubscriptions Command

        private bool CanGetAllEventSubscriptions(object argument)
        {
            return !string.IsNullOrEmpty(this.TeamProjectCollectionUrl);
        }

        private void GetAllEventSubscriptions(object argument)
        {
            var worker = new BackgroundWorker();
            worker.DoWork += (sender, e) =>
            {
                SetStatus("Working...");
                using (var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(this.TeamProjectCollectionUrl)))
                {
                    var eventService = tfs.GetService<IEventService>();
                    e.Result = eventService.GetAllEventSubscriptions();
                }
            };
            worker.RunWorkerCompleted += (sender, e) =>
            {
                if (e.Error != null)
                {
                    HandleException(e.Error);
                }
                else
                {
                    var subscriptions = (IList<Subscription>)e.Result;
                    this.EventSubscriptions.Clear();
                    foreach (var subscription in subscriptions)
                    {
                        this.EventSubscriptions.Add(subscription);
                    }
                    SetStatus(string.Format(CultureInfo.CurrentCulture, "Retrieved {0} event subscription(s).", subscriptions.Count()));
                }
            };
            worker.RunWorkerAsync();
        }

        #endregion

        #region DeleteSelectedEventSubscriptions Command

        private bool CanDeleteSelectedEventSubscriptions(object argument)
        {
            return !string.IsNullOrEmpty(this.TeamProjectCollectionUrl) && this.SelectedEventSubscriptions != null && this.SelectedEventSubscriptions.Any();
        }

        private void DeleteSelectedEventSubscriptions(object argument)
        {
            var subscriptionsToDelete = this.SelectedEventSubscriptions.ToArray();
            var result = MessageBox.Show(string.Format(CultureInfo.CurrentCulture, "This will permanently delete the {0} selected subscription(s). Are you sure?", subscriptionsToDelete.Length), "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes)
            {
                return;
            }
            var worker = new BackgroundWorker();
            worker.DoWork += (sender, e) =>
            {
                SetStatus("Working...");
                using (var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(this.TeamProjectCollectionUrl)))
                {
                    var eventService = tfs.GetService<IEventService>();
                    foreach (var subscriptionToDelete in subscriptionsToDelete)
                    {
                        eventService.UnsubscribeEvent(subscriptionToDelete.ID);
                    }
                }
            };
            worker.RunWorkerCompleted += (sender, e) =>
            {
                if (e.Error != null)
                {
                    HandleException(e.Error);
                }
                else
                {
                    foreach (var subscriptionToDelete in subscriptionsToDelete)
                    {
                        this.EventSubscriptions.Remove(subscriptionToDelete);
                    }
                    SetStatus(string.Format(CultureInfo.CurrentCulture, "Deleted {0} event subscription(s).", subscriptionsToDelete.Length));
                }
            };
            worker.RunWorkerAsync();
        }

        #endregion

        #region GetEventJobHistory Command

        private bool CanGetEventJobHistory(object argument)
        {
            return !string.IsNullOrEmpty(this.TeamProjectCollectionUrl);
        }

        private void GetEventJobHistory(object argument)
        {
            var worker = new BackgroundWorker();
            worker.DoWork += (sender, e) =>
            {
                SetStatus("Working...");
                using (var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(this.TeamProjectCollectionUrl)))
                {
                    var jobService = tfs.GetService<ITeamFoundationJobService>();
                    // Retrieve the history for the "Team Foundation Server Event Processing" job (using its Guid).
                    e.Result = jobService.QueryJobHistory(new Guid[] { new Guid("a4804dcf-4bb6-4109-b61c-e59c2e8a9ff7") });
                }
            };
            worker.RunWorkerCompleted += (sender, e) =>
            {
                if (e.Error != null)
                {
                    HandleException(e.Error);
                }
                else
                {
                    this.allEventJobHistoryEntries = (IList<TeamFoundationJobHistoryEntry>)e.Result;
                    FilterEventJobHistoryEntries();
                }
            };
            worker.RunWorkerAsync();
        }

        private void FilterEventJobHistoryEntries()
        {
            this.EventJobHistoryEntries.Clear();
            foreach (var entry in this.allEventJobHistoryEntries.Where(e => !this.HideSucceededEventJobHistoryEntries || e.Result != TeamFoundationJobResult.Succeeded).OrderByDescending(e => e.EndTime))
            {
                this.EventJobHistoryEntries.Add(entry);
            }
            SetStatus(string.Format(CultureInfo.CurrentCulture, "Found {0} event job history item(s), showing {1}.", this.allEventJobHistoryEntries.Count, this.EventJobHistoryEntries.Count));
        }

        #endregion

        #region GetNotificationJobLogLevel Command

        private bool CanGetNotificationJobLogLevel(object argument)
        {
            return !string.IsNullOrEmpty(this.TeamProjectCollectionUrl);
        }

        private void GetNotificationJobLogLevel(object argument)
        {
            var worker = new BackgroundWorker();
            worker.DoWork += (sender, e) =>
            {
                SetStatus("Working...");
                using (var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(this.TeamProjectCollectionUrl)))
                {
                    var registry = tfs.GetService<ITeamFoundationRegistry>();
                    var logLevelValue = registry.GetValue("/Service/Integration/Settings/NotificationJobLogLevel");
                    var logLevelNumber = 0;
                    if (!string.IsNullOrEmpty(logLevelValue) && !int.TryParse(logLevelValue, out logLevelNumber))
                    {
                        throw new InvalidOperationException("The NotificationJobLogLevel registry value is unknown: " + logLevelValue);
                    }
                    e.Result = (NotificationJobLogLevel)logLevelNumber;
                }
            };
            worker.RunWorkerCompleted += (sender, e) =>
            {
                if (e.Error != null)
                {
                    HandleException(e.Error);
                }
                else
                {
                    this.NotificationJobLogLevel = (NotificationJobLogLevel)e.Result;
                    SetStatus(string.Format(CultureInfo.CurrentCulture, "The notification job log level is set to {0}.", this.NotificationJobLogLevel.ToString()));
                    CommandManager.InvalidateRequerySuggested();
                }
            };
            worker.RunWorkerAsync();
        }

        #endregion

        #region SetNotificationJobLogLevel Command

        private bool CanSetNotificationJobLogLevel(object argument)
        {
            return !string.IsNullOrEmpty(this.TeamProjectCollectionUrl) && this.NotificationJobLogLevel.HasValue;
        }

        private void SetNotificationJobLogLevel(object argument)
        {
            var worker = new BackgroundWorker();
            worker.DoWork += (sender, e) =>
            {
                SetStatus("Working...");
                var logLevel = this.NotificationJobLogLevel.Value;
                using (var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(this.TeamProjectCollectionUrl)))
                {
                    var registry = tfs.GetService<ITeamFoundationRegistry>();
                    var logLevelValue = ((int)logLevel).ToString();
                    registry.SetValue("/Service/Integration/Settings/NotificationJobLogLevel", logLevelValue);
                    SetStatus(string.Format(CultureInfo.CurrentCulture, "Changed the notification job log level to {0}.", logLevel));
                }
            };
            worker.RunWorkerCompleted += (sender, e) =>
            {
                if (e.Error != null)
                {
                    HandleException(e.Error);
                }
                else
                {
                    var result = e.Result;
                }
            };
            worker.RunWorkerAsync();
        }

        #endregion

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion

        #region Helper Methods

        private void SetStatus(string message, string detail = null, bool isError = false)
        {
            this.StatusText = message == null ? null : message.Replace(Environment.NewLine, " "); // Force single line
            this.StatusDetail = detail;
            this.StatusColorBackground = isError ? Brushes.Red : null;
            this.StatusColorForeground = isError ? Brushes.White : Brushes.Black;
        }

        private void HandleException(Exception exc)
        {
            SetStatus("An error occurred: " + exc.Message, exc.ToString(), true);
            MainWindow.ShowException(exc);
        }

        #endregion
    }
}
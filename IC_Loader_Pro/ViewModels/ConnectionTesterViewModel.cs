using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Threading.Tasks;
using System.Windows.Input;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.ViewModels
{
    public class ConnectionTesterViewModel : PropertyChangedBase
    {
        private string _instance = "sde:postgresql:dep-v-agis4";
        public string Instance { get => _instance; set => SetProperty(ref _instance, value); }

        private string _database = "srp";
        public string Database { get => _database; set => SetProperty(ref _database, value); }

        private string _user;
        public string User { get => _user; set => SetProperty(ref _user, value); }

        private string _password;
        public string Password { get => _password; set => SetProperty(ref _password, value); }

        private string _statusMessage = "Ready to test.";
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        private bool _isSuccess;
        public bool IsSuccess { get => _isSuccess; set => SetProperty(ref _isSuccess, value); }

        public ICommand TestConnectionCommand { get; }

        public ConnectionTesterViewModel()
        {
            TestConnectionCommand = new RelayCommand(async () => await OnTestConnectionAsync(), () => !string.IsNullOrEmpty(Instance));
        }

        private async Task OnTestConnectionAsync()
        {
            StatusMessage = "Testing...";
            IsSuccess = false;

            await QueuedTask.Run(() =>
            {
                Geodatabase gdb = null;
                try
                {
                    var dbConnectionProperties = new DatabaseConnectionProperties(EnterpriseDatabaseType.PostgreSQL)
                    {
                        AuthenticationMode = AuthenticationMode.DBMS,
                        Instance = this.Instance,
                        Database = this.Database,
                        User = this.User,
                        Password = this.Password
                    };

                    // The act of creating the Geodatabase object attempts the connection.
                    gdb = new Geodatabase(dbConnectionProperties);

                    // If we get here without an exception, the connection was successful.
                    StatusMessage = "Connection Successful!";
                    IsSuccess = true;
                }
                catch (Exception ex)
                {
                    // The most useful info is the specific error message from the database.
                    StatusMessage = $"Failed: {ex.Message}";
                    IsSuccess = false;
                }
                finally
                {
                    // Always dispose of the geodatabase object to close the connection.
                    gdb?.Dispose();
                }
            });
        }
    }
}
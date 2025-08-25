using IC_Loader_Pro.ViewModels;
using System.Windows;

namespace IC_Loader_Pro.Views
{
    public partial class ConnectionTesterWindow : Window
    {
        // Helper to easily access the ViewModel
        private ConnectionTesterViewModel ViewModel => DataContext as ConnectionTesterViewModel;

        public ConnectionTesterWindow()
        {
            InitializeComponent();
        }

        private void TestButton_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null)
            {
                // Manually pass the password from the PasswordBox to the ViewModel
                ViewModel.Password = PasswordBox.Password;
            }
        }
    }
}
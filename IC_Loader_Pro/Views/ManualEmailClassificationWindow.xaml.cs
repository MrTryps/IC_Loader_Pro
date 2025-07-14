using System.Windows;

namespace IC_Loader_Pro.Views
{
    public partial class ManualEmailClassificationWindow : Window
    {
        public ManualEmailClassificationWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // This sets the window's result to true so the calling code knows the user clicked "OK"
            DialogResult = true;
        }
    }
}
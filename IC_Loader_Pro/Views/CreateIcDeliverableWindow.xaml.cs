using System.Windows;

namespace IC_Loader_Pro.Views
{
    public partial class CreateIcDeliverableWindow : Window
    {
        public CreateIcDeliverableWindow()
        {
            InitializeComponent();
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // Set the DialogResult to true so the calling code knows the user clicked "Create"
            this.DialogResult = true;
        }
    }
}
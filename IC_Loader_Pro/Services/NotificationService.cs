using System.Threading.Tasks;
using static BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    public class NotificationService
    {
        public async Task SendConfirmationEmailAsync(string deliverableId)
        {
            Log.RecordMessage($"Building and sending confirmation email for {deliverableId} (shelled)...", BisLogMessageType.Note);
            // TODO: Implement logic to build and send email based on templates and test results.
            ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show("Pretend this is an outgoing email.","TODO");
            await Task.CompletedTask;
        }
    }
}
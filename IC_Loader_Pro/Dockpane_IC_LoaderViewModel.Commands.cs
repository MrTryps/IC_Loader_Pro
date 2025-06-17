using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel : DockPane
    {
        #region Commands
        public ICommand SaveCommand { get; private set; }
        public ICommand SkipCommand { get; private set; }
        public ICommand RejectCommand { get; private set; }
        public ICommand ShowNotesCommand { get; }
        public ICommand SearchCommand { get; }
        public ICommand ToolsCommand { get; }
        public ICommand OptionsCommand { get; }
        #endregion

        #region Command Methods
        private async Task OnSave()
        {
            Log.recordMessage("Save button was clicked.", Bis_Log_Message_Type.Note);

            // We can now use 'await' here for any async GIS work in the future,
            // for example, saving features to the 'manually_added' layer.
            await QueuedTask.Run(() => {
                // ... future GIS logic ...
            });
        }

        // Also update OnSkip and OnReject
        private async Task OnSkip()
        {
            Log.recordMessage("Skip button was clicked.", Bis_Log_Message_Type.Note);
            // Use Task.CompletedTask as a placeholder since there's no async work yet.
            await Task.CompletedTask;
        }

        private async Task OnReject()
        {
            Log.recordMessage("Reject button was clicked.", Bis_Log_Message_Type.Note);
            await Task.CompletedTask;
        }

        #endregion
    }
}

using IC_Loader_Pro.Models;
using IC_Loader_Pro.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro
{
    internal partial class Dockpane_IC_LoaderViewModel
    {
        #region Core Workflow Logic

        /// <summary>
        /// Fetches the list of emails for the selected queue and displays the first one.
        /// </summary>
        private async Task LoadEmailsForQueueAsync(ICQueueSummary queue)
        {
            StatusMessage = $"Fetching emails for queue: {queue.Name}...";
            IsUIEnabled = false;

            await Task.Run(() =>
            {
                try
                {
                    var outlookService = new OutlookService(); // From your library

                    // This call gets the detailed list for only the selected queue
                    _emailsForCurrentQueue = outlookService.GetEmailsFromSubfolder(queue.Name);

                    _currentEmailIndex = -1; // Reset the index for the new list
                }
                catch (Exception ex)
                {
                    Log.recordError($"Failed to load emails for queue '{queue.Name}'.", ex, nameof(LoadEmailsForQueueAsync));
                    _emailsForCurrentQueue?.Clear(); // Clear any partial results
                }
            });

            Log.recordMessage($"Found {_emailsForCurrentQueue?.Count ?? 0} emails for queue '{queue.Name}'.", Bis_Log_Message_Type.Note);

            // Display the first email
            ShowNextEmail();

            IsUIEnabled = true;
        }

        /// <summary>
        /// Advances to the next email in the current list and updates the CurrentEmail property.
        /// </summary>
        private void ShowNextEmail()
        {
            if (_emailsForCurrentQueue == null || _emailsForCurrentQueue.Count == 0)
            {
                CurrentEmail = null;
                StatusMessage = "No emails found in this queue.";
                return;
            }

            _currentEmailIndex++; // Move to the next index

            if (_currentEmailIndex < _emailsForCurrentQueue.Count)
            {
                // If we are still within the bounds of the list, set the CurrentEmail
                CurrentEmail = _emailsForCurrentQueue[_currentEmailIndex];
                StatusMessage = $"Reviewing: {CurrentEmail.Subject}";
            }
            else
            {
                // We have finished the list
                CurrentEmail = null;
                StatusMessage = $"All emails in queue '{SelectedIcType.Name}' have been processed.";
                Log.recordMessage("End of email queue reached.", Bis_Log_Message_Type.Note);
            }
        }

        #endregion
    }
}

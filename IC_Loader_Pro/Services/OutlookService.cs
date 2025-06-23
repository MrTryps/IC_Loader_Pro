using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace IC_Loader_Pro.Services
{
    internal class OutlookService
    {     

        /// <summary>
        /// Gets all emails from a specified subfolder of the Outlook Inbox, sorted by received date.
        /// </summary>
        /// <param name="targetFolderName">The name of the subfolder inside the Inbox to search.</param>
        /// <returns>A list of simple EmailItem objects.</returns>
        public List<EmailItem> GetEmailsFromSubfolder(string targetFolderName)
        {
            var results = new List<EmailItem>();
            Outlook.Application outlookApp = null;
            Outlook.MAPIFolder inboxFolder = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items outlookItems = null;

            try
            {
                outlookApp = new Outlook.Application();
                inboxFolder = outlookApp.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                // Find the specific subfolder by name
                targetFolder = inboxFolder.Folders[targetFolderName] as Outlook.MAPIFolder;
                if (targetFolder == null)
                {
                    throw new DirectoryNotFoundException($"The Outlook folder '{targetFolderName}' was not found in the Inbox.");
                }

                outlookItems = targetFolder.Items;
                // Sort the items by ReceivedTime, descending (newest first) for processing
                outlookItems.Sort("[ReceivedTime]", true);

                // This is the MAPI property for the permanent Internet Message ID
                const string MessageIdProp = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";

                foreach (object item in outlookItems)
                {
                    if (item is Outlook.MailItem mailItem)
                    {
                        try
                        {
                            string internetMessageId = mailItem.PropertyAccessor.GetProperty(MessageIdProp)?.ToString();
                            if (string.IsNullOrEmpty(internetMessageId)) continue; // Skip if no permanent ID

                            // Create our clean data object with data from the email
                            var emailData = new EmailItem
                            {
                                PermanentId = internetMessageId,
                                Subject = mailItem.Subject,
                                ReceivedTime = mailItem.ReceivedTime,
                                SenderName = mailItem.SenderName,
                                AttachmentCount = mailItem.Attachments.Count
                            };
                            results.Add(emailData);
                        }
                        finally
                        {
                            // CRITICAL: Release each MailItem COM object inside the loop
                            Marshal.ReleaseComObject(mailItem);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Here you can use your logger to record the exception before re-throwing it
                // Log.recordError("Failed to retrieve emails from Outlook.", ex, "GetEmailsFromSubfolder");
                throw;
            }
            finally
            {
                // CRITICAL: Clean up all other COM objects to ensure Outlook can close properly
                if (outlookItems != null) Marshal.ReleaseComObject(outlookItems);
                if (targetFolder != null) Marshal.ReleaseComObject(targetFolder);
                if (inboxFolder != null) Marshal.ReleaseComObject(inboxFolder);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }

            return results;
        }
    }
}

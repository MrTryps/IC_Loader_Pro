using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using static BIS_Tools_2025_Core.BIS_Log;
using static IC_Loader_Pro.Module1;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace IC_Loader_Pro.Services
{
    internal class OutlookService
    {
        /// <summary>
        /// Gets emails from a folder, applying a three-state test filter.
        /// </summary>
        /// <param name="fullFolderPath">The full path to the folder.</param>
        /// <param name="testSenderEmail">The email address to use for filtering.</param>
        /// <param name="isInTestMode">The flag controlling the filter mode (null, true, or false).</param>
        /// <returns>A list of EmailItem objects.</returns>
        public List<EmailItem> GetEmailsFromFolderPath(string fullFolderPath, string testSenderEmail, bool? isInTestMode)
        {
            if (string.IsNullOrWhiteSpace(fullFolderPath) || !fullFolderPath.StartsWith("\\\\"))
            {
                throw new ArgumentException("Invalid folder path format. Path must start with '\\\\'.", nameof(fullFolderPath));
            }

            var results = new List<EmailItem>();
            Outlook.Application outlookApp = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items outlookItems = null;

            try
            {
                var (storeName, folderPath) = ParseOutlookPath(fullFolderPath);
                outlookApp = new Outlook.Application();
                targetFolder = GetFolderFromPath(outlookApp.GetNamespace("MAPI"), storeName, folderPath);

                if (targetFolder == null)
                {
                    throw new System.IO.DirectoryNotFoundException($"The Outlook folder specified by the path '{fullFolderPath}' could not be found.");
                }

                outlookItems = targetFolder.Items;
                outlookItems.Sort("[ReceivedTime]", true);

                foreach (object item in outlookItems)
                {
                    if (item is Outlook.MailItem mailItem)
                    {
                        try
                        {
                            string senderEmail;
                            if (mailItem.SenderEmailType == "EX")
                            {
                                senderEmail = mailItem.Sender?.GetExchangeUser()?.PrimarySmtpAddress;
                            }
                            else
                            {
                                senderEmail = mailItem.SenderEmailAddress;
                            }

                            if (string.IsNullOrEmpty(senderEmail)) continue;

                            string internetMessageId = mailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")?.ToString();
                            if (string.IsNullOrEmpty(internetMessageId)) continue;

                            results.Add(new EmailItem
                            {
                                PermanentId = internetMessageId,
                                Subject = mailItem.Subject,
                                ReceivedTime = mailItem.ReceivedTime,
                                SenderName = mailItem.SenderName,
                                SenderEmailAddress = senderEmail,
                                AttachmentCount = mailItem.Attachments.Count
                            });
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mailItem);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.recordError($"Failed to retrieve emails using path '{fullFolderPath}'.", ex, "GetEmailsFromFolderPath");
                throw;
            }
            finally
            {
                if (outlookItems != null) Marshal.ReleaseComObject(outlookItems);
                if (targetFolder != null) Marshal.ReleaseComObject(targetFolder);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }

            // --- NEW THREE-STATE FILTERING LOGIC ---
            // If the flag is null or the test email is not set, do nothing.
            if (!isInTestMode.HasValue || string.IsNullOrWhiteSpace(testSenderEmail))
            {
                return results;
            }

            // If the flag is true, filter FOR the test sender.
            if (isInTestMode.Value)
            {
                Log.recordMessage($"TEST MODE (Include): Filtering for emails from {testSenderEmail}", Bis_Log_Message_Type.Warning);
                return results.Where(e => e.SenderEmailAddress.Equals(testSenderEmail, StringComparison.OrdinalIgnoreCase)).ToList();
            }
            // If the flag is false, filter OUT the test sender.
            else
            {
                Log.recordMessage($"TEST MODE (Exclude): Filtering out emails from {testSenderEmail}", Bis_Log_Message_Type.Warning);
                return results.Where(e => !e.SenderEmailAddress.Equals(testSenderEmail, StringComparison.OrdinalIgnoreCase)).ToList();
            }
        }

        /// <summary>
        /// Parses a path like "\\Store Name\Folder\Subfolder" into its components.
        /// </summary>
        private (string storeName, string folderPath) ParseOutlookPath(string fullPath)
        {
            var parts = fullPath.TrimStart('\\').Split(new[] { '\\' }, 2);
            if (parts.Length < 2)
            {
                throw new ArgumentException("Path must include at least a store name and a folder name.", nameof(fullPath));
            }
            return (parts[0], parts[1]);
        }

        /// <summary>
        /// Navigates to and retrieves a folder object based on the store and folder path.
        /// </summary>
        private Outlook.MAPIFolder GetFolderFromPath(Outlook.NameSpace mapiNamespace, string storeName, string folderPath)
        {
            Outlook.Store targetStore = null;
            Outlook.MAPIFolder currentFolder = null;

            try
            {
                targetStore = mapiNamespace.Stores
                    .Cast<Outlook.Store>()
                    .FirstOrDefault(s => s.DisplayName.Equals(storeName, StringComparison.OrdinalIgnoreCase));

                if (targetStore == null) return null;

                currentFolder = targetStore.GetRootFolder();
                var folderNames = folderPath.Split('\\');

                foreach (var name in folderNames)
                {
                    Outlook.MAPIFolder nextFolder = null;
                    try
                    {
                        nextFolder = currentFolder.Folders[name];
                        Marshal.ReleaseComObject(currentFolder); // Release previous folder
                        currentFolder = nextFolder;
                    }
                    catch
                    {
                        // Folder not found, clean up and return null
                        if (nextFolder != null) Marshal.ReleaseComObject(nextFolder);
                        Marshal.ReleaseComObject(currentFolder);
                        return null;
                    }
                }
                return currentFolder;
            }
            finally
            {
                if (targetStore != null) Marshal.ReleaseComObject(targetStore);
            }
        }
    }
}
using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using static BIS_Log;
using static IC_Loader_Pro.Models.EmailItem;
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
                                Emailid = internetMessageId,
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
                Log.RecordError($"Failed to retrieve emails using path '{fullFolderPath}'.", ex, "GetEmailsFromFolderPath");
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
                Log.RecordMessage($"TEST MODE (Include): Filtering for emails from {testSenderEmail}", BisLogMessageType.Warning);
                return results.Where(e => e.SenderEmailAddress.Equals(testSenderEmail, StringComparison.OrdinalIgnoreCase)).ToList();
            }
            // If the flag is false, filter OUT the test sender.
            else
            {
                Log.RecordMessage($"TEST MODE (Exclude): Filtering out emails from {testSenderEmail}", BisLogMessageType.Warning);
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

        /// <summary>
        /// Retrieves a fully populated EmailItem by its unique Internet Message ID.
        /// </summary>
        public EmailItem GetEmailById(string internetMessageId)
        {
            const string methodName = "GetEmailById";
            Outlook.Application outlookApp = null;
            Outlook.MailItem foundMail = null;
            string tempAttachmentPath = null;

            try
            {
                outlookApp = new Outlook.Application();

                // --- THE FIX ---
                // Sanitize the ID by removing the angle brackets before searching.
                string sanitizedId = internetMessageId.Trim('<', '>');

                foundMail = FindMailItemById(outlookApp, sanitizedId);

                if (foundMail != null)
                {
                    // We found it, now build the complete EmailItem
                    string senderEmail = GetSenderAddress(foundMail);
                    tempAttachmentPath = SaveAttachmentsToTempFolder(foundMail.Attachments);

                    var emailItem = new EmailItem
                    {
                        Emailid = internetMessageId, // Store the original, full ID
                        Subject = foundMail.Subject,
                        ReceivedTime = foundMail.ReceivedTime,
                        SenderName = foundMail.SenderName,
                        SenderEmailAddress = senderEmail,
                        Body = foundMail.Body
                    };

                    if (Directory.Exists(tempAttachmentPath))
                    {
                        emailItem.Attachments = Directory.GetFiles(tempAttachmentPath)
                            .Select(f => new AttachmentItem { FileName = Path.GetFileName(f), SavedPath = f })
                            .ToList();
                    }

                    return emailItem;
                }
            }
            catch (Exception ex)
            {
                Log.RecordError($"Error processing email with ID '{internetMessageId}'.", ex, methodName);
                throw;
            }
            finally
            {
                // Ensure all COM objects and temp files are cleaned up
                if (foundMail != null) Marshal.ReleaseComObject(foundMail);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
                if (Directory.Exists(tempAttachmentPath))
                {
                    try { Directory.Delete(tempAttachmentPath, true); }
                    catch (Exception ex) { Log.RecordError($"Failed to delete temp folder: {tempAttachmentPath}", ex, "GetEmailById_Finally"); }
                }
            }

            // If we get here, the email was not found
            throw new FileNotFoundException($"An email with the ID '{internetMessageId}' could not be found using AdvancedSearch.");
        }

        #region Private Helpers

        /// <summary>
        /// Finds a single email using the fast AdvancedSearch method, waiting for it to complete.
        /// </summary>
        private Outlook.MailItem FindMailItemById(Outlook.Application app, string sanitizedId)
        {
            const string MessageIdPropSchema = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
            string filter = $"{MessageIdPropSchema} = '{sanitizedId}'";

            var searchCompleteEvent = new AutoResetEvent(false);
            Outlook.Search advancedSearch = null;
            Outlook.MailItem result = null;

            // Define the event handler within the method's scope
            void SearchCompleteHandler(Outlook.Search Search)
            {
                advancedSearch = Search;
                searchCompleteEvent.Set(); // Signal that the search has finished
            }
            ;

            app.AdvancedSearchComplete += SearchCompleteHandler;

            try
            {
                app.AdvancedSearch("SCOPE_ALL_STORES", filter, false, "GetByIdSearchTag");

                // Wait for the event to be signaled, with a reasonable timeout
                if (searchCompleteEvent.WaitOne(15000) && advancedSearch != null && advancedSearch.Results.Count > 0)
                {
                    // The COM object is returned, caller is responsible for releasing it
                    result = advancedSearch.Results[1] as Outlook.MailItem;
                }
                else if (advancedSearch == null)
                {
                    Log.RecordError("Outlook advanced search timed out.", null, "FindMailItemById");
                }
            }
            finally
            {
                // Unsubscribe from the event to prevent memory leaks
                app.AdvancedSearchComplete -= SearchCompleteHandler;
                if (advancedSearch != null) Marshal.ReleaseComObject(advancedSearch);
                searchCompleteEvent.Close();
            }

            return result;
        }

        private string GetSenderAddress(Outlook.MailItem mailItem)
        {
            if (mailItem.SenderEmailType == "EX")
                return mailItem.Sender?.GetExchangeUser()?.PrimarySmtpAddress;

            return mailItem.SenderEmailAddress;
        }

        private string SaveAttachmentsToTempFolder(Outlook.Attachments attachments)
        {
            string tempFolderPath = Path.Combine(Path.GetTempPath(), "IC_Loader", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolderPath);

            if (attachments == null || attachments.Count == 0) return tempFolderPath;

            foreach (Outlook.Attachment attachment in attachments)
            {
                try
                {
                    string fullPath = Path.Combine(tempFolderPath, attachment.FileName);
                    int count = 1;
                    while (File.Exists(fullPath))
                    {
                        string newFileName = $"{Path.GetFileNameWithoutExtension(attachment.FileName)} ({count++}){Path.GetExtension(attachment.FileName)}";
                        fullPath = Path.Combine(tempFolderPath, newFileName);
                    }
                    attachment.SaveAsFile(fullPath);
                }
                catch (Exception ex)
                {
                    Log.RecordError($"Failed to save attachment '{attachment.FileName}'.", ex, "SaveAttachmentsToTempFolder");
                }
            }
            return tempFolderPath;
        }
        #endregion
    }
}
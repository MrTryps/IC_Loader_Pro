using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using static BIS_Log;
using static IC_Loader_Pro.Models.EmailItem;
using static IC_Loader_Pro.Module1;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace IC_Loader_Pro.Services
{
    internal class OutlookService
    {
        // The DASL property name for the Message-ID.
        //private const string _messageIdPropSchema = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        private const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";

        /// <summary>
        /// Gets emails from a folder, applying a three-state test filter.
        /// </summary>
        /// <param name="fullFolderPath">The full path to the folder.</param>
        /// <param name="testSenderEmail">The email address to use for filtering.</param>
        /// <param name="isInTestMode">The flag controlling the filter mode (null, true, or false).</param>
        /// <returns>A list of EmailItem objects.</returns>
        public List<EmailItem> GetEmailsFromFolderPath(Outlook.Application outlookApp, string fullFolderPath, string testSenderEmail, bool? isInTestMode)
        {
            if (!IsOutlookResponsive(outlookApp))
            {
                // If Outlook is not responsive, return an empty list immediately to prevent a freeze.
                throw new OutlookNotResponsiveException("Outlook is not running or is not responsive.");
            }

            if (string.IsNullOrWhiteSpace(fullFolderPath) || !fullFolderPath.StartsWith("\\\\"))
            {
                throw new ArgumentException("Invalid folder path format. Path must start with '\\\\'.", nameof(fullFolderPath));
            }

            var results = new List<EmailItem>();
           // Outlook.Application outlookApp = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items outlookItems = null;

            try
            {
                var (storeName, folderPath) = ParseOutlookPath(fullFolderPath);
                var mapiNamespace = outlookApp.GetNamespace("MAPI");
                targetFolder = GetFolderFromPath(mapiNamespace, storeName, folderPath);

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
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                throw;
            }
            finally
            {
                if (outlookItems != null) Marshal.ReleaseComObject(outlookItems);
                if (targetFolder != null) Marshal.ReleaseComObject(targetFolder);
               // if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }

            // First, filter out any emails that have been previously skipped in this session.
            var filteredResults = results
                .Where(e => !Module1.SkippedEmailIds.Contains(e.Emailid))
                .ToList();

            // Now, apply the test mode filtering to the already-filtered list.
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
        public static (string storeName, string folderPath) ParseOutlookPath(string fullPath)
        {
            var parts = fullPath.TrimStart('\\').Split(new[] { '\\' }, 2);
            if (parts.Length < 2)
            {
                throw new ArgumentException("Path must include at least a store name and a folder name.", nameof(fullPath));
            }
            return (parts[0], parts[1]);
        }

        public Outlook.MAPIFolder GetFolderFromPath(Outlook.NameSpace mapiNamespace, string storeName, string folderPath)
        {
            const string methodName = "GetFolderFromPath";
            Outlook.Store targetStore = null;
            Outlook.MAPIFolder parentFolder = null;
            Outlook.MAPIFolder childFolder = null;

            try
            {
                targetStore = mapiNamespace.Stores
                    .Cast<Outlook.Store>()
                    .FirstOrDefault(s => s.DisplayName.Equals(storeName, StringComparison.OrdinalIgnoreCase));

                if (targetStore == null)
                {
                    Log.RecordError($"[GetFolderFromPath] Could not find a store with DisplayName: '{storeName}'.", null, methodName);
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                    return null;
                }

                parentFolder = targetStore.GetRootFolder();
                var folderNames = folderPath.Split('\\');

                foreach (var name in folderNames)
                {
                    try
                    {
                        // Add a small delay to give Outlook time to populate the subfolders collection
                        System.Threading.Thread.Sleep(250);

                        childFolder = parentFolder.Folders[name];

                        if (parentFolder != targetStore.GetRootFolder())
                        {
                            Marshal.ReleaseComObject(parentFolder);
                        }
                        parentFolder = childFolder;
                    }
                    catch
                    {
                        Log.RecordError($"[GetFolderFromPath] CRITICAL FAILURE: Could not find the subfolder named '{name}' inside of '{parentFolder.Name}'.", null, methodName);
                        if (childFolder != null) Marshal.ReleaseComObject(childFolder);
                        if (parentFolder != null) Marshal.ReleaseComObject(parentFolder);
                        if (targetStore != null) Marshal.ReleaseComObject(targetStore);
                        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.","Processing Error");
                        return null;
                    }
                }
                return parentFolder;
            }
            catch (Exception ex)
            {
                Log.RecordError($"[GetFolderFromPath] An unexpected top-level exception occurred.", ex, methodName);
                return null;
            }
            finally
            {
                if (targetStore != null) Marshal.ReleaseComObject(targetStore);
            }
        }


        public Outlook.MAPIFolder GetFolderFromPath_old(Outlook.NameSpace mapiNamespace, string storeName, string folderPath)
        {
            const string methodName = "GetFolderFromPath";
            Log.RecordMessage($"[GetFolderFromPath] Attempting to find folder. Store: '{storeName}', Path: '{folderPath}'.", BisLogMessageType.Note);
            Outlook.Store targetStore = null;
            Outlook.MAPIFolder parentFolder = null;
            Outlook.MAPIFolder childFolder = null;

            try
            {
                targetStore = mapiNamespace.Stores
                    .Cast<Outlook.Store>()
                    .FirstOrDefault(s => s.DisplayName.Equals(storeName, StringComparison.OrdinalIgnoreCase));

                if (targetStore == null)
                {
                    Log.RecordError($"[GetFolderFromPath] Could not find a store with DisplayName: '{storeName}'.", null, methodName);
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                    return null;
                }

                parentFolder = targetStore.GetRootFolder();
                var folderNames = folderPath.Split('\\');

                // Traverse the path, carefully managing each folder object.
                Log.RecordMessage($"--- Searching inside '{parentFolder.Name}'. Found subfolders: ---", BisLogMessageType.Note);
                foreach (var name in folderNames)
                {
                    System.Threading.Thread.Sleep(250); // A quarter-second pause
                    // --- NEW DIAGNOSTIC LOGGING ---
                    // Log all available subfolders inside the current parent folder.
                    Log.RecordMessage($"--- Searching inside '{parentFolder.Name}'. Found subfolders: ---", BisLogMessageType.Note);
                    foreach (Outlook.MAPIFolder subfolder in parentFolder.Folders)
                    {
                        Log.RecordMessage($" -> '{subfolder.Name}'", BisLogMessageType.Note);
                        Marshal.ReleaseComObject(subfolder);
                    }
                    Log.RecordMessage("--------------------------------------------------", BisLogMessageType.Note);
                    // --- END OF DIAGNOSTIC LOGGING ---
                    try
                    {
                        childFolder = parentFolder.Folders[name];

                        // Release the previous parent folder before assigning the new one.
                        // We don't release the absolute root folder.
                        if (parentFolder != targetStore.GetRootFolder())
                        {
                            Marshal.ReleaseComObject(parentFolder);
                        }
                        parentFolder = childFolder;
                    }
                    catch
                    {
                        Log.RecordError($"[GetFolderFromPath] CRITICAL FAILURE: Could not find the subfolder named '{name}' inside of '{parentFolder.Name}'.", null, methodName);
                        // Clean up everything and exit immediately.
                        if (childFolder != null) Marshal.ReleaseComObject(childFolder);
                        if (parentFolder != null) Marshal.ReleaseComObject(parentFolder);
                        if (targetStore != null) Marshal.ReleaseComObject(targetStore);
                        ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                        return null;
                    }
                }
                // If we successfully looped through all parts, parentFolder now holds our target folder.
                return parentFolder;
            }
            catch (Exception ex)
            {
                Log.RecordError($"[GetFolderFromPath] An unexpected top-level exception occurred.", ex, methodName);
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                return null;
            }
            finally
            {
                // Final cleanup of the store object.
                if (targetStore != null) Marshal.ReleaseComObject(targetStore);
            }
        }

        /// <summary>
        /// Navigates to and retrieves a folder object based on the store and folder path.
        /// </summary>
        private Outlook.MAPIFolder GetFolderFromPath2(Outlook.NameSpace mapiNamespace, string storeName, string folderPath)
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
        /// Retrieves a specific email from Outlook and maps it to a custom EmailItem object.
        /// </summary>
        /// <param name="folderPath">The path to the folder to search in (e.g., "Inbox/My Project").</param>
        /// <param name="messageId">The Internet Message ID of the email to find.</param>
        /// <returns>A populated EmailItem object, or null if the email is not found.</returns>
        /// <summary>
        /// Retrieves a specific email from Outlook and maps it to a custom EmailItem object.
        /// </summary>
        public EmailItem GetEmailById(Outlook.Application outlookApp, string folderPath, string messageId, string storeName = null)
        {
            // --- Start of Diagnostic Logging ---
            Log.RecordMessage($"Attempting to get email. ID: '{messageId}', Folder Path: '{folderPath}', Store: '{storeName ?? "Default"}'.", BisLogMessageType.Note);
            // --- End of Diagnostic Logging ---

           // Outlook.Application outlookApp = null;
            Outlook.NameSpace mapiNamespace = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.MailItem mailItem = null;
            EmailItem result = null;

            // It's possible the message ID is missing the angle brackets, which are often required.
            if (!messageId.StartsWith("<")) messageId = "<" + messageId;
            if (!messageId.EndsWith(">")) messageId = messageId + ">";

            try
            {
                mapiNamespace = outlookApp.GetNamespace("MAPI");

                string actualStoreName = storeName;
                if (string.IsNullOrEmpty(actualStoreName))
                {
                    actualStoreName = mapiNamespace.DefaultStore.DisplayName;
                    // --- Start of Diagnostic Logging ---
                    Log.RecordMessage($"No store name provided. Defaulting to: '{actualStoreName}'.", BisLogMessageType.Note);
                    // --- End of Diagnostic Logging ---
                }

                targetFolder = this.GetFolderFromPath(mapiNamespace, actualStoreName, folderPath);

                if (targetFolder == null)
                {
                    // --- Start of Diagnostic Logging ---
                    Log.RecordError($"GetFolderFromPath returned null. Could not find folder '{folderPath}' in store '{actualStoreName}'.", null, nameof(GetEmailById));
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                    // --- End of Diagnostic Logging ---
                    return null;
                }

                // --- Start of Diagnostic Logging ---
                Log.RecordMessage($"Successfully found folder '{targetFolder.FolderPath}'. Now searching for email.", BisLogMessageType.Note);
                // --- End of Diagnostic Logging ---

                string filter = $"@SQL=\"{PR_INTERNET_MESSAGE_ID}\" = '{messageId}'";

                // --- Start of Diagnostic Logging ---
                Log.RecordMessage($"Using DASL filter: {filter}", BisLogMessageType.Note);
                // --- End of Diagnostic Logging ---

                object item = targetFolder.Items.Find(filter);

                if (item is Outlook.MailItem foundMailItem)
                {
                    mailItem = foundMailItem;
                    result = MapToEmailItem(mailItem);
                    // --- Start of Diagnostic Logging ---
                    Log.RecordMessage($"SUCCESS: Found email with subject: '{result.Subject}'.", BisLogMessageType.Note);
                    // --- End of Diagnostic Logging ---
                }
                else
                {
                    // --- Start of Diagnostic Logging ---
                    Log.RecordError($"The DASL query returned null. The email with ID '{messageId}' was not found in the folder '{targetFolder.FolderPath}'.", null, nameof(GetEmailById));
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                    // --- End of Diagnostic Logging ---
                }
            }
            catch (Exception ex)
            {
                Log.RecordError($"An exception occurred within GetEmailById.", ex, nameof(GetEmailById));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
            }
            finally
            {
                if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                if (targetFolder != null) Marshal.ReleaseComObject(targetFolder);
                if (mapiNamespace != null) Marshal.ReleaseComObject(mapiNamespace);
               // if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }

            return result;
        }
        #region Private Helpers
        /// <summary>
        /// Maps an Outlook.MailItem to the custom EmailItem model and saves attachments.
        /// </summary>
        /// <param name="mailItem">The source Outlook.MailItem.</param>
        /// <returns>A new, populated EmailItem object.</returns>
        // In Services/OutlookService.cs

        private EmailItem MapToEmailItem(Outlook.MailItem mailItem)
        {
            if (mailItem == null) return null;

            var emailItem = new EmailItem
            {
                Emailid = mailItem.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) as string,
                Subject = mailItem.Subject,
                ReceivedTime = mailItem.ReceivedTime,
                SenderName = mailItem.SenderName,
                SenderEmailAddress = mailItem.SenderEmailAddress,
                AttachmentCount = mailItem.Attachments.Count,
                Body = mailItem.Body
            };

            // Call our new helper method to handle saving all attachments.
            SaveAttachmentsToTempFolder(emailItem, mailItem.Attachments);

            return emailItem;
        }



        /// <summary>
        /// Finds an email by its ID in a source folder and moves it to a destination folder.
        /// </summary>
        /// <param name="messageId">The Internet Message ID of the email to move.</param>
        /// <param name="sourceFolderPath">The path of the folder where the email currently resides.</param>
        /// <param name="storeName">The name of the store (mailbox) for both source and destination.</param>
        /// <param name="destinationFolderPath">The path of the folder to move the email to.</param>
        /// <returns>True if the move was successful, otherwise false.</returns>
        public bool MoveEmailToFolder(Outlook.Application outlookApp, string messageId, string sourceFolderPath, string storeName, string destinationFolderPath)
        {
            Log.RecordMessage($"Attempting to move email '{messageId}' to folder '{destinationFolderPath}'.", BisLogMessageType.Note);

            //Outlook.Application outlookApp = null;
            Outlook.NameSpace mapiNamespace = null;
            Outlook.MAPIFolder sourceFolder = null;
            Outlook.MAPIFolder destinationFolder = null;
            object itemToMove = null;
            bool success = false;

            if (string.IsNullOrEmpty(messageId) || string.IsNullOrEmpty(sourceFolderPath) || string.IsNullOrEmpty(destinationFolderPath))
            {
                Log.RecordError("MoveEmailToFolder failed: One or more required parameters were null or empty.", null, nameof(MoveEmailToFolder));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                return false;
            }

            try
            {
                mapiNamespace = outlookApp.GetNamespace("MAPI");
                string actualStoreName = string.IsNullOrEmpty(storeName) ? mapiNamespace.DefaultStore.DisplayName : storeName;

                // Find both the source and destination folders
                sourceFolder = this.GetFolderFromPath(mapiNamespace, actualStoreName, sourceFolderPath);
                destinationFolder = this.GetFolderFromPath(mapiNamespace, actualStoreName, destinationFolderPath);

                if (sourceFolder == null)
                {
                    Log.RecordError($"Move failed: Could not find source folder '{sourceFolderPath}' in store '{actualStoreName}'.", null, nameof(MoveEmailToFolder));
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                    return false;
                }

                if (destinationFolder == null)
                {
                    Log.RecordError($"Move failed: Could not find destination folder '{destinationFolderPath}' in store '{actualStoreName}'.", null, nameof(MoveEmailToFolder));
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                    return false;
                }

                // Find the specific email item using the same DASL query logic
                string filter = $"@SQL=\"{PR_INTERNET_MESSAGE_ID}\" = '{messageId}'";
                itemToMove = sourceFolder.Items.Find(filter);

                if (itemToMove is Outlook.MailItem mailItem)
                {
                    mailItem.Move(destinationFolder);
                    success = true;
                    Log.RecordMessage($"Successfully moved email '{messageId}'.", BisLogMessageType.Note);
                }
                else
                {
                    Log.RecordError($"Move failed: Could not find email with ID '{messageId}' in source folder.", null, nameof(MoveEmailToFolder));
                }
            }
            catch (Exception ex)
            {
                Log.RecordError($"An exception occurred while trying to move email '{messageId}'.", ex, nameof(MoveEmailToFolder));
                ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                success = false;
            }
            finally
            {
                // Release all COM objects
                if (itemToMove != null) Marshal.ReleaseComObject(itemToMove);
                if (destinationFolder != null) Marshal.ReleaseComObject(destinationFolder);
                if (sourceFolder != null) Marshal.ReleaseComObject(sourceFolder);
                if (mapiNamespace != null) Marshal.ReleaseComObject(mapiNamespace);
               // if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }

            return success;
        }               

        private void SaveAttachmentsToTempFolder(EmailItem emailItem, Outlook.Attachments attachments)
        {
            if (attachments == null || attachments.Count == 0) return;

            // Create a unique temporary folder for this email's attachments.
            string tempFolderPath = Path.Combine(Path.GetTempPath(), "IC_Loader", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolderPath);

            // Store the path so we can access it later for processing and cleanup.
            emailItem.TempFolderPath = tempFolderPath;

            foreach (Outlook.Attachment attachment in attachments)
            {
                try
                {
                    // Note: We can add the sanitize logic from BisFileTools here if needed,
                    // or just save with the original name for now.
                    string savedPath = Path.Combine(tempFolderPath, attachment.FileName);
                    attachment.SaveAsFile(savedPath);

                    // Add the saved attachment info to our custom EmailItem.
                    emailItem.Attachments.Add(new EmailItem.AttachmentItem
                    {
                        FileName = attachment.FileName,
                        SavedPath = savedPath
                    });
                }
                catch (Exception ex)
                {
                    Log.RecordError($"Failed to save attachment '{attachment.FileName}'.", ex, "SaveAttachmentsToTempFolder");
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                }
                finally
                {
                    // Release the individual attachment COM object.
                    Marshal.ReleaseComObject(attachment);
                }
            }
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
                    ArcGIS.Desktop.Framework.Dialogs.MessageBox.Show($"An unexpected error occurred while processing this email. The application will advance to the next email.", "Processing Error");
                }
            }
            return tempFolderPath;
        }


        /// <summary>
        /// Checks if the main Outlook window is responsive by sending it a message with a timeout.
        /// </summary>
        /// <returns>True if Outlook is running and responsive, otherwise false.</returns>
        public bool IsOutlookResponsive(Outlook.Application outlookApp)
        {
            // The class name for the main Outlook window is "rctrl_renwnd32"
            IntPtr outlookHandle = Win32Helper.FindWindow("rctrl_renwnd32", null);

            if (outlookHandle == IntPtr.Zero)
            {
                Log.RecordMessage("Outlook process not found.", BisLogMessageType.Warning);
                return false; // Outlook is not running.
            }

            IntPtr result;
            const uint timeoutMilliseconds = 2000; // 2-second timeout

            // Send a null message to the Outlook window. If it's responsive, it will reply quickly.
            // If it's hung, this call will time out.
            IntPtr response = Win32Helper.SendMessageTimeout(
                outlookHandle,
                Win32Helper.WM_NULL,
                IntPtr.Zero,
                IntPtr.Zero,
                Win32Helper.SendMessageTimeoutFlags.SMTO_ABORTIFHUNG | Win32Helper.SendMessageTimeoutFlags.SMTO_NORMAL,
                timeoutMilliseconds,
                out result
            );

            // A non-zero response indicates success. A zero response indicates a timeout or error.
            if (response != IntPtr.Zero)
            {
                return true; // Outlook is responsive.
            }
            else
            {
                Log.RecordError("Outlook appears to be running but is not responsive.", null, nameof(IsOutlookResponsive));
                return false; // Outlook is hung.
            }
        }
        #endregion


        /// <summary>
        /// Attempts to forcefully restart the Microsoft Outlook application after getting user confirmation.
        /// </summary>
        /// <returns>True if the restart was attempted, false if the user canceled.</returns>
        public static async Task<bool> TryRestartOutlook()
        {
            // 1. Ask the user for permission and warn them about data loss.
            var message = "The application has detected that Microsoft Outlook is not responding.\n\n" +
                          "Would you like to try and forcefully restart it?\n\n" +
                          "WARNING: This will close Outlook without saving any unsaved work (such as draft emails).";

            var result = MessageBox.Show(message, "Outlook Unresponsive", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                // User clicked "No" or closed the dialog.
                return false;
            }

            try
            {
                // 2. Find all running Outlook processes.
                Process[] outlookProcesses = Process.GetProcessesByName("outlook");
                if (!outlookProcesses.Any())
                {
                    // Outlook is not running, so we can just try to start it.
                    Process.Start("outlook.exe");
                    return true;
                }

                // 3. Forcefully terminate (kill) each Outlook process.
                foreach (var process in outlookProcesses)
                {
                    process.Kill();
                }

                // Give the OS a moment to release the processes.
                await Task.Delay(2000); // 2-second delay

                // 4. Start a new instance of Outlook.
                Process.Start("outlook.exe");

                return true;
            }
            catch (Exception ex)
            {
                Log.RecordError("Failed to restart Outlook.", ex, "TryRestartOutlook");
                MessageBox.Show($"An error occurred while trying to restart Outlook: {ex.Message}", "Restart Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

    }
}
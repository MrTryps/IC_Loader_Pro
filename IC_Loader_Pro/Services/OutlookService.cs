using IC_Loader_Pro.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using static BIS_Log;
using static IC_Loader_Pro.Models.EmailItem;
using static IC_Loader_Pro.Module1;
using Outlook = Microsoft.Office.Interop.Outlook;

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
        public List<EmailItem> GetEmailsFromFolderPath(Outlook.Application outlookApp,string fullFolderPath, string testSenderEmail, bool? isInTestMode)
        {
            if (!IsOutlookResponsive())
            {
                // If Outlook is not responsive, return an empty list immediately to prevent a freeze.
                throw new OutlookNotResponsiveException("Outlook is not running or is not responsive.");
            }

            if (string.IsNullOrWhiteSpace(fullFolderPath) || !fullFolderPath.StartsWith("\\\\"))
            {
                throw new ArgumentException("Invalid folder path format. path must start with '\\\\'.", nameof(fullFolderPath));
            }

            var results = new List<EmailItem>();
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items outlookItems = null;

            try
            {
                var (storeName, folderPath) = ParseOutlookPath(fullFolderPath);
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
        public static (string storeName, string folderPath) ParseOutlookPath(string fullPath)
        {
            var parts = fullPath.TrimStart('\\').Split(new[] { '\\' }, 2);
            if (parts.Length < 2)
            {
                throw new ArgumentException("path must include at least a store name and a folder name.", nameof(fullPath));
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
                    .FirstOrDefault(s =>
                s != null && // First, check if the store object itself is not null
                s.DisplayName.Equals(storeName, StringComparison.OrdinalIgnoreCase));

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
        // In Services/OutlookService.cs

        /// <summary>
        /// Retrieves a specific email from Outlook and maps it to a custom EmailItem object.
        /// This version uses a shared Outlook Application instance.
        /// </summary>
        public EmailItem GetEmailById(Outlook.Application outlookApp, string folderPath, string messageId, string storeName = null)
        {
            Log.RecordMessage($"Attempting to get email. ID: '{messageId}', Folder path: '{folderPath}', Store: '{storeName ?? "Default"}'.", BisLogMessageType.Note);

            Outlook.NameSpace mapiNamespace = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.MailItem mailItem = null;
            EmailItem result = null;

            if (!messageId.StartsWith("<")) messageId = "<" + messageId;
            if (!messageId.EndsWith(">")) messageId = messageId + ">";

            try
            {
                // Use the passed-in outlookApp instance instead of creating a new one.
                mapiNamespace = outlookApp.GetNamespace("MAPI");

                string actualStoreName = storeName;
                if (string.IsNullOrEmpty(actualStoreName))
                {
                    actualStoreName = mapiNamespace.DefaultStore.DisplayName;
                }

                targetFolder = this.GetFolderFromPath(mapiNamespace, actualStoreName, folderPath);

                if (targetFolder == null)
                {
                    Log.RecordError($"GetFolderFromPath returned null for folder '{folderPath}' in store '{actualStoreName}'.", null, nameof(GetEmailById));
                    return null;
                }

                string filter = $"@SQL=\"{PR_INTERNET_MESSAGE_ID}\" = '{messageId}'";
                object item = targetFolder.Items.Find(filter);

                if (item is Outlook.MailItem foundMailItem)
                {
                    mailItem = foundMailItem;
                    result = MapToEmailItem(mailItem);
                }
                else
                {
                    Log.RecordError($"The DASL query returned null. Email ID '{messageId}' not found in folder '{targetFolder.FolderPath}'.", null, nameof(GetEmailById));
                }
            }
            catch (Exception ex)
            {
                Log.RecordError($"An exception occurred within GetEmailById.", ex, nameof(GetEmailById));
            }
            finally
            {
                // Release only the objects created within this method.
                // The outlookApp object is managed by the calling method.
                if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                if (targetFolder != null) Marshal.ReleaseComObject(targetFolder);
                if (mapiNamespace != null) Marshal.ReleaseComObject(mapiNamespace);
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

            Outlook.NameSpace mapiNamespace = null;
            Outlook.MAPIFolder sourceFolder = null;
            Outlook.MAPIFolder destinationFolder = null;
            object itemToMove = null;
            bool success = false;

            if (string.IsNullOrEmpty(messageId) || string.IsNullOrEmpty(sourceFolderPath) || string.IsNullOrEmpty(destinationFolderPath))
            {
                Log.RecordError("MoveEmailToFolder failed: One or more required parameters were null or empty.", null, nameof(MoveEmailToFolder));
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
                    return false;
                }

                if (destinationFolder == null)
                {
                    Log.RecordError($"Move failed: Could not find destination folder '{destinationFolderPath}' in store '{actualStoreName}'.", null, nameof(MoveEmailToFolder));
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
                success = false;
            }
            finally
            {
                // Release all COM objects
                if (itemToMove != null) Marshal.ReleaseComObject(itemToMove);
                if (destinationFolder != null) Marshal.ReleaseComObject(destinationFolder);
                if (sourceFolder != null) Marshal.ReleaseComObject(sourceFolder);
                if (mapiNamespace != null) Marshal.ReleaseComObject(mapiNamespace);
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
                }
            }
            return tempFolderPath;
        }


        /// <summary>
        /// Checks if the main Outlook window is responsive by sending it a message with a timeout.
        /// </summary>
        /// <returns>True if Outlook is running and responsive, otherwise false.</returns>
        public bool IsOutlookResponsive()
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

    }
}
using System.Collections.Generic;
using System.Linq;

namespace IC_Loader_Pro.Models
{
    public class OutgoingEmail
    {
        public List<string> ToRecipients { get; } = new List<string>();
        public List<string> CcRecipients { get; } = new List<string>();
        public List<string> BccRecipients { get; } = new List<string>();
        public string Subject { get; set; }
        public List<string> Attachments { get; } = new List<string>();

        public List<string> OpeningText { get; } = new List<string>();
        public List<string> MainBodyText { get; } = new List<string>();
        public List<string> ClosingText { get; } = new List<string>();

        /// <summary>
        /// A read-only property that combines the different parts of the email
        /// into a single HTML-formatted string.
        /// </summary>
        public string Body
        {
            get
            {
                var allText = new List<string>();
                allText.AddRange(OpeningText);
                allText.AddRange(MainBodyText);
                allText.AddRange(ClosingText);

                // Join with <br> for simple HTML line breaks
                return string.Join("<br>", allText.Where(s => !string.IsNullOrEmpty(s)));
            }
        }

        public void AddToMainBody(string textToAdd, int position = -1)
        {
            if (string.IsNullOrWhiteSpace(textToAdd)) textToAdd = "<BR>";
            if (position < 0 || position >= MainBodyText.Count) MainBodyText.Add(textToAdd);
            else MainBodyText.Insert(position, textToAdd);
        }


        /// <summary>
        /// Adds a line of text to the opening section of the email body.
        /// </summary>
        /// <param name="textToAdd">The text to add.</param>
        /// <param name="position">Optional. The zero-based index at which to insert the text. Defaults to the end.</param>
        public void AddToOpeningText(string textToAdd, int position = -1)
        {
            if (string.IsNullOrWhiteSpace(textToAdd))
            {
                textToAdd = "<BR>";
            }

            if (position < 0 || position >= OpeningText.Count)
            {
                OpeningText.Add(textToAdd);
            }
            else
            {
                OpeningText.Insert(position, textToAdd);
            }
        }

        /// <summary>
        /// Adds a line of text to the closing section of the email body.
        /// </summary>
        /// <param name="textToAdd">The text to add.</param>
        /// <param name="position">Optional. The zero-based index at which to insert the text. Defaults to the end.</param>
        public void AddToClosingText(string textToAdd, int position = -1)
        {
            if (string.IsNullOrWhiteSpace(textToAdd))
            {
                textToAdd = "<BR>";
            }

            if (position < 0 || position >= ClosingText.Count)
            {
                ClosingText.Add(textToAdd);
            }
            else
            {
                ClosingText.Insert(position, textToAdd);
            }
        }
        // --- END OF NEW METHODS ---
    }
}
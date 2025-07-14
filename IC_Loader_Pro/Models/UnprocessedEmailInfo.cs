namespace IC_Loader_Pro.Models
{
    /// <summary>
    /// A simple data class to hold information about an email that was
    /// not fully processed and requires manual user attention.
    /// </summary>
    public class UnprocessedEmailInfo
    {
        public string Subject { get; set; }
        public string Reason { get; set; } // e.g., "Moved to DNA queue", "User canceled", "Move to Junk failed"
    }
}
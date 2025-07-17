using System;

namespace IC_Loader_Pro.Services
{
    /// <summary>
    /// An exception that is thrown when the Outlook application is not running or is not responsive.
    /// </summary>
    public class OutlookNotResponsiveException : Exception
    {
        public OutlookNotResponsiveException(string message) : base(message) { }
    }
}
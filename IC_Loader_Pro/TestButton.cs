using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using System;
using System.Windows;

namespace IC_Loader_Pro
{
    internal class TestButton : Button
    {
        protected override void OnClick()
        {
            try
            {
                // This is the only line we care about.
                // We are just trying to get the Type information from the problematic library.
                Type rulesType = typeof(IC_Rules_2025.IC_Rules);

                // If the line above succeeds, we will see this message.
                MessageBox.Show($"SUCCESS! The type information for IC_Rules was accessed.\n\nAssembly loaded from:\n{rulesType.Assembly.Location}",
                                "Success!");
            }
            catch (Exception ex)
            {
                // If we get here, this proves the assembly cannot be loaded correctly.
                // This message box will contain the true root cause.
                MessageBox.Show($"This confirms the runtime error. The system failed to load the IC_Rules type.\n\nERROR:\n{ex.ToString()}",
                                "Runtime Load Failure");
            }
        }
    }
}
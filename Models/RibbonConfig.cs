namespace RibbonXWizard.Models
{
    /// <summary>
    /// Configuration collected from user input for ribbon customization
    /// </summary>
    public class RibbonConfig
    {
        /// <summary>
        /// The built-in Office ribbon tab where the button will appear (e.g., "TabHome")
        /// </summary>
        public string TargetTabId { get; set; } = "TabHome";

        /// <summary>
        /// Display name for the custom group
        /// </summary>
        public string GroupLabel { get; set; } = "My Tools";

        /// <summary>
        /// Text displayed on the button
        /// </summary>
        public string ButtonLabel { get; set; } = "My Button";

        /// <summary>
        /// Brief description shown on hover
        /// </summary>
        public string ScreenTip { get; set; } = "";

        /// <summary>
        /// Built-in Office icon identifier (e.g., "HappyFace")
        /// </summary>
        public string ImageMso { get; set; } = "MacroPlay";

        /// <summary>
        /// Name of the VBA macro to call (must exist in the .dotm)
        /// </summary>
        public string MacroName { get; set; } = "";

        /// <summary>
        /// Path to the source .dotm file
        /// </summary>
        public string SourceDotmPath { get; set; } = "";

        /// <summary>
        /// Generate a unique button ID based on the label
        /// </summary>
        public string GetButtonId()
        {
            // Remove spaces and special characters for a valid XML ID
            var cleaned = System.Text.RegularExpressions.Regex.Replace(
                ButtonLabel, @"[^a-zA-Z0-9]", "");
            return "btn" + cleaned;
        }

        /// <summary>
        /// Generate the VBA callback function name
        /// </summary>
        public string GetCallbackName()
        {
            var cleaned = System.Text.RegularExpressions.Regex.Replace(
                ButtonLabel, @"[^a-zA-Z0-9]", "");
            return "On" + cleaned + "Click";
        }
    }
}
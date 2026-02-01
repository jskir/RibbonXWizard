namespace RibbonXWizard.Models
{
    /// <summary>
    /// Contains all generated outputs from the wizard
    /// </summary>
    public class GeneratedRibbon
    {
        /// <summary>
        /// The CustomUI XML that will be embedded in the Office file
        /// </summary>
        public string XmlContent { get; set; } = "";

        /// <summary>
        /// VBA callback code that the user needs to add to their document
        /// </summary>
        public string VbaCallbackCode { get; set; } = "";

        /// <summary>
        /// User-friendly instructions for completing the setup
        /// </summary>
        public string Instructions { get; set; } = "";

        /// <summary>
        /// Path to the generated output .dotm file
        /// </summary>
        public string OutputFilePath { get; set; } = "";
    }
}
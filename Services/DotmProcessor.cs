using System;
using System.IO;
using OfficeRibbonXEditor.Common;

namespace RibbonXWizard.Services
{
    /// <summary>
    /// Handles manipulation of .dotm files to inject CustomUI XML
    /// </summary>
    public class DotmProcessor
    {
        /// <summary>
        /// Apply ribbon customization to a .dotm file
        /// </summary>
        /// <param name="sourcePath">Path to the original .dotm file</param>
        /// <param name="outputPath">Path where the modified file will be saved</param>
        /// <param name="customUiXml">The CustomUI XML content to inject</param>
        public void ApplyCustomization(string sourcePath, string outputPath, string customUiXml)
        {
            // Validation
            if (string.IsNullOrWhiteSpace(sourcePath))
                throw new ArgumentException("Source file path is required", nameof(sourcePath));

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output file path is required", nameof(outputPath));

            if (string.IsNullOrWhiteSpace(customUiXml))
                throw new ArgumentException("Custom UI XML is required", nameof(customUiXml));

            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Ensure the source file is actually a valid Office document
            if (!IsValidOfficeFile(sourcePath))
                throw new InvalidOperationException("The source file is not a valid Office document (.dotm)");

            try
            {
                // Open the Office document
                using (var doc = new OfficeDocument(sourcePath))
                {
                    // Office 2010+ uses CustomUI14 (Office 2007 uses CustomUI12)
                    // We'll use CustomUI14 as it's compatible with all modern Office versions
                    var partType = XmlPart.RibbonX14;

                    // Check if a CustomUI part already exists
                    var existingPart = doc.RetrieveCustomPart(partType);

                    if (existingPart != null)
                    {
                        // Replace the existing customization
                        existingPart.Save(customUiXml);
                    }
                    else
                    {
                        // Create a new CustomUI part
                        var newPart = doc.CreateCustomPart(partType);
                        newPart.Save(customUiXml);
                    }

                    // Save the modified document to the output path
                    doc.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to process the .dotm file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Verify that the file is a valid Office Open XML document
        /// </summary>
        private bool IsValidOfficeFile(string filePath)
        {
            var extension = Path.GetExtension(filePath)?.ToLowerInvariant();

            // Check for valid Office template extensions
            return extension == ".dotm" ||
                   extension == ".docm" ||
                   extension == ".dotx" ||
                   extension == ".docx";
        }

        /// <summary>
        /// Check if a .dotm file already has ribbon customizations
        /// </summary>
        public bool HasExistingCustomization(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                using (var doc = new OfficeDocument(filePath))
                {
                    return doc.HasCustomUi;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
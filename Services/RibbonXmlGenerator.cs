using System;
using System.Text;
using RibbonXWizard.Models;

namespace RibbonXWizard.Services
{
    /// <summary>
    /// Generates CustomUI XML for Office ribbon customization
    /// </summary>
    public class RibbonXmlGenerator
    {
        /// <summary>
        /// Generate a complete CustomUI XML based on configuration
        /// </summary>
        public GeneratedRibbon Generate(RibbonConfig config)
        {
            if (config == null)
                throw new ArgumentNullException(nameof(config));

            var result = new GeneratedRibbon();

            // Generate the CustomUI XML
            result.XmlContent = GenerateCustomUiXml(config);

            // Generate VBA callback code
            result.VbaCallbackCode = GenerateVbaCallback(config);

            // Generate user instructions
            result.Instructions = GenerateInstructions(config);

            return result;
        }

        /// <summary>
        /// Creates the CustomUI14.xml content
        /// </summary>
        private string GenerateCustomUiXml(RibbonConfig config)
        {
            var xml = new StringBuilder();

            xml.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            xml.AppendLine(@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">");
            xml.AppendLine("  <ribbon>");
            xml.AppendLine("    <tabs>");
            xml.AppendLine($"      <tab idMso=\"{config.TargetTabId}\">");
            xml.AppendLine($"        <group id=\"grp{config.GetButtonId()}\" label=\"{EscapeXml(config.GroupLabel)}\">");

            xml.Append($"          <button id=\"{config.GetButtonId()}\" ");
            xml.Append($"label=\"{EscapeXml(config.ButtonLabel)}\" ");
            xml.Append($"size=\"large\" ");
            xml.Append($"imageMso=\"{config.ImageMso}\" ");

            if (!string.IsNullOrWhiteSpace(config.ScreenTip))
            {
                xml.Append($"screentip=\"{EscapeXml(config.ScreenTip)}\" ");
            }

            xml.AppendLine($"onAction=\"{config.GetCallbackName()}\" />");

            xml.AppendLine("        </group>");
            xml.AppendLine("      </tab>");
            xml.AppendLine("    </tabs>");
            xml.AppendLine("  </ribbon>");
            xml.AppendLine("</customUI>");

            return xml.ToString();
        }

        /// <summary>
        /// Generate VBA callback code for the user to paste
        /// </summary>
        private string GenerateVbaCallback(RibbonConfig config)
        {
            var vba = new StringBuilder();

            vba.AppendLine("' ============================================");
            vba.AppendLine("' RIBBON CALLBACK CODE");
            vba.AppendLine("' Paste this into a standard module in your .dotm file");
            vba.AppendLine("' ============================================");
            vba.AppendLine();
            vba.AppendLine($"' Callback for button: {config.ButtonLabel}");
            vba.AppendLine($"Sub {config.GetCallbackName()}(control As IRibbonControl)");
            vba.AppendLine($"    ' This callback is triggered when the ribbon button is clicked");
            vba.AppendLine($"    ' It calls your existing macro: {config.MacroName}");
            vba.AppendLine($"    Call {config.MacroName}");
            vba.AppendLine("End Sub");

            return vba.ToString();
        }

        /// <summary>
        /// Generate step-by-step instructions for the user
        /// </summary>
        private string GenerateInstructions(RibbonConfig config)
        {
            var instructions = new StringBuilder();

            instructions.AppendLine("RIBBON CUSTOMIZATION INSTRUCTIONS");
            instructions.AppendLine("=" + new string('=', 50));
            instructions.AppendLine();
            instructions.AppendLine("Your customized .dotm file has been created!");
            instructions.AppendLine();
            instructions.AppendLine("NEXT STEPS:");
            instructions.AppendLine();
            instructions.AppendLine("1. Open the generated .dotm file in Microsoft Word");
            instructions.AppendLine("   (You may need to click 'Enable Content' if prompted)");
            instructions.AppendLine();
            instructions.AppendLine("2. Press Alt+F11 to open the VBA Editor");
            instructions.AppendLine();
            instructions.AppendLine("3. In the VBA Editor:");
            instructions.AppendLine("   - If you don't already have a standard module, create one:");
            instructions.AppendLine("     Insert → Module");
            instructions.AppendLine("   - Paste the VBA callback code (provided below)");
            instructions.AppendLine();
            instructions.AppendLine("4. Make sure your original macro is present:");
            instructions.AppendLine($"   - The macro '{config.MacroName}' must exist in this document");
            instructions.AppendLine();
            instructions.AppendLine("5. Save the file (Ctrl+S)");
            instructions.AppendLine();
            instructions.AppendLine("6. Close and reopen Word");
            instructions.AppendLine();
            instructions.AppendLine($"7. Look for your button '{config.ButtonLabel}' on the {GetTabDisplayName(config.TargetTabId)} tab");
            instructions.AppendLine();
            instructions.AppendLine("DISTRIBUTION:");
            instructions.AppendLine("To share with colleagues:");
            instructions.AppendLine("- They can place this .dotm in their Word Startup folder:");
            instructions.AppendLine("  %APPDATA%\\Microsoft\\Word\\STARTUP\\");
            instructions.AppendLine("- The button will appear automatically when they start Word");
            instructions.AppendLine();

            return instructions.ToString();
        }

        /// <summary>
        /// Escape XML special characters
        /// </summary>
        private string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        /// <summary>
        /// Get friendly tab name from idMso
        /// </summary>
        private string GetTabDisplayName(string tabIdMso)
        {
            return tabIdMso switch
            {
                "TabHome" => "Home",
                "TabInsert" => "Insert",
                "TabDesign" => "Design",
                "TabPageLayoutWord" => "Layout",
                "TabReferences" => "References",
                "TabMailings" => "Mailings",
                "TabReview" => "Review",
                "TabView" => "View",
                "TabDeveloper" => "Developer",
                "TabAddIns" => "Add-ins",
                _ => tabIdMso
            };
        }
    }
}
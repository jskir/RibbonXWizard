using System.Collections.Generic;

namespace RibbonXWizard.Data
{
    /// <summary>
    /// Provides data about built-in Office ribbon tabs
    /// </summary>
    public static class BuiltInTabs
    {
        /// <summary>
        /// Common Word ribbon tabs with their idMso values
        /// </summary>
        public static List<TabInfo> GetWordTabs()
        {
            return new List<TabInfo>
            {
                new TabInfo { DisplayName = "Home", IdMso = "TabHome" },
                new TabInfo { DisplayName = "Insert", IdMso = "TabInsert" },
                new TabInfo { DisplayName = "Design", IdMso = "TabDesign" },
                new TabInfo { DisplayName = "Layout", IdMso = "TabPageLayoutWord" },
                new TabInfo { DisplayName = "References", IdMso = "TabReferences" },
                new TabInfo { DisplayName = "Mailings", IdMso = "TabMailings" },
                new TabInfo { DisplayName = "Review", IdMso = "TabReview" },
                new TabInfo { DisplayName = "View", IdMso = "TabView" },
                new TabInfo { DisplayName = "Developer", IdMso = "TabDeveloper" },
                new TabInfo { DisplayName = "Add-ins", IdMso = "TabAddIns" }
            };
        }

        public class TabInfo
        {
            public string DisplayName { get; set; } = "";
            public string IdMso { get; set; } = "";

            public override string ToString() => DisplayName;
        }
    }
}
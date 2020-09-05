namespace PowerPointLabs.TextCollection
{
    internal static class LiveCodingLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "LiveCodingLabMenu";
        public const string HighlightDifferenceTag = "HighlightDifference";
        public const string AnimateNewLinesTag = "AnimateNewLines";
        public const string AnimateScrollDownTag = "AnimateScrollDown";
        public const string AnimateScrollUpTag = "AnimateScrollUp";
        public const string SettingsTag = "LiveCodingLabSettings";
        #endregion

        #region GUI Text
        public const string RibbonMenuLabel = "Live Coding";
        public const string HighlightDifferenceButtonLabel = "Highlight Difference";
        public const string AnimateNewLinesButtonLabel = "Animate New Lines";
        public const string AnimateScrollDownButtonLabel = "Animate Scroll Down";
        public const string AnimateScrollUpButtonLabel = "Animate Scroll Up";
        public const string SettingsButtonLabel = "Settings";

        public const string RibbonMenuSupertip =
            "Use Live Coding Lab to create live coding simulations for your slides.\n\n" +
            "Click this button to open the Live Coding Lab interface.";
        public const string HighlightDifferenceButtonSupertip = 
            "Highlight any changes made between two pieces of code by changing the font colour.\n\n" +
            "To perform this action, select the two text boxes containing the two code extracts, then click this button.";
        public const string AnimateNewLinesButtonSupertip = 
            "Creates an animation to show the transition from one piece of code to the other piece of code.\n\n" +
            "To perform this action, select the two text boxes containing the two code extracts, then click this button.";
        public const string AnimateScrollDownButtonSupertip = 
            "Creates an animation to simulate a scroll down of a piece of code.\n\n" +
            "To perform this action, select the two text boxes containing the two code extracts, then click this button.";
        public const string AnimateScrollUpButtonSupertip =
            "Creates an animation to simulate a scroll up of a piece of code.\n\n" +
            "To perform this action, select the two text boxes containing the two code extracts, then click this button.";
        public const string SettingsButtonSupertip = "Configure the settings for Live Coding Lab.";

        public const string SettingsScrollSpeedTooltip = "The duration (in seconds) for the animation to scroll down to the desired portion in the code.";

        public const string ErrorHighlightDifferenceDialogTitle = "Unable to execute action";
        public const string ErrorHighlightDifferenceWrongSlide = "Please select the correct slide.";
        public const string ErrorHighlightDifferenceNoSelection = "Please select the code snippet.";
        public const string ErrorHighlightDifferenceCodeSnippet = "Missing or extra code snippet.";
        #endregion
    }
}

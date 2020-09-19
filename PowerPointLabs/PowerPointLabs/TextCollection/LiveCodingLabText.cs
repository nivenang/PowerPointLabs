namespace PowerPointLabs.TextCollection
{
    internal static class LiveCodingLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "LiveCodingLabMenu";
        public const string HighlightDifferenceTag = "HighlightDifference";
        public const string AnimateNewLinesTag = "AnimateNewLines";
        public const string LiveCodingLabPaneTag = "LiveCodingLab";
        public const string AnimateScrollDownTag = "AnimateScrollDown";
        public const string AnimateScrollUpTag = "AnimateScrollUp";
        public const string SettingsTag = "LiveCodingLabSettings";
        public const string TaskPanelTitle = "Live Coding Lab";
        #endregion

        #region Identifiers
        public const string SlideNumIdentifier = "SlideNum";
        public const string CodeTextIdentifier = "CodeText";
        public const string CodeURLIdentifier = "CodeURL";
        public const string IsCodeURLIdentifier = "IsCodeURL";
        public const string TagNoIdentifier = "TagNo";
        public const string CodeBoxShapeNameFormat = Identifier + Underscore + "{0}" + Underscore + CodeBoxIdentifier;
        #endregion

        #region Code Box Identifiers
        public const string CodeBox_IsURL = "IsURL";
        public const string CodeBox_IsFile = "IsFile";
        public const string CodeBox_IsText = "IsText";
        public const string CodeBox_URLText = "URLText";
        public const string CodeBox_FileText = "FileText";
        public const string CodeBox_UserText = "UserText";
        public const string CodeBox_Id = "CodeBoxId";
        public const string CodeBox_Slide = "Slide";
        public const string CodeBox_CodeShape = "CodeShape";
        #endregion

        #region Storage Identifiers
        public const string Identifier = "PPTL";
        public const string Underscore = "_";
        public const string TextStorageIdentifier = "Storage";
        public const string CodeBoxItemIdentifier = "Item";
        public const string CodeBoxIdentifier = "CodeBox";
        public const string CodeBoxStorageIdentifier = "CodeBoxStorage";
        public const string LiveCodingLabTextStorageShapeName = Identifier + Underscore + TextStorageIdentifier;
        public const string ExtractTagNoRegex = Identifier + Underscore + @"([1-9][0-9]*)" +
            Underscore + "(" + CodeBoxIdentifier + @").*";
        public const string CodeBoxShapeNameRegex = Identifier + Underscore + @"[1-9][0-9]*" + Underscore + CodeBoxIdentifier;
        #endregion

        #region GUI Text

        public const string NoSlideSelectedMessage = "No slide is selected";
        public const string OnLoadingMessage = "Now Loading...";

        public const string RibbonMenuLabel = "Live Coding";
        public const string HighlightDifferenceButtonLabel = "Highlight Difference";
        public const string AnimateNewLinesButtonLabel = "Animate New Lines";
        public const string LiveCodingLabPaneButtonLabel = "Live Coding Lab Pane";
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
        public const string LiveCodingLabPaneButtonSupertip =
            "Opens the Live Coding Lab Pane\n\n" +
            "To perform this action, click this button.";
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

        public const string ErrorInvalidFileName = "The file specified does not exist.";
        public const string ErrorInvalidSelection = "You need to select {1} {2} before applying '{0}'.";
        public const string ErrorValueLessThanEqualsZero = "Please enter a value greater than 0.";
        public const string ErrorValueLessThanEqualsZeroWithShape = "Please enter a value greater than 0 (Shape {0}).";
        public const string ErrorUndefined = "Undefined error in Live Coding Lab.";
        #endregion
    }
}

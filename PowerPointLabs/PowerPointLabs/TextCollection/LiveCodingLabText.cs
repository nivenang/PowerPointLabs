namespace PowerPointLabs.TextCollection
{
    internal static class LiveCodingLabText
    {
        #region Action Framework Tags
        public const string RibbonMenuId = "LiveCodingLabButton";
        public const string HighlightDifferenceTag = "HighlightDifference";
        public const string AnimateNewLinesTag = "AnimateNewLines";
        public const string LiveCodingLabPaneTag = "LiveCodingLab";
        public const string SettingsTag = "LiveCodingLabSettings";
        public const string TaskPanelTitle = "Live Coding Lab";
        public const string PaneTag = "LiveCodingLab";
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
        public const string CodeBox_IsFile = "IsFile";
        public const string CodeBox_IsText = "IsText";
        public const string CodeBox_IsDiff = "IsDiff";
        public const string CodeBox_FileText = "FileText";
        public const string CodeBox_UserText = "UserText";
        public const string CodeBox_DiffText = "DiffText";
        public const string CodeBox_DiffIndex = "DiffIndex";
        public const string CodeBox_Id = "CodeBoxId";
        public const string CodeBox_Slide = "Slide";
        public const string CodeBox_CodeShape = "CodeShape";
        public const string CodeBox_Group = "Group";
        public const string CodeBox_ShapeName = "ShapeName";
        #endregion

        #region Storage Identifiers
        public const string Identifier = "PPTL";
        public const string Underscore = "_";
        public const string TextStorageIdentifier = "LiveCodingStorage";
        public const string CodeBoxItemIdentifier = "Item";
        public const string CodeBoxIdentifier = "CodeBox";
        public const string CodeBoxStorageIdentifier = "CodeBoxStorage";
        public const string LiveCodingLabTextStorageShapeName = Identifier + Underscore + TextStorageIdentifier;
        public const string ExtractTagNoRegex = Identifier + Underscore + @"([1-9][0-9]*)" +
            Underscore + "(" + CodeBoxIdentifier + @").*";
        public const string CodeBoxShapeNameRegex = Identifier + Underscore + @"[1-9][0-9]*" + Underscore + CodeBoxIdentifier;
        public const string TransitionSlideIdentifier = Identifier + Underscore + "LiveCodingTransitionSlide";
        public const string TransitionTextIdentifier = "PPTLabsTransitionText";
        #endregion

        #region Animation Identifiers
        public const string AnimateLineDiffIdentifier = "AnimateLineDiff_";
        public const string AnimateBlockDiffIdentifier = "AnimateBlockDiff_";
        public const string AnimateWordDiffIdentifier = "AnimateWordDiff_";
        public const string AnimateCharDiffIdentifier = "AnimateCharDiff_";
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

        public const string PromptToRecreateCodeBox = "Live Coding Lab has detected that this Code Box does not exist in the slides.\n" +
            "Do you want to recreate the Code Box in the current slide?";

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

        public const string ErrorHighlightDifferenceDialogTitle = "Unable to execute Highlight Difference action";
        public const string ErrorHighlightDifferenceWrongSlide = "Please select the correct slide.";
        public const string ErrorHighlightDifferenceNoSelection = "Please select the code snippet.";
        public const string ErrorHighlightDifferenceMissingCodeSnippet = "Missing code snippet.";
        public const string ErrorHighlightDifferenceWrongCodeSnippet = "Mismatched code snippets. Please ensure that code snippets have the same number of lines.";

        public const string ErrorAnimateNewLinesDialogTitle = "Unable to execute Animate New Lines action";
        public const string ErrorAnimateNewLinesWrongSlide = "Please select the correct slide.";
        public const string ErrorAnimateNewLinesMissingCodeSnippet = "Missing code snippet.";
        public const string ErrorAnimateNewLinesWrongCodeSnippet = "Mismatched code snippets. Please ensure that the 'after' code snippet have more lines than the 'before' code snippet.";

        public const string ErrorAnimateLineDiffDialogTitle = "Unable to execute Animate Line Diff action";
        public const string ErrorAnimateBlockDiffDialogTitle = "Unable to execute Animate Block Diff action";
        public const string ErrorAnimateWordDiffDialogTitle = "Unable to execute Animate Word Diff action";
        public const string ErrorAnimateCharDiffDialogTitle = "Unable to execute Animate Char Diff action";
        public const string ErrorAnimateDiffWrongSlide = "Please select the correct slide.";
        public const string ErrorAnimateDiffMissingCodeSnippet = "Missing code snippet. Please ensure that there is a 'before' and 'after' code snippet and that your current slide selected is on the 'before' code snippet.";
        public const string ErrorAnimateDiffWrongCodeSnippet = "Mismatched code snippets. Please ensure that the 'before' and 'after' code snippets have the same diff file.";

        public const string ErrorInvalidFileName = "The file specified does not exist.";
        public const string ErrorInvalidCodeBox = "You need to have {1} {2} on both the current slide and next slide before applying '{0}'.";
        public const string ErrorValueLessThanEqualsZero = "Please enter a value greater than 0.";
        public const string ErrorValueLessThanEqualsZeroWithShape = "Please enter a value greater than 0 (Shape {0}).";
        public const string ErrorUndefined = "Undefined error in Live Coding Lab.";
        #endregion
    }
}

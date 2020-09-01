using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.LiveCodingLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportActionRibbonId(LiveCodingLabText.HighlightDifferenceTag)]
    class HighlightDifferenceActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            if (this.GetAddIn().Application.ActiveWindow.Selection.ShapeRange.Count == 2)
            {
                HighlightDifference.HighlightDifferences();
            }
        }
    }
}

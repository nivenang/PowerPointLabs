using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.LiveCodingLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportActionRibbonId(LiveCodingLabText.LiveCodingLabPaneTag)]
    class LiveCodingLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(LiveCodingLabPane), LiveCodingLabText.TaskPanelTitle);
            CustomTaskPane liveCodingPane = this.GetTaskPane(typeof(LiveCodingLabPane));
            // if currently the pane is hidden, show the pane
            if (!liveCodingPane.Visible)
            {
                // fire the pane visble change event
                liveCodingPane.Visible = true;
            }
            else
            {
                liveCodingPane.Visible = false;
            }
        }
    }
}

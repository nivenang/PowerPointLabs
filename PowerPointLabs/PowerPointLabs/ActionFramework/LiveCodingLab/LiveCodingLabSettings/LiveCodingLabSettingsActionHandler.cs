using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.LiveCodingLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportActionRibbonId(LiveCodingLabText.SettingsTag)]
    class LiveCodingLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            LiveCodingLabSettings.ShowAnimationSettingsDialog();
        }
    }
}

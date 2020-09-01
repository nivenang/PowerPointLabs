using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Label.LiveCodingLab
{
    [ExportLabelRibbonId(LiveCodingLabText.SettingsTag)]
    class LiveCodingLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return LiveCodingLabText.SettingsButtonLabel;
        }
    }
}

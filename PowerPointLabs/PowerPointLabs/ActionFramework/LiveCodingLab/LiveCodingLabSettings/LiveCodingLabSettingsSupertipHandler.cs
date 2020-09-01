using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Supertip.LiveCodingLab
{
    [ExportSupertipRibbonId(LiveCodingLabText.SettingsTag)]
    class LiveCodingLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return LiveCodingLabText.SettingsButtonSupertip;
        }
    }
}

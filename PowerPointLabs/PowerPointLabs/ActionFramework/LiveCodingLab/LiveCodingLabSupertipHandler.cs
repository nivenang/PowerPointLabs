using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportSupertipRibbonId(LiveCodingLabText.LiveCodingLabPaneTag)]
    class LiveCodingLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return LiveCodingLabText.RibbonMenuSupertip;
        }
    }
}

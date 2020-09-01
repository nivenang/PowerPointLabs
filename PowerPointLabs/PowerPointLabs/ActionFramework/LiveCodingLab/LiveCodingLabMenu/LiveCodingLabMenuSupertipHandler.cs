using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Supertip.LiveCodingLab
{
    [ExportSupertipRibbonId(LiveCodingLabText.RibbonMenuId)]
    class LiveCodingLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return LiveCodingLabText.RibbonMenuSupertip;
        }
    }
}

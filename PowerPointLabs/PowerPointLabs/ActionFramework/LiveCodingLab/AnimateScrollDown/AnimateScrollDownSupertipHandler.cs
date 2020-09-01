using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportSupertipRibbonId(LiveCodingLabText.AnimateScrollDownTag)]
    class AnimateScrollDownSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return LiveCodingLabText.AnimateScrollDownButtonSupertip;
        }
    }
}

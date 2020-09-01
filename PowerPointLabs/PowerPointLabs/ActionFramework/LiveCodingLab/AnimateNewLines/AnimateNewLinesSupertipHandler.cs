using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportSupertipRibbonId(LiveCodingLabText.AnimateNewLinesTag)]
    class AnimateNewLinesSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return LiveCodingLabText.AnimateNewLinesButtonSupertip;
        }
    }
}

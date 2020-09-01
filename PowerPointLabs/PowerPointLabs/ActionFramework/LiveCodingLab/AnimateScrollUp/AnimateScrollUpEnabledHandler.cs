using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportEnabledRibbonId(LiveCodingLabText.AnimateScrollUpTag)]
    class AnimateScrollDownUpHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            
        }
    }
}

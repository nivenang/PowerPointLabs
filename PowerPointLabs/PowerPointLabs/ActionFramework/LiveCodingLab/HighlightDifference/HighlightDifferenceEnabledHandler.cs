using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.LiveCodingLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.LiveCodingLab
{
    [ExportEnabledRibbonId(LiveCodingLabText.HighlightDifferenceTag)]
    class HighlightDifferenceHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return HighlightDifference.IsHighlightDifferenceEnabled;
        }
    }
}

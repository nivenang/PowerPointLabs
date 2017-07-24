﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportLabelRibbonId(AgendaLabText.UpdateAgendaTag)]
    class UpdateAgendaLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return AgendaLabText.UpdateAgendaButtonLabel;
        }
    }
}

﻿using System.Collections.Generic;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportContentRibbonId(
        ShortcutsLabText.MenuShape, ShortcutsLabText.MenuLine, ShortcutsLabText.MenuFreeform,
        ShortcutsLabText.MenuPicture, ShortcutsLabText.MenuSlide, ShortcutsLabText.MenuGroup,
        ShortcutsLabText.MenuInk, ShortcutsLabText.MenuVideo, ShortcutsLabText.MenuTextEdit,
        ShortcutsLabText.MenuChart, ShortcutsLabText.MenuTable, ShortcutsLabText.MenuTableCell,
        ShortcutsLabText.MenuSmartArt, ShortcutsLabText.MenuEditSmartArt, ShortcutsLabText.MenuEditSmartArtText)]

    class ContextMenuContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            List<ContextMenuGroup> contextMenuGroups = GetContextMenuGroups(ribbonId);
            StringBuilder xmlString = new System.Text.StringBuilder();

            foreach (ContextMenuGroup group in contextMenuGroups)
            {
                string id = group.Name.Replace(" ", "");
                xmlString.Append(string.Format(CommonText.DynamicMenuXmlTitleMenuSeparator,
                    id + ShortcutsLabText.MenuSeparator, group.Name));

                foreach (string groupItem in group.Items)
                {
                    xmlString.Append(string.Format(CommonText.DynamicMenuXmlImageButton, groupItem + ribbonId, groupItem));
                }
            }

            return string.Format(CommonText.DynamicMenuXmlMenu, xmlString);
        }

        private List<ContextMenuGroup> GetContextMenuGroups(string ribbonId)
        {
            List<ContextMenuGroup> contextMenuGroups = new List<ContextMenuGroup>();
            ContextMenuGroup pasteLab = new ContextMenuGroup(ShortcutsLabText.PasteMenuLabel, new List<string>());
            ContextMenuGroup shortcuts = new ContextMenuGroup(ShortcutsLabText.ShortcutsMenuLabel, new List<string>());
            contextMenuGroups.Add(pasteLab);

            // All context menus will have these buttons
            pasteLab.Items.Add(PasteLabText.PasteAtCursorPositionTag);
            pasteLab.Items.Add(PasteLabText.PasteAtOriginalPositionTag);
            pasteLab.Items.Add(PasteLabText.PasteToFillSlideTag);
            pasteLab.Items.Add(PasteLabText.PasteToFitSlideTag);


            // Context menus other than slide will have these buttons
            if (ribbonId != ShortcutsLabText.MenuSlide)
            {
                // We only add shortcuts group if context menu is not for slide
                contextMenuGroups.Add(shortcuts);

                if (!ribbonId.Contains(ShortcutsLabText.MenuEditSmartArtBase) &&
                    ribbonId != ShortcutsLabText.MenuTextEdit &&
                    ribbonId != ShortcutsLabText.MenuTable)
                {
                    pasteLab.Items.Add(PasteLabText.ReplaceWithClipboardTag);
                }

                shortcuts.Items.Add(ShortcutsLabText.EditNameTag);

                // Context menus other than picture will have these buttons
                if (ribbonId != ShortcutsLabText.MenuPicture)
                {
                    shortcuts.Items.Add(ShortcutsLabText.ConvertToPictureTag);
                }

                shortcuts.Items.Add(ShortcutsLabText.HideShapeTag);
                shortcuts.Items.Add(ShortcutsLabText.AddCustomShapeTag);

                // Context menu group will have these buttons
                if (ribbonId == ShortcutsLabText.MenuGroup)
                {
                    pasteLab.Items.Add(PasteLabText.PasteIntoGroupTag);
                    shortcuts.Items.Add(ShortcutsLabText.AddIntoGroupTag);
                }
            }
            return contextMenuGroups;
        }

        public class ContextMenuGroup
        {
            public string Name { get; private set; }
            public List<string> Items { get; private set; }

            public ContextMenuGroup(string groupName, List<string> groupItems)
            {
                Name = groupName;
                Items = groupItems;
            }
        }
    }
}

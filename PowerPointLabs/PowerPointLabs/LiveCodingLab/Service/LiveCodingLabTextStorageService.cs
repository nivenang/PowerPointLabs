using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Utility;
using PowerPointLabs.LiveCodingLab.Views;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.LiveCodingLab.Service
{
    public class LiveCodingLabTextStorageService
    {
        public static void StoreCodeBoxToSlide(List<CodeBoxPaneItem> codeBoxItems,
            PowerPointSlide slide)
        {
            string shapeName = LiveCodingLabText.LiveCodingLabTextStorageShapeName;
            slide.DeleteShapeWithName(shapeName);
            List<Dictionary<string, string>> codeBoxDict =
                ConvertListToDictionary(codeBoxItems);
            XElement textInxml = new XElement(LiveCodingLabText.CodeBoxStorageIdentifier,
                codeBoxDict.Select(kv =>
                new XElement(LiveCodingLabText.CodeBoxItemIdentifier,
               from text in kv select new XElement(text.Key, text.Value))));
            Shape shape = ShapeUtility.InsertStorageCodeBoxToSlide(slide, shapeName, textInxml.ToString());
        }

        public static List<Dictionary<string, string>> LoadCodeBoxesFromSlide(PowerPointSlide slide)
        {
            List<Shape> shapes = slide.GetShapeWithName(LiveCodingLabText.LiveCodingLabTextStorageShapeName);
            if (shapes.Count > 0)
            {
                Shape shape = shapes[0];
                return LoadCodeBoxFromString(shape.TextFrame.TextRange.Text);
            }
            return null;
        }

        private static List<Dictionary<string, string>> LoadCodeBoxFromString(string text)
        {
            List<Dictionary<string, string>> codeBoxDict =
                new List<Dictionary<string, string>>();
            XElement xml = XElement.Parse(text);
            foreach (var codeBox in xml.Elements())
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add(LiveCodingLabText.CodeTextIdentifier, codeBox.Element(LiveCodingLabText.CodeTextIdentifier).Value);
                dic.Add(LiveCodingLabText.CodeBox_IsURL, codeBox.Element(LiveCodingLabText.CodeBox_IsURL).Value);
                dic.Add(LiveCodingLabText.CodeBox_IsFile, codeBox.Element(LiveCodingLabText.CodeBox_IsFile).Value);
                dic.Add(LiveCodingLabText.CodeBox_IsText, codeBox.Element(LiveCodingLabText.CodeBox_IsText).Value);
                dic.Add(LiveCodingLabText.CodeBox_Id, codeBox.Element(LiveCodingLabText.CodeBox_Id).Value);
                dic.Add(LiveCodingLabText.CodeBox_Slide, codeBox.Element(LiveCodingLabText.CodeBox_Slide).Value);
                codeBoxDict.Add(dic);
            }
            return codeBoxDict;
        }

        private static List<Dictionary<string, string>> ConvertListToDictionary(List<CodeBoxPaneItem> codeBoxItems)
        {
            List<Dictionary<string, string>> keyValuePairs =
                new List<Dictionary<string, string>>();
            foreach (CodeBoxPaneItem paneItem in codeBoxItems)
            {
                CodeBox item = paneItem.CodeBox;
                Dictionary<string, string> value = new Dictionary<string, string>();
                value.Add(LiveCodingLabText.CodeBox_IsURL, item.IsURL ? "Y" : "N");
                value.Add(LiveCodingLabText.CodeBox_IsFile, item.IsFile ? "Y" : "N");
                value.Add(LiveCodingLabText.CodeBox_IsText, item.IsText ? "Y" : "N");
                value.Add(LiveCodingLabText.CodeBox_Id, item.Id.ToString());
                value.Add(LiveCodingLabText.CodeTextIdentifier, item.Text);
                if (item.Slide == null)
                {
                    value.Add(LiveCodingLabText.CodeBox_Slide, "-1");
                }
                else
                {
                    value.Add(LiveCodingLabText.CodeBox_Slide, item.Slide.Index.ToString());
                }
                keyValuePairs.Add(value);
            }
            return keyValuePairs;
        }
    }
}

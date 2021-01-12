using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        /// <summary>
        /// Converts all code boxes into string format and stores it in the slides for storage purposes.
        /// </summary>
        /// <param name="codeBoxItems">list of all created code boxes</param>
        /// <param name="slide">slide to insert the storage code box into</param>
        public static void StoreCodeBoxToSlide(ObservableCollection<CodeBoxPaneItem> codeBoxItems,
            PowerPointSlide slide)
        {
            // Remove pre-existing storage code box in the slide
            string shapeName = LiveCodingLabText.LiveCodingLabTextStorageShapeName;
            slide.DeleteShapeWithName(shapeName);

            // Convert updated code boxes into storage format and store in the slide
            List<Dictionary<string, string>> codeBoxDict =
                ConvertListToDictionary(codeBoxItems);
            XElement textInxml = new XElement(LiveCodingLabText.CodeBoxStorageIdentifier,
                codeBoxDict.Select(kv =>
                new XElement(LiveCodingLabText.CodeBoxItemIdentifier,
               from text in kv select new XElement(text.Key, text.Value))));
            Shape shape = ShapeUtility.InsertStorageCodeBoxToSlide(slide, shapeName, textInxml.ToString());
        }

        /// <summary>
        /// Loads all stored code boxes upon opening Live Coding Lab
        /// </summary>
        /// <param name="slide">slide with the storage code box</param>
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

        /// <summary>
        /// Converts the stored string format of code boxes to dictionary format for loading
        /// </summary>
        /// <param name="text">stored string format of code boxes</param>
        /// <returns>list of dictionaries with each dictionary representing one code box</returns>
        private static List<Dictionary<string, string>> LoadCodeBoxFromString(string text)
        {
            List<Dictionary<string, string>> codeBoxDict =
                new List<Dictionary<string, string>>();
            XElement xml = XElement.Parse(text);
            foreach (var codeBox in xml.Elements())
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add(LiveCodingLabText.CodeTextIdentifier, codeBox.Element(LiveCodingLabText.CodeTextIdentifier).Value);
                dic.Add(LiveCodingLabText.CodeBox_IsFile, codeBox.Element(LiveCodingLabText.CodeBox_IsFile).Value);
                dic.Add(LiveCodingLabText.CodeBox_IsText, codeBox.Element(LiveCodingLabText.CodeBox_IsText).Value);
                dic.Add(LiveCodingLabText.CodeBox_IsDiff, codeBox.Element(LiveCodingLabText.CodeBox_IsDiff).Value);
                dic.Add(LiveCodingLabText.CodeBox_Id, codeBox.Element(LiveCodingLabText.CodeBox_Id).Value);
                dic.Add(LiveCodingLabText.CodeBox_Group, codeBox.Element(LiveCodingLabText.CodeBox_Group).Value);
                dic.Add(LiveCodingLabText.CodeBox_ShapeName, codeBox.Element(LiveCodingLabText.CodeBox_ShapeName).Value);
                dic.Add(LiveCodingLabText.CodeBox_DiffIndex, codeBox.Element(LiveCodingLabText.CodeBox_DiffIndex).Value);
                codeBoxDict.Add(dic);
            }
            return codeBoxDict;
        }

        /// <summary>
        /// Converts a list of code boxes into dictionary format for storage.
        /// </summary>
        /// <param name="codeBoxItems">list of code box items in the presentation</param>
        /// <returns>list of dictionaries with each dictionary representing one code box</returns>
        private static List<Dictionary<string, string>> ConvertListToDictionary(ObservableCollection<CodeBoxPaneItem> codeBoxItems)
        {
            List<Dictionary<string, string>> keyValuePairs =
                new List<Dictionary<string, string>>();
            foreach (CodeBoxPaneItem paneItem in codeBoxItems)
            {
                CodeBox item = paneItem.CodeBox;
                Dictionary<string, string> value = new Dictionary<string, string>();
                value.Add(LiveCodingLabText.CodeBox_IsFile, item.IsFile ? "Y" : "N");
                value.Add(LiveCodingLabText.CodeBox_IsText, item.IsText ? "Y" : "N");
                value.Add(LiveCodingLabText.CodeBox_IsDiff, item.IsDiff ? "Y" : "N");
                value.Add(LiveCodingLabText.CodeBox_Id, item.Id.ToString());
                value.Add(LiveCodingLabText.CodeTextIdentifier, item.Text);
                value.Add(LiveCodingLabText.CodeBox_Group, paneItem.Group);
                value.Add(LiveCodingLabText.CodeBox_ShapeName, item.ShapeName);
                value.Add(LiveCodingLabText.CodeBox_DiffIndex, item.DiffIndex.ToString());
                keyValuePairs.Add(value);
            }
            return keyValuePairs;
        }
    }
}
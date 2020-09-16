using System;
using System.Drawing;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.LiveCodingLab.Views
{
#pragma warning disable 0618
    /// <summary>
    /// Interaction logic for SyncFormatPaneItem.xaml
    /// </summary>
    public partial class CodeBoxPaneItem : UserControl
    {

        private LiveCodingPaneWPF parent;
        private CodeBox codeBox;

        #region Constructors

        public CodeBoxPaneItem(LiveCodingPaneWPF parent)
        {
            InitializeComponent();
            this.parent = parent;
            codeBox = new CodeBox(CodeBoxIdService.GenerateUniqueId());
            insertCode.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabEditButton);
            deleteImage.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabDeleteButton);
            isText.IsChecked = true;
            isURL.IsChecked = false;
            isFile.IsChecked = false;

        }

        public CodeBoxPaneItem(LiveCodingPaneWPF parent, CodeBox codeBox)
        {
            InitializeComponent();
            this.parent = parent;
            this.codeBox = codeBox;
            insertCode.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabEditButton);
            deleteImage.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabDeleteButton);
            if (codeBox.IsURL)
            {
                isText.IsChecked = false;
                isURL.IsChecked = true;
                isFile.IsChecked = false;
            }
            else if (codeBox.IsFile)
            {
                isText.IsChecked = false;
                isURL.IsChecked = false;
                isFile.IsChecked = true;
            }
            else
            {
                isText.IsChecked = true;
                isURL.IsChecked = false;
                isFile.IsChecked = false;
            }
        }
        #endregion

        #region Getters and Setters
        public string Text
        {
            get
            {
                return codeBox.Text;
            }
            set
            {
                codeBox.Text = value;
            }
        }

        public CodeBox CodeBox
        {
            get
            {
                return codeBox;
            }
        }

        #endregion

        #region Helper Functions
        public void PopulateTextBox()
        {
            codeTextBox.Text = codeBox.Text;
        }

        #endregion

        #region Event Handlers

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            parent.RemoveCodeBox(this);
            if (codeBox != null)
            {
                codeBox.Delete();
            }
            parent.SaveCodeBox();
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            codeBox.Text = codeTextBox.Text;
            if (codeBox.Shape == null)
            {
                codeBox = ShapeUtility.InsertCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox);
            }
            else
            {
                codeBox = ShapeUtility.ReplaceTextForShape(codeBox);
            }

            parent.SaveCodeBox();
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (codeBox.Slide != null)
            {
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(codeBox.Slide.Index);
            }
        }

        private void IsURL_Checked(object sender, RoutedEventArgs e)
        {
            codeBox.IsURL = true;
            codeBox.IsText = false;
            codeBox.IsFile = false;
        }

        private void IsFile_Checked(object sender, RoutedEventArgs e)
        {
            codeBox.IsURL = false;
            codeBox.IsText = false;
            codeBox.IsFile = true;
        }
        private void IsText_Checked(object sender, RoutedEventArgs e)
        {
            codeBox.IsURL = false;
            codeBox.IsText = true;
            codeBox.IsFile = false;
        }
        #endregion
    }
}

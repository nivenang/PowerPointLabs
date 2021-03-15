using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
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
    /// Interaction logic for CodeBoxPaneItem.xaml
    /// </summary>
    public partial class CodeBoxPaneItem : UserControl, INotifyPropertyChanged
    {

        private LiveCodingPaneWPF parent;
        private CodeBox codeBox;
        private string group;

        #region Constructors

        public CodeBoxPaneItem(LiveCodingPaneWPF parent)
        {
            InitializeComponent();
            this.parent = parent;
            codeBox = new CodeBox(CodeBoxIdService.GenerateUniqueId());
            //insertCode.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabEditButton);
            deleteImage.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabDeleteButton);
            group = "Ungrouped";
        }

        public CodeBoxPaneItem(LiveCodingPaneWPF parent, CodeBox codeBox, string group)
        {
            InitializeComponent();
            this.parent = parent;
            this.codeBox = codeBox;
            //insertCode.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabEditButton);
            deleteImage.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.SyncLabDeleteButton);
            this.group = group;

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

        public string InputType
        {
            get
            {
                return codeBox.InputType;
            }
        }

        public string Slide
        {
            get
            {
                if (codeBox.Slide == null)
                {
                    return "";
                }
                try
                {
                    return codeBox.Slide.Index.ToString();
                }
                catch (COMException)
                {
                    return "";
                }
            }
        }

        public CodeBox CodeBox
        {
            get
            {
                return codeBox;
            }

            set
            {
                codeBox = value;
            }
        }

        public string Group
        {
            get
            {
                return group;
            }

            set
            {
                group = value;
                NotifyPropertyChanged("Group");
            }
        }

        #endregion

        #region Helper Functions
        public void SetDiff()
        {
            codeBox.IsText = false;
            codeBox.IsFile = false;
            codeBox.IsDiff = true;
        }

        public void SetFile()
        {
            codeBox.IsText = false;
            codeBox.IsFile = true;
            codeBox.IsDiff = false;
        }
        public void PopulateTextBox()
        {
            codeTextBox.Text = codeBox.Text;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
        protected void NotifyPropertyChanged(String propertyName)
        {
            OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
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
            if (codeBox.IsDiff && codeBox.DiffIndex < 0)
            {
                MessageBox.Show("Diff file is uninitialised. Please use Insert Diff feature to insert a Diff file",
                                    "Unable to execute action");
                return;
            }
            else if (codeBox.IsDiff && codeBox.DiffIndex >= 0)
            {
                codeBox = ShapeUtility.InsertDiffCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox, CodeBoxFileService.ParseDiff(codeBox.Text)[0]);
            }
            else
            {
                codeBox = ShapeUtility.InsertCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox);
            }

            parent.SaveCodeBox();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            codeBox.Text = codeTextBox.Text;
            codeBox = ShapeUtility.ReplaceTextForShape(codeBox);
            parent.SaveCodeBox();
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            parent.MoveUpCodeBox(this);
            parent.SaveCodeBox();
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            parent.MoveDownCodeBox(this);
            parent.SaveCodeBox();
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            parent.RefreshCode();
            bool hasCodeBoxInSlide = true;
            if (codeBox.Slide == null || codeBox.Shape == null)
            {
                hasCodeBoxInSlide = false;
            }
            else
            {
                try
                {
                    int slideId = codeBox.Slide.ID;
                    int shapeId = codeBox.Shape.Id;
                }
                catch (COMException)
                {
                    hasCodeBoxInSlide = false;
                }
            }

            if (hasCodeBoxInSlide)
            {
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(codeBox.Slide.Index);
                codeBox.Shape.Select();
            }
            else
            {
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                    LiveCodingLabText.PromptToRecreateCodeBox, 
                    LiveCodingLabText.TaskPanelTitle, 
                    System.Windows.Forms.MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    codeBox.Text = codeTextBox.Text;
                    if (codeBox.IsDiff && codeBox.DiffIndex < 0)
                    {
                        MessageBox.Show("Diff file is uninitialised. Please use Insert Diff feature to insert a Diff file",
                                            "Unable to execute action");
                        return;
                    }
                    else if (codeBox.IsDiff && codeBox.DiffIndex >= 0)
                    {
                        codeBox = ShapeUtility.InsertDiffCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox, CodeBoxFileService.ParseDiff(codeBox.Text)[0]);
                    }
                    else
                    {
                        codeBox = ShapeUtility.InsertCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox);
                    }

                    parent.SaveCodeBox();
                }
            }
        }

        private void OnMouseSingleClick(object sender, MouseButtonEventArgs e)
        {
            parent.RefreshCode();
            bool hasCodeBoxInSlide = true;
            if (codeBox.Slide == null || codeBox.Shape == null)
            {
                hasCodeBoxInSlide = false;
            }
            else
            {
                try
                {
                    int slideId = codeBox.Slide.ID;
                    int shapeId = codeBox.Shape.Id;
                }
                catch (COMException)
                {
                    hasCodeBoxInSlide = false;
                }
            }

            if (hasCodeBoxInSlide)
            {
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(codeBox.Slide.Index);
                codeBox.Shape.Select();
            }
            else
            {
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                    LiveCodingLabText.PromptToRecreateCodeBox,
                    LiveCodingLabText.TaskPanelTitle,
                    System.Windows.Forms.MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    codeBox.Text = codeTextBox.Text;
                    if (codeBox.IsDiff && codeBox.DiffIndex < 0)
                    {
                        MessageBox.Show("Diff file is uninitialised. Please use Insert Diff feature to insert a Diff file",
                                            "Unable to execute action");
                        return;
                    }
                    else if (codeBox.IsDiff && codeBox.DiffIndex >= 0)
                    {
                        codeBox = ShapeUtility.InsertDiffCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox, CodeBoxFileService.ParseDiff(codeBox.Text)[0]);
                    }
                    else
                    {
                        codeBox = ShapeUtility.InsertCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox);
                    }

                    parent.SaveCodeBox();
                }
            }
        }

        private void IsFile_Checked(object sender, RoutedEventArgs e)
        {
            codeBox.IsText = false;
            codeBox.IsFile = true;
            codeBox.IsDiff = false;
        }
        private void IsText_Checked(object sender, RoutedEventArgs e)
        {
            codeBox.IsText = true;
            codeBox.IsFile = false;
            codeBox.IsDiff = false;
        }
        private void IsDiff_Checked(object sender, RoutedEventArgs e)
        {
            codeBox.IsText = false;
            codeBox.IsFile = false;
            codeBox.IsDiff = true;
        }

        #endregion
    }
}

using System.Collections.Generic;
using System.Drawing.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace PowerPointLabs.LiveCodingLab.Views
{
    /// <summary>
    /// Interaction logic for InsertDiffDialog.xaml
    /// </summary>
    public partial class FileCodeBoxDialog
    {
        private string filePath;
        private string fileGroupName;

        #region Constructor
        public FileCodeBoxDialog()
        {
            InitializeComponent();
            filePath = "";
            fileGroupName = "Ungrouped";
        }
        #endregion

        #region Attributes
        public string FilePath
        {
            get { return filePath; }
            set { filePath = value; }
        }

        public string FileGroup
        {
            get { return fileGroupName; }
            set { fileGroupName = value; }
        }
        #endregion

        #region Event Handlers
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            filePath = fileInput.Text;
            if (groupNameInput.Text != "")
            {
                fileGroupName = groupNameInput.Text;
            }
            DialogResult = true;
            Close();
        }
        #endregion
    }
}

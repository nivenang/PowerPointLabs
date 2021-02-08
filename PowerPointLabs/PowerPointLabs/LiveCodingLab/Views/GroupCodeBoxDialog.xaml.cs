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
    /// Interaction logic for GroupCodeBoxDialog.xaml
    /// </summary>
    public partial class GroupCodeBoxDialog
    {
        private string groupName;

        #region Constructor
        public GroupCodeBoxDialog(string defaultName)
        {
            InitializeComponent();
            groupName = defaultName;
            groupNameInput.Text = defaultName;
        }
        #endregion

        #region Attributes
        public string ResponseText
        {
            get { return groupName; }
            set { groupName = value; }
        }
        #endregion

        #region Event Handlers
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            groupName = groupNameInput.Text;
            DialogResult = true;
            Close();
        }
        #endregion
    }
}

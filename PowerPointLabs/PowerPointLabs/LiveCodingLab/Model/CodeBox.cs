using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using TestInterface;


using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab.Model
{
    public class CodeBox : INotifyPropertyChanged, IEquatable<CodeBox>
    {
        public event PropertyChangedEventHandler PropertyChanged;

        #region Public properties

        public bool IsFile
        {
            get
            {
                return isFile;
            }
            set
            {
                isFile = (bool)value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_IsFile);
            }
        }

        public bool IsText
        {
            get
            {
                return isText;
            }
            set
            {
                isText = (bool)value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_IsText);
            }
        }

        public bool IsDiff
        {
            get
            {
                return isDiff;
            }
            set
            {
                isDiff = (bool)value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_IsDiff);
            }
        }
        public string FileText
        {
            get
            {
                return codeFile;
            }
            set
            {
                codeFile = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_FileText);
            }
        }

        public string UserText
        {
            get
            {
                return codeText;
            }
            set
            {
                codeText = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_UserText);
            }
        }

        public string DiffText
        {
            get
            {
                return codeDiff;
            }
            set
            {
                codeDiff = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_DiffText);
            }
        }

        public int DiffIndex
        {
            get
            {
                return diffIndex;
            }
            set
            {
                diffIndex = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_DiffIndex);
            }
        }

        public string Text
        {
            get
            {
                if (isFile)
                {
                    return codeFile;
                }
                else if (isText)
                {
                    return codeText;
                }
                else if (isDiff)
                {
                    return codeDiff;
                }
                else
                {
                    return "";
                }
            }

            set
            {
                if (isFile)
                {
                    codeFile = value;
                }
                else if (isText)
                {
                    codeText = value;
                }
                else if (isDiff)
                {
                    codeDiff = value;
                }
                else
                {
                    return;
                }
            }
        }

        public string InputType
        {
            get
            {
                if (isFile)
                {
                    return "File";
                }
                else if (isText)
                {
                    return "Text";
                }
                else if (isDiff)
                {
                    return "Diff";
                }
                else
                {
                    return "";
                }
            }
        }

        public bool IsEmpty
        {
            get
            {
                return string.IsNullOrEmpty(FileText.Trim())
                    && string.IsNullOrEmpty(UserText.Trim())
                    && string.IsNullOrEmpty(DiffText.Trim());
            }
        }

        public int Id
        {
            get
            {
                return codeBoxId;
            }
            set
            {
                codeBoxId = value;
            }
        }

        public PowerPointSlide Slide
        {
            get
            {
                return slide;
            }
            set
            {
                slide = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_Slide);
            }
        }

        public PowerPoint.Shape Shape
        {
            get
            {
                return codeShape;
            }
            set
            {
                codeShape = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_CodeShape);
            }
        }

        public string ShapeName
        {
            get
            {
                return shapeName;
            }
            set
            {
                shapeName = value;
            }
        }

        #endregion

        #region Methods

        public void Delete()
        {
            if (codeShape != null)
            {
                try
                {
                    codeShape.Delete();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return;
                }
            }
        }
        #endregion

        #region Attributes

        private int codeBoxId;
        private PowerPointSlide slide;
        private int diffIndex;

        private bool isFile;
        private bool isText;
        private bool isDiff;

        private string codeFile;
        private string codeText;
        private string codeDiff;

        private PowerPoint.Shape codeShape;

        private string shapeName;
        #endregion

        #region Constructor
        public CodeBox(int codeBoxId, string codeFile = "", string codeText = "", string codeDiff = "", bool isFile = false, bool isText = true, bool isDiff = false, PowerPointSlide slide = null, string shapeName = "", int diffIndex = -1)
        {
            this.codeBoxId = codeBoxId;
            this.slide = slide;
            this.isFile = isFile;
            this.isText = isText;
            this.isDiff = isDiff;
            this.codeFile = codeFile;
            this.codeText = codeText;
            this.codeDiff = codeDiff;
            this.diffIndex = diffIndex;
            this.shapeName = shapeName;
        }
        #endregion

        public override bool Equals(object other)
        {
            if (other == null || other.GetType() != GetType())
            {
                return false;
            }

            if (ReferenceEquals(other, this))
            {
                return true;
            }
            return Equals(other as CodeBox);
        }

        public bool Equals(CodeBox other)
        {
            return codeBoxId == other.Id
                && isFile == other.IsFile
                && isText == other.IsText
                && isDiff == other.IsDiff
                && diffIndex == other.diffIndex
                && codeText.Equals(other.UserText)
                && codeFile.Equals(other.FileText)
                && codeDiff.Equals(other.DiffText);
        }

        public override int GetHashCode()
        {
            var hashCode = -1571720738;
            hashCode = hashCode * -1521134295 + codeBoxId.GetHashCode();
            hashCode = hashCode * -1521134295 + isFile.GetHashCode();
            hashCode = hashCode * -1521134295 + isText.GetHashCode();
            hashCode = hashCode * -1521134295 + isDiff.GetHashCode();
            hashCode = hashCode * -1521134295 + diffIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(codeFile);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(codeText);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(codeDiff);
            return hashCode;
        }

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void CodeBox_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case LiveCodingLabText.CodeBox_CodeShape:
                    try
                    {
                        int shapeId = codeShape.Id;
                    }
                    catch (COMException)
                    {
                        codeShape = null;
                    }
                    break;
                case LiveCodingLabText.CodeBox_Slide:
                    try
                    {
                        int slideID = slide.ID;
                    }
                    catch (COMException)
                    {
                        slide = null;
                        codeShape = null;
                    }
                    break;
                default:
                    break;
            }
        }
    }
}

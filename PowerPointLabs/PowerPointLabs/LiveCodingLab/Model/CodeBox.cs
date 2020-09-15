using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

using Microsoft.Office.Interop.PowerPoint;
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
        public bool IsURL
        {
            get
            {
                return isURL;
            }
            set
            {
                isURL = (bool)value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_IsURL);
            }
        }

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

        public string URLText
        {
            get
            {
                return codeURL;
            }
            set
            {
                codeURL = value;
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_URLText);
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

        public string Text
        {
            get
            {
                if (isURL)
                {
                    return codeURL;
                }
                else if (isFile)
                {
                    return codeFile;
                }
                else if (isText)
                {
                    return codeText;
                }
                else
                {
                    return "";
                }
            }

            set
            {
                if (isURL)
                {
                    codeURL = value;
                }
                else if (isFile)
                {
                    codeFile = value;
                }
                else if (isText)
                {
                    codeText = value;
                }
                else
                {
                    return;
                }
            }
        }

        public bool IsEmpty
        {
            get
            {
                return string.IsNullOrEmpty(URLText.Trim())
                    && string.IsNullOrEmpty(FileText.Trim())
                    && string.IsNullOrEmpty(UserText.Trim());
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
                NotifyPropertyChanged(LiveCodingLabText.CodeBox_Id);
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

        #endregion

        #region Methods

        public void Delete()
        {
            codeShape.Delete();
        }
        #endregion

        #region Attributes

        private int codeBoxId;
        private PowerPointSlide slide;

        private bool isURL;
        private bool isFile;
        private bool isText;

        private string codeURL;
        private string codeFile;
        private string codeText;

        private PowerPoint.Shape codeShape;

        #endregion

        public CodeBox(int codeBoxId, string codeURL = "", string codeFile = "", string codeText = "", bool isURL = false, bool isFile = false, bool isText = true, PowerPointSlide slide=null)
        {
            this.codeBoxId = codeBoxId;
            this.slide = slide;
            this.isURL = isURL;
            this.isFile = isFile;
            this.isText = isText;
            this.codeFile = codeFile;
            this.codeText = codeText;
            this.codeURL = codeURL;
        }

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
                && isURL == other.IsURL
                && isFile == other.IsFile
                && isText == other.IsText
                && codeURL.Equals(other.URLText)
                && codeText.Equals(other.UserText)
                && codeFile.Equals(other.FileText);
        }

        public override int GetHashCode()
        {
            var hashCode = -1571720738;
            hashCode = hashCode * -1521134295 + codeBoxId.GetHashCode();
            hashCode = hashCode * -1521134295 + isURL.GetHashCode();
            hashCode = hashCode * -1521134295 + isFile.GetHashCode();
            hashCode = hashCode * -1521134295 + isText.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(codeFile);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(codeText);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(codeURL);
            return hashCode;
        }

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

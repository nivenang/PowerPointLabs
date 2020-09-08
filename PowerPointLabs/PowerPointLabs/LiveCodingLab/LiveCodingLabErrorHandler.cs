using PowerPointLabs.TextCollection;

namespace PowerPointLabs.LiveCodingLab
{
    internal class LiveCodingLabErrorHandler
    {
        public const int ErrorCodeInvalidSelection = 0;
        public const int ErrorCodeNotSameShapes = 1;
        public const int ErrorCodeGroupShapeNotSupported = 2;

        private ILiveCodingLabPane View { get; set; }
        private static LiveCodingLabErrorHandler _errorHandler;

        private const string ErrorMessageInvalidSelection = LiveCodingLabText.ErrorInvalidSelection;
        private const string ErrorMessageUndefined = LiveCodingLabText.ErrorUndefined;

        private LiveCodingLabErrorHandler(ILiveCodingLabPane view = null)
        {
            View = view;
        }

        public static LiveCodingLabErrorHandler InitializeErrorHandler(ILiveCodingLabPane view = null)
        {
            if (_errorHandler == null)
            {
                _errorHandler = new LiveCodingLabErrorHandler(view);
            }
            else if (view != null) // Allow the view to change
            {
                _errorHandler.View = view;
            }
            return _errorHandler;
        }

        /// <summary>
        /// Store error code in the culture info.
        /// </summary>
        /// <param name="errorType"></param>
        /// <param name="optionalParameters"></param>
        public void ProcessErrorCode(int errorType, params string[] optionalParameters)
        {
            if (View == null) // Nothing to display on
            {
                return;
            }
            string errorMsg = string.Format(GetErrorMessage(errorType), optionalParameters);
            View.ShowErrorMessageBox(errorMsg);
        }

        /// <summary>
        /// Get error message corresponds to the error code.
        /// </summary>
        /// <param name="errorCode"></param>
        /// <returns></returns>
        private string GetErrorMessage(int errorCode)
        {   
            switch (errorCode)
            {
                case ErrorCodeInvalidSelection:
                    return ErrorMessageInvalidSelection;
                default:
                    return ErrorMessageUndefined;
            }
        }
    }
}

using System;

namespace OfficeOpenXml.Utils
{
    public class ValidationResult : IValidationResult
    {
        private string _errorMessage;

        private bool _result;

        public ValidationResult(bool result)
                            : this(result, null)
        {
        }

        public ValidationResult(bool result, string errorMessage)
        {
            _result = result;
            _errorMessage = errorMessage;
        }

        void IValidationResult.IsFalse()
        {
            if (_result)
            {
                Throw();
            }
        }

        void IValidationResult.IsTrue()
        {
            if (!_result)
            {
                Throw();
            }
        }

        private void Throw()
        {
            if(string.IsNullOrEmpty(_errorMessage))
            {
                throw new InvalidOperationException();
            }
            throw new InvalidOperationException(_errorMessage);
        }
    }
}
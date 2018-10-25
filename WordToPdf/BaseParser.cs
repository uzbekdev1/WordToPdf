using System;

namespace WordToPdf
{
    public abstract class BaseParser
    {
        protected BaseParser(string input)
        {
            Input = input;
        }

        protected string Input { get; }

        /// <summary>
        /// Save to pdf include password
        /// </summary>
        /// <param name="output"></param>
        /// <param name="password">Minimum length is 6 characters</param>
        public virtual void ConvertTo(string output, string password = "")
        {
            if (string.IsNullOrWhiteSpace(output))
                throw new ArgumentNullException("output");

            if (!string.IsNullOrWhiteSpace(password) && password.Length < 6)
                throw new Exception("Minimum password length fixing size is 6 characters");

        }
    }
}
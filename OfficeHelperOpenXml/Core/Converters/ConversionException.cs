using System;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// Exception thrown during JSON to PPTX conversion process
    /// </summary>
    public class ConversionException : Exception
    {
        /// <summary>
        /// Gets the context in which the error occurred (e.g., "JSON Deserialization", "Shape Creation")
        /// </summary>
        public string Context { get; set; }

        /// <summary>
        /// Gets the JSON path or location where the error occurred
        /// </summary>
        public string JsonPath { get; set; }

        /// <summary>
        /// Initializes a new instance of ConversionException with a message and context
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="context">The context where the error occurred</param>
        public ConversionException(string message, string context)
            : base(message)
        {
            Context = context;
        }

        /// <summary>
        /// Initializes a new instance of ConversionException with a message, context, and inner exception
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="context">The context where the error occurred</param>
        /// <param name="innerException">The inner exception that caused this exception</param>
        public ConversionException(string message, string context, Exception innerException)
            : base(message, innerException)
        {
            Context = context;
        }

        /// <summary>
        /// Returns a string representation of the exception with context information
        /// </summary>
        public override string ToString()
        {
            var result = $"ConversionException in {Context}: {Message}";
            if (!string.IsNullOrEmpty(JsonPath))
            {
                result += $"\nJSON Path: {JsonPath}";
            }
            if (InnerException != null)
            {
                result += $"\nInner Exception: {InnerException.Message}";
            }
            return result;
        }
    }
}

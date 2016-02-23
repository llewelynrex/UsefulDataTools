using System;

namespace UsefulDataTools
{
    public static class StringExtensions
    {
        /// <summary>
        /// Returns a new string containing the specified number of characters from the left of the input string.
        /// </summary>
        /// <param name="input">The source string for the operation.</param>
        /// <param name="length">The number of characters to return.</param>
        /// <returns>string</returns>
        public static string Left(this string input, int length)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("The input string is null or empty.", nameof(input));

            if (length <= 0)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be greater than 0.");

            if (length > input.Length)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be smaller or equal to the length of the input string.");

            return input.Substring(0, length);
        }

        /// <summary>
        /// Returns a new string containing the specified number of characters from the right of the input string.
        /// </summary>
        /// <param name="input">The source string for the operation.</param>
        /// <param name="length">The number of characters to return.</param>
        /// <returns>string</returns>
        public static string Right(this string input, int length)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("The input string is null or empty.", nameof(input));

            if (length <= 0)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be greater than 0.");

            if (length > input.Length)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be smaller or equal to the length of the input string.");

            return input.Substring(input.Length - length, length);
        }

        /// <summary>
        /// Converts any object to string and then trims the new string.
        /// </summary>
        /// <param name="input">The source string for the operation.</param>
        /// <returns>string</returns>
        public static string ToTrimmedString(this object input)
        {
            var type = input.GetType();
            return type == typeof (string) ? ((string)input).Trim() : input.ToString().Trim();
        }

        /// <summary>
        /// Returns a new string with the specified number of characters added to the left of the input string.
        /// </summary>
        /// <param name="input">The source string for the operation.</param>
        /// <param name="length">The number of characters to add.</param>
        /// <returns>string</returns>
        public static string AddWhitespaceLeft(this string input,int length)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("The input string is null or empty.", nameof(input));

            if (length <= 0)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be greater than 0.");

            return string.Concat(new string(' ', length), input);
        }

        /// <summary>
        /// Returns a new string with the specified number of characters added to the right of the input string.
        /// </summary>
        /// <param name="input">The source string for the operation.</param>
        /// <param name="length">The number of characters to add.</param>
        /// <returns>string</returns>
        public static string AddWhitespaceRight(this string input, int length)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("The input string is null or empty.", nameof(input));

            if (length <= 0)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be greater than 0.");

            return string.Concat(input, new string(' ', length));
        }
    }
}

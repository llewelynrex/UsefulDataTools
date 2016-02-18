﻿using System;

namespace UsefulDataTools
{
    public static class StringExtensions
    {
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

        public static string ToTrimmedString(this object input)
        {
            var type = input.GetType();
            return type == typeof (string) ? ((string)input).Trim() : input.ToString().Trim();
        }

        public static string AddWhitespaceLeft(this string input,int length)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("The input string is null or empty.", nameof(input));

            if (length <= 0)
                throw new ArgumentOutOfRangeException(nameof(length), "The length must be greater than 0.");

            return string.Concat(new string(' ', length), input);
        }

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

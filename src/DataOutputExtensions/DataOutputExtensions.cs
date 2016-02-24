using System;
using System.Reflection;

namespace UsefulDataTools
{
    public static class DataOutputExtensions
    {
        internal static readonly Func<Type, bool> SimpleTypesPredicate;

        static DataOutputExtensions()
        {
            SimpleTypesPredicate = t =>
                                   t == typeof (byte) ||
                                   t == typeof (sbyte) ||
                                   t == typeof (int) ||
                                   t == typeof (uint) ||
                                   t == typeof (short) ||
                                   t == typeof (ushort) ||
                                   t == typeof (long) ||
                                   t == typeof (ulong) ||
                                   t == typeof (float) ||
                                   t == typeof (double) ||
                                   t == typeof (decimal) ||
                                   t == typeof (bool) ||
                                   t == typeof (char) ||
                                   t == typeof (DateTime) ||
                                   t == typeof (string) ||
                                   t.GetTypeInfo().BaseType == typeof (Enum);
        }

        public static bool IsSimpleType<T>(this T type)
            where T : Type
        {
            if (type.Name == "Nullable`1")
                return SimpleTypesPredicate(Nullable.GetUnderlyingType(type));
            return SimpleTypesPredicate(type);
        }

        public static bool IsSimpleType(Type type)
        {
            return type.IsSimpleType();
        }
    }
}
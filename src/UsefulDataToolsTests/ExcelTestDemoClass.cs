using System;
using LinqToExcel.Attributes;

namespace UsefulDataToolsTests
{
    public class ExcelTestDemoClass
    {
        [ExcelColumn("ByteProperty")]
        public byte ByteProperty { get; set; }
        [ExcelColumn("SByteProperty")]
        public sbyte SByteProperty { get; set; }
        [ExcelColumn("IntProperty")]
        public int IntProperty { get; set; }
        [ExcelColumn("UIntProperty")]
        public uint UIntProperty { get; set; }
        [ExcelColumn("ShortProperty")]
        public short ShortProperty { get; set; }
        [ExcelColumn("UShortProperty")]
        public ushort UShortProperty { get; set; }
        [ExcelColumn("LongProperty")]
        public long LongProperty { get; set; }
        [ExcelColumn("ULongProperty")]
        public ulong ULongProperty { get; set; }
        [ExcelColumn("FloatProperty")]
        public float FloatProperty { get; set; }
        [ExcelColumn("DoubleProperty")]
        public double DoubleProperty { get; set; }
        [ExcelColumn("DecimalProperty")]
        public decimal DecimalProperty { get; set; }
        [ExcelColumn("BooleanProperty")]
        public bool BooleanProperty { get; set; }
        [ExcelColumn("CharacterProperty")]
        public char CharacterProperty { get; set; }
        [ExcelColumn("DateTimeProperty")]
        public DateTime DateTimeProperty { get; set; }
        [ExcelColumn("NullableByteProperty")]
        public byte? NullableByteProperty { get; set; }
        [ExcelColumn("NullableSByteProperty")]
        public sbyte? NullableSByteProperty { get; set; }
        [ExcelColumn("NullableIntProperty")]
        public int? NullableIntProperty { get; set; }
        [ExcelColumn("NullableUIntProperty")]
        public uint? NullableUIntProperty { get; set; }
        [ExcelColumn("NullableShortProperty")]
        public short? NullableShortProperty { get; set; }
        [ExcelColumn("NullableUShortProperty")]
        public ushort? NullableUShortProperty { get; set; }
        [ExcelColumn("NullableLongProperty")]
        public long? NullableLongProperty { get; set; }
        [ExcelColumn("NullableULongProperty")]
        public ulong? NullableULongProperty { get; set; }
        [ExcelColumn("NullableFloatProperty")]
        public float? NullableFloatProperty { get; set; }
        [ExcelColumn("NullableDoubleProperty")]
        public double? NullableDoubleProperty { get; set; }
        [ExcelColumn("NullableDecimalProperty")]
        public decimal? NullableDecimalProperty { get; set; }
        [ExcelColumn("NullableBooleanProperty")]
        public bool? NullableBooleanProperty { get; set; }
        [ExcelColumn("NullableCharacterProperty")]
        public char? NullableCharacterProperty { get; set; }
        [ExcelColumn("NullableDateTimeProperty")]
        public DateTime? NullableDateTimeProperty { get; set; }
        [ExcelColumn("StringProperty")]
        public string StringProperty { get; set; }
        public ExcelTestDemoClass ExcelTestDemoClassProperty { get; set; }
    }
}
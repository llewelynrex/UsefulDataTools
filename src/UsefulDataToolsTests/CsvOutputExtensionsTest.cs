using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class CsvOutputExtensionsTest
    {
        public List<TestDemoClass> TestDemoClasses { get; set; }

        [TestInitialize]
        public void Initialize()
        {
            TestDemoClasses = new List<TestDemoClass>();

            var today = DateTime.Today;
            var utc = DateTime.UtcNow;

            var testDemoClass1 = new TestDemoClass
                                 {
                                     BooleanField = true,
                                     BooleanProperty = false,
                                     ByteField = 1,
                                     ByteProperty = 0,
                                     CharacterField = 'a',
                                     CharacterProperty = 'A',
                                     DateTimeField = today,
                                     DateTimeProperty = utc,
                                     DecimalField = 1M,
                                     DecimalProperty = 0M,
                                     DoubleField = 1,
                                     DoubleProperty = 0,
                                     FloatField = 0,
                                     FloatProperty = 0,
                                     IntField = 1,
                                     IntProperty = 0,
                                     LongField = 1,
                                     LongProperty = 0,
                                     NullableBooleanField = false,
                                     NullableBooleanProperty = true,
                                     NullableByteField = 0,
                                     NullableByteProperty = 1,
                                     NullableCharacterField = 'A',
                                     NullableCharacterProperty = 'a',
                                     NullableDateTimeField = utc,
                                     NullableDateTimeProperty = today,
                                     NullableDecimalField = 0M,
                                     NullableDecimalProperty = 1M,
                                     NullableDoubleField = 0,
                                     NullableDoubleProperty = 1,
                                     NullableFloatField = 0,
                                     NullableFloatProperty = 1,
                                     NullableIntField = 0,
                                     NullableIntProperty = 1,
                                     NullableLongField = 0,
                                     NullableLongProperty = 1,
                                     NullableSByteField = 0,
                                     NullableSByteProperty = 1,
                                     NullableShortField = 0,
                                     NullableShortProperty = 1,
                                     NullableUIntField = 0,
                                     NullableUIntProperty = 1,
                                     NullableULongField = 0,
                                     NullableULongProperty = 1,
                                     NullableUShortField = 0,
                                     NullableUShortProperty = 1,
                                     SByteField = 1,
                                     SByteProperty = 0,
                                     ShortField = 1,
                                     ShortProperty = 0,
                                     StringField = "Foo Bar",
                                     StringProperty = "Bar Foo",
                                     UIntField = 1,
                                     UIntProperty = 0,
                                     ULongField = 1,
                                     ULongProperty = 0,
                                     UShortField = 1,
                                     UShortProperty = 0
                                 };

            var testDemoClass2 = new TestDemoClass
                                 {
                                     BooleanField = true,
                                     BooleanProperty = false,
                                     ByteField = 1,
                                     ByteProperty = 0,
                                     CharacterField = 'a',
                                     CharacterProperty = 'A',
                                     DateTimeField = today,
                                     DateTimeProperty = utc,
                                     DecimalField = 1M,
                                     DecimalProperty = 0M,
                                     DoubleField = 1,
                                     DoubleProperty = 0,
                                     FloatField = 0,
                                     FloatProperty = 0,
                                     IntField = 1,
                                     IntProperty = 0,
                                     LongField = 1,
                                     LongProperty = 0,
                                     NullableBooleanField = false,
                                     NullableBooleanProperty = true,
                                     NullableByteField = 0,
                                     NullableByteProperty = 1,
                                     NullableCharacterField = 'A',
                                     NullableCharacterProperty = 'a',
                                     NullableDateTimeField = utc,
                                     NullableDateTimeProperty = today,
                                     NullableDecimalField = 0M,
                                     NullableDecimalProperty = 1M,
                                     NullableDoubleField = 0,
                                     NullableDoubleProperty = 1,
                                     NullableFloatField = 0,
                                     NullableFloatProperty = 1,
                                     NullableIntField = 0,
                                     NullableIntProperty = 1,
                                     NullableLongField = 0,
                                     NullableLongProperty = 1,
                                     NullableSByteField = 0,
                                     NullableSByteProperty = 1,
                                     NullableShortField = 0,
                                     NullableShortProperty = 1,
                                     NullableUIntField = 0,
                                     NullableUIntProperty = 1,
                                     NullableULongField = 0,
                                     NullableULongProperty = 1,
                                     NullableUShortField = 0,
                                     NullableUShortProperty = 1,
                                     SByteField = 1,
                                     SByteProperty = 0,
                                     ShortField = 1,
                                     ShortProperty = 0,
                                     StringField = "Foo Bar",
                                     StringProperty = "Bar Foo",
                                     UIntField = 1,
                                     UIntProperty = 0,
                                     ULongField = 1,
                                     ULongProperty = 0,
                                     UShortField = 1,
                                     UShortProperty = 0
                                 };

            TestDemoClasses.Add(testDemoClass1);
            TestDemoClasses.Add(testDemoClass2);
        }

        [TestMethod]
        public void TestCsvOutput()
        {
            var today = DateTime.Today;
            var utc = DateTime.UtcNow;

            var csv = TestDemoClasses.ToCsv();
        }
    }
}
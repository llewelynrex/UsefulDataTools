﻿using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class XmlOutputExtensionTests
    {
        public List<TestDemoClass> TestDemoClasses { get; set; }
        public DateTime utc { get; set; }
        public DateTime today { get; set; }

        [TestInitialize]
        public void Initialize()
        {
            TestDemoClasses = new List<TestDemoClass>();

            today = DateTime.Today;
            utc = DateTime.UtcNow;

            var testDemoClass1 = new TestDemoClass
            {
                BooleanProperty = false,
                ByteProperty = 0,
                CharacterProperty = 'A',
                DateTimeProperty = utc,
                DecimalProperty = 0M,
                DoubleProperty = 0,
                FloatProperty = 0,
                IntProperty = 0,
                LongProperty = 0,
                NullableBooleanProperty = true,
                NullableByteProperty = 1,
                NullableCharacterProperty = 'a',
                NullableDateTimeProperty = today,
                NullableDecimalProperty = 1M,
                NullableDoubleProperty = 1,
                NullableFloatProperty = 1,
                NullableIntProperty = 1,
                NullableLongProperty = 1,
                NullableSByteProperty = 1,
                NullableShortProperty = 1,
                NullableUIntProperty = 1,
                NullableULongProperty = 1,
                NullableUShortProperty = 1,
                NullableTestDemoEnumProperty = TestDemoEnum.Value2,
                SByteProperty = 0,
                ShortProperty = 0,
                StringProperty = "Bar Foo",
                UIntProperty = 0,
                ULongProperty = 0,
                UShortProperty = 0,
                TestDemoClassProperty = new TestDemoClass(),
                TestDemoEnumProperty = TestDemoEnum.Value1,
                BooleanField = true,
                ByteField = 1,
                CharacterField = 'a',
                DateTimeField = today,
                DecimalField = 1M,
                DoubleField = 1,
                FloatField = 0,
                IntField = 1,
                LongField = 1,
                NullableBooleanField = false,
                NullableByteField = 0,
                NullableCharacterField = 'A',
                NullableDateTimeField = utc,
                NullableDecimalField = 0M,
                NullableDoubleField = 0,
                NullableFloatField = 0,
                NullableIntField = 0,
                NullableLongField = 0,
                NullableSByteField = 0,
                NullableShortField = 0,
                NullableUIntField = 0,
                NullableULongField = 0,
                NullableUShortField = 0,
                NullableTestDemoEnumField = TestDemoEnum.Value2,
                SByteField = 1,
                ShortField = 1,
                StringField = "Foo Bar",
                UIntField = 1,
                ULongField = 1,
                UShortField = 1,
                TestDemoClassField = new TestDemoClass(),
                TestDemoEnumField = TestDemoEnum.Value1
            };
            var testDemoClass2 = new TestDemoClass
            {
                BooleanProperty = false,
                ByteProperty = 0,
                CharacterProperty = 'A',
                DateTimeProperty = utc,
                DecimalProperty = 0M,
                DoubleProperty = 0,
                FloatProperty = 0,
                IntProperty = 0,
                LongProperty = 0,
                NullableBooleanProperty = true,
                NullableByteProperty = 1,
                NullableCharacterProperty = 'a',
                NullableDateTimeProperty = today,
                NullableDecimalProperty = 1M,
                NullableDoubleProperty = 1,
                NullableFloatProperty = 1,
                NullableIntProperty = 1,
                NullableLongProperty = 1,
                NullableSByteProperty = 1,
                NullableShortProperty = 1,
                NullableUIntProperty = 1,
                NullableULongProperty = 1,
                NullableUShortProperty = 1,
                NullableTestDemoEnumProperty = TestDemoEnum.Value2,
                SByteProperty = 0,
                ShortProperty = 0,
                StringProperty = "Bar Foo",
                UIntProperty = 0,
                ULongProperty = 0,
                UShortProperty = 0,
                TestDemoClassProperty = new TestDemoClass(),
                TestDemoEnumProperty = TestDemoEnum.Value1,
                BooleanField = true,
                ByteField = 1,
                CharacterField = 'a',
                DateTimeField = today,
                DecimalField = 1M,
                DoubleField = 1,
                FloatField = 0,
                IntField = 1,
                LongField = 1,
                NullableBooleanField = false,
                NullableByteField = 0,
                NullableCharacterField = 'A',
                NullableDateTimeField = utc,
                NullableDecimalField = 0M,
                NullableDoubleField = 0,
                NullableFloatField = 0,
                NullableIntField = 0,
                NullableLongField = 0,
                NullableSByteField = 0,
                NullableShortField = 0,
                NullableUIntField = 0,
                NullableULongField = 0,
                NullableUShortField = 0,
                NullableTestDemoEnumField = TestDemoEnum.Value2,
                SByteField = 1,
                ShortField = 1,
                StringField = "Foo Bar",
                UIntField = 1,
                ULongField = 1,
                UShortField = 1,
                TestDemoClassField = new TestDemoClass(),
                TestDemoEnumField = TestDemoEnum.Value1
            };

            TestDemoClasses.Add(testDemoClass1);
            TestDemoClasses.Add(testDemoClass2);
        }

        [TestMethod]
        public void TestCsvOutput()
        {
        }
    }
}
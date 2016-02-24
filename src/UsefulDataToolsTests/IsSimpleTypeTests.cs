using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class IsSimpleTypeTests
    {
        private TestDemoClass _testDemoClass;
        private DateTime _today;
        private DateTime _utc;

        [TestInitialize]
        public void Initialize()
        {
            _today = DateTime.Today;
            _utc = DateTime.UtcNow;

            _testDemoClass = new TestDemoClass
                             {
                                 BooleanProperty = false,
                                 ByteProperty = 0,
                                 CharacterProperty = 'A',
                                 DateTimeProperty = _utc,
                                 DecimalProperty = 0M,
                                 DoubleProperty = 0,
                                 FloatProperty = 0,
                                 IntProperty = 0,
                                 LongProperty = 0,
                                 NullableBooleanProperty = true,
                                 NullableByteProperty = 1,
                                 NullableCharacterProperty = 'a',
                                 NullableDateTimeProperty = _today,
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
                                 DateTimeField = _today,
                                 DecimalField = 1M,
                                 DoubleField = 1,
                                 FloatField = 0,
                                 IntField = 1,
                                 LongField = 1,
                                 NullableBooleanField = false,
                                 NullableByteField = 0,
                                 NullableCharacterField = 'A',
                                 NullableDateTimeField = _utc,
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
        }

        [TestMethod]
        public void NonNullableSimpleFieldsTest()
        {
            var fields = typeof (TestDemoClass).GetFields().Where(x => !x.Name.Contains("Nullable") && !x.Name.Contains("Class"));

            foreach (var fieldInfo in fields)
                if (!fieldInfo.FieldType.IsSimpleType()) throw new AssertFailedException($"{fieldInfo.Name}");
        }

        [TestMethod]
        public void NonNullableSimplePropertiesTest()
        {
            var properties = typeof (TestDemoClass).GetProperties().Where(x => !x.Name.Contains("Nullable") && !x.Name.Contains("Class"));

            foreach (var propertyInfo in properties)
                if (!propertyInfo.PropertyType.IsSimpleType()) throw new AssertFailedException($"{propertyInfo.Name}");
        }

        [TestMethod]
        public void NullableSimpleFieldsTest()
        {
            var nullableFields = typeof (TestDemoClass).GetFields().Where(x => x.Name.Contains("Nullable") && !x.Name.Contains("Class"));

            foreach (var fieldInfo in nullableFields)
                if (!fieldInfo.FieldType.IsSimpleType()) throw new AssertFailedException($"{fieldInfo.Name}");
        }

        [TestMethod]
        public void NullableSimplePropertiesTest()
        {
            var nullableProperties = typeof (TestDemoClass).GetProperties().Where(x => x.Name.Contains("Nullable") && !x.Name.Contains("Class"));

            foreach (var propertyInfo in nullableProperties)
                if (!propertyInfo.PropertyType.IsSimpleType()) throw new AssertFailedException($"{propertyInfo.Name}");
        }

        [TestMethod]
        public void NonSimpleFieldTest()
        {
            var fields = typeof (TestDemoClass).GetFields().Where(x => x.Name.Contains("Class"));

            foreach (var fieldInfo in fields)
                if (fieldInfo.FieldType.IsSimpleType()) throw new AssertFailedException($"{fieldInfo.Name}");
        }

        [TestMethod]
        public void NonSimplePropertyTest()
        {
            var properties = typeof (TestDemoClass).GetProperties().Where(x => x.Name.Contains("Class"));

            foreach (var propertyInfo in properties)
                if (propertyInfo.PropertyType.IsSimpleType()) throw new AssertFailedException($"{propertyInfo.Name}");
        }
    }
}
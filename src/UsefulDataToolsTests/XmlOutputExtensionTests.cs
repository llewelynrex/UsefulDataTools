using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class XmlOutputExtensionTests
    {
        private TestDemoClass _testDemoClass;
        private List<TestDemoClass> _testDemoClasses;

        private XDocument _xDocument;

        [TestInitialize]
        public void Initialize()
        {
            var today = DateTime.Today;
            var utc = DateTime.UtcNow;

            _testDemoClass = new TestDemoClass
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

            var _testDemoClass2 = new TestDemoClass
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

            _testDemoClasses = new List<TestDemoClass>(2);
            _testDemoClasses.AddRange(new[] {_testDemoClass, _testDemoClass2});

            _xDocument = new XDocument(new XDeclaration("1.0", "UTF-8", null));
            var rootXElement = new XElement(XName.Get("Root"));

            _xDocument.Add(rootXElement);
        }

        [TestMethod]
        public void TestToXmlSimpleTypeFieldOutput()
        {
            var fields = typeof(TestDemoClass).GetFields().Where(x => !x.Name.Contains("Nullable") && x.FieldType.IsSimpleType());

            foreach (var fieldInfo in fields)
            {
                var testXDocumentString = fieldInfo.GetValue(_testDemoClass).ToXml().ToString();
                var rootXDocumentString = _xDocument.ToString();

                Assert.AreNotEqual(testXDocumentString,rootXDocumentString);
            }
        }

        [TestMethod]
        public void TestToXmlSimpleTypePropertyOutput()
        {
            var properties = typeof(TestDemoClass).GetProperties().Where(x => !x.Name.Contains("Nullable") && x.PropertyType.IsSimpleType());

            foreach (var propertyInfo in properties)
            {
                var testXDocumentString = propertyInfo.GetValue(_testDemoClass).ToXml().ToString();
                var rootXDocumentString = _xDocument.ToString();

                Assert.AreNotEqual(testXDocumentString, rootXDocumentString);
            }
        }

        [TestMethod]
        public void TestToXmlNullableSimpleTypeFieldOutput()
        {
            var fields = typeof(TestDemoClass).GetFields().Where(x => x.Name.Contains("Nullable") && x.FieldType.IsSimpleType());

            foreach (var fieldInfo in fields)
            {
                var testXDocumentString = fieldInfo.GetValue(_testDemoClass).ToXml().ToString();
                var rootXDocumentString = _xDocument.ToString();

                Assert.AreNotEqual(testXDocumentString, rootXDocumentString);
            }
        }

        [TestMethod]
        public void TestToXmlNullableSimpleTypePropertyOutput()
        {
            var properties = typeof(TestDemoClass).GetProperties().Where(x => x.Name.Contains("Nullable") && x.PropertyType.IsSimpleType());

            foreach (var propertyInfo in properties)
            {
                var testXDocumentString = propertyInfo.GetValue(_testDemoClass).ToXml().ToString();
                var rootXDocumentString = _xDocument.ToString();

                Assert.AreNotEqual(testXDocumentString, rootXDocumentString);
            }
        }

        [TestMethod]
        public void TestToXmlComplexTypeOutput() {}

        [TestMethod]
        public void TestToXmlRecursiveOutput() {}

        [TestMethod]
        [ExpectedException(typeof (ArgumentNullException))]
        public void TestNullToXmlOutput()
        {
            string nullVariable = null;
            nullVariable.ToXml();
        }
    }
}
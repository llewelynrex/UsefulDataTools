using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class ExcelOutputExtensionTests
    {
        private const string SimpleTypePath = @"D:\TEMP\SimpleTypeTest.xlsx";
        private const string ComplexTypePath = @"D:\TEMP\ComplexTypeTest.xlsx";

        private DateTime utc { get; set; }
        private DateTime today { get; set; }

        private List<string> _stringList;
        private List<ExcelTestDemoClass> ExcelTestDemoClasses { get; set; }


        [TestInitialize]
        public void Initialize()
        {
            _stringList = new List<string>(new[] {"Test1", "Test2", "Test3"});

            ExcelTestDemoClasses = new List<ExcelTestDemoClass>();

            today = DateTime.Today;
            utc = DateTime.UtcNow;

            var testDemoClass1 = new ExcelTestDemoClass
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
                SByteProperty = 0,
                ShortProperty = 0,
                StringProperty = "Bar Foo",
                UIntProperty = 0,
                ULongProperty = 0,
                UShortProperty = 0,
                ExcelTestDemoClassProperty = new ExcelTestDemoClass()
            };
            var testDemoClass2 = new ExcelTestDemoClass()
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
                SByteProperty = 0,
                ShortProperty = 0,
                StringProperty = "Bar Foo",
                UIntProperty = 0,
                ULongProperty = 0,
                UShortProperty = 0,
                ExcelTestDemoClassProperty = new ExcelTestDemoClass()
            };

            ExcelTestDemoClasses.Add(testDemoClass1);
            ExcelTestDemoClasses.Add(testDemoClass2);
            Cleanup();
        }

        [TestMethod]
        public void ExcelOutputSimpleTypeEnumerationTest()
        {
            _stringList.ToExcel(SimpleTypePath, PostCreationActions.SaveAndClose);
            var excel = new LinqToExcel.ExcelQueryFactory(SimpleTypePath);
            var excelArray = excel.Worksheet(0).Select(x => x[0]).ToArray();

            for (int i = 0; i < _stringList.Count; i++)
            {
                Assert.AreEqual(_stringList[i],excelArray[i]);
            }
        }

        [TestMethod]
        public void ExcelOutputComplexTypeEnumerationTest()
        {
            var type = typeof(ExcelTestDemoClass);

            ExcelTestDemoClasses.ToExcel(ComplexTypePath, PostCreationActions.SaveAndClose);
            var excel = new LinqToExcel.ExcelQueryFactory(ComplexTypePath);
            var excelArray = excel.Worksheet<ExcelTestDemoClass>(0).ToArray();

            for (int i = 0; i < ExcelTestDemoClasses.Count; i++)
            {
                foreach (var propertyInfo in type.GetProperties().Where(x=>x.PropertyType.IsSimpleType()))
                {
                    if (propertyInfo.PropertyType != typeof(DateTime))
                        Assert.AreEqual(propertyInfo.GetValue(excelArray[i]),propertyInfo.GetValue(ExcelTestDemoClasses[i]));
                    else
                    {
                        Assert.AreEqual(((DateTime)propertyInfo.GetValue(excelArray[i])).Year, ((DateTime)propertyInfo.GetValue(ExcelTestDemoClasses[i])).Year);
                        Assert.AreEqual(((DateTime)propertyInfo.GetValue(excelArray[i])).Month, ((DateTime)propertyInfo.GetValue(ExcelTestDemoClasses[i])).Month);
                        Assert.AreEqual(((DateTime)propertyInfo.GetValue(excelArray[i])).Day, ((DateTime)propertyInfo.GetValue(ExcelTestDemoClasses[i])).Day);
                        Assert.AreEqual(((DateTime)propertyInfo.GetValue(excelArray[i])).Hour, ((DateTime)propertyInfo.GetValue(ExcelTestDemoClasses[i])).Hour);
                        Assert.AreEqual(((DateTime)propertyInfo.GetValue(excelArray[i])).Minute, ((DateTime)propertyInfo.GetValue(ExcelTestDemoClasses[i])).Minute);
                        Assert.AreEqual(((DateTime)propertyInfo.GetValue(excelArray[i])).Second, ((DateTime)propertyInfo.GetValue(ExcelTestDemoClasses[i])).Second);
                    }
                }
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(SimpleTypePath))
                File.Delete(SimpleTypePath);
            if (File.Exists(ComplexTypePath))
                File.Delete(ComplexTypePath);
        }
    }
}
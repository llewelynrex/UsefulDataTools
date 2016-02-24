using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class IsSimpleTypeTests
    {
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
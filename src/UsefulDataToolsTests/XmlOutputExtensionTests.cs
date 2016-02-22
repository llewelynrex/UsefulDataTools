using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;


namespace UsefulDataToolsTests
{
    [TestClass]
    public class XmlOutputExtensionTests
    {
        public List<ToCsvTestDemoClass> TestDemoClasses { get; set; }
        public DateTime utc { get; set; }

        [TestInitialize]
        public void Initialize()
        {
            utc = DateTime.UtcNow;
        }

        [TestMethod]
        public void TestToXmlSimpleTypeOutput()
        {
            var byteOutput = ((byte)1).ToXml();
            var sbyteOutput = ((sbyte)1).ToXml();
            var intOutput = ((int)1).ToXml();
            var uintOutput = ((uint)1).ToXml();
            var shortOutput = ((short)1).ToXml();
            var ushortOutput = ((ushort)1).ToXml();
            var longOutput = ((long)1).ToXml();
            var ulongOutput = ((ulong)1).ToXml();
            var floatOutput = ((float)1).ToXml();
            var doubleOutput = ((double)1).ToXml();
            var decimalOutput = ((decimal)1).ToXml();
            var boolOutput = true.ToXml();
            var charOutput = ' '.ToXml();
            var dateTimeOutput = utc.ToXml();
            var nullableByteOutput = ((byte?)1).ToXml();
            var nullableSByteOutput = ((sbyte?)1).ToXml();
            var nullableIntOutput = ((int?)1).ToXml();
            var nullableUIntOutput = ((uint?)1).ToXml();
            var nullableShortOutput = ((short?)1).ToXml();
            var nullableUShortOutput = ((ushort?)1).ToXml();
            var nullableLongOutput = ((long?)1).ToXml();
            var nullableULongOutput = ((ulong?)1).ToXml();
            var nullableFloatOutput = ((float?)1).ToXml();
            var nullableDoubleOutput = ((double?)1).ToXml();
            var nullableDecimalOutput = ((decimal?)1).ToXml();
            var nullableBoolOutput = ((bool?)true).ToXml();
            var nullableCharOutput = ((char?)' ').ToXml();
            var nullableDateTimeOutput = ((DateTime?)utc).ToXml();
            var stringOutput = "Hello World".ToXml();
        }

        [TestMethod]
        public void TestToXmlComplexTypeOutput()
        {

        }

        [TestMethod]
        public void TestToXmlRecursiveOutput()
        {
        }
    }
}
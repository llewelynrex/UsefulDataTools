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
            var intOutput = 1.ToXml();
            var uintOutput = ((uint)1).ToXml();
            var shortOutput = ((short)1).ToXml();
            var ushortOutput = ((ushort)1).ToXml();
            var longOutput = 1L.ToXml();
            var ulongOutput = ((ulong)1).ToXml();
            var floatOutput = 1F.ToXml();
            var doubleOutput = 1D.ToXml();
            var decimalOutput = 1M.ToXml();
            var boolOutput = true.ToXml();
            var charOutput = ' '.ToXml();
            var dateTimeOutput = utc.ToXml();
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

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void TestNullToXmlOutput()
        {
            string nullVariable = null;
            nullVariable.ToXml();
        }
    }
}
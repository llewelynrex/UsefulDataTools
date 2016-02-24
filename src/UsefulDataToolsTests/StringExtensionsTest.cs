using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UsefulDataTools;

namespace UsefulDataToolsTests
{
    [TestClass]
    public class StringExtensionsTest
    {
        [TestMethod]
        public void LeftWorkingTest()
        {
            var teststring = "Bar".Left(2);

            Assert.AreEqual("Ba",teststring);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void LeftArgumentOutOfRangeExceptionTest1()
        {
            "Bar".Left(0);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void LeftArgumentOutOfRangeExceptionTest2()
        {
            "Bar".Left(-1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void LeftArgumentOutOfRangeExceptionTest3()
        {
            "Bar".Left(4);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void LeftArgumentExceptionTest1()
        {
            "".Left(5);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void LeftArgumentExceptionTest2()
        {
            string teststring = null;
            teststring.Left(5);
        }

        [TestMethod]
        public void RightWorkingTest()
        {
            var teststring = "Bar".Right(2);

            Assert.AreEqual("ar", teststring);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void RightArgumentOutOfRangeExceptionTest1()
        {
            "Bar".Right(0);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void RightArgumentOutOfRangeExceptionTest2()
        {
            "Bar".Right(-1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void RightArgumentOutOfRangeExceptionTest3()
        {
            "Bar".Right(4);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void RightArgumentExceptionTest1()
        {
            "".Right(5);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void RightArgumentExceptionTest2()
        {
            string teststring = null;
            teststring.Right(5);
        }

        [TestMethod]
        public void AddWhitespaceLeftWorkingTest()
        {
            var teststring = "Foo".AddWhitespaceLeft(5);

            Assert.AreEqual("     Foo",teststring);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void AddWhitespaceLeftArgumentOutOfRangeExceptionTest1()
        {
            "Foo".AddWhitespaceLeft(0);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void AddWhitespaceLeftArgumentOutOfRangeExceptionTest2()
        {
            "Foo".AddWhitespaceLeft(-1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddWhitespaceLeftArgumentExceptionTest1()
        {
            "".AddWhitespaceLeft(5);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddWhitespaceLeftArgumentExceptionTest2()
        {
            string teststring = null;
            teststring.AddWhitespaceLeft(5);
        }

        [TestMethod]
        public void AddWhitespaceRightWorkingTest()
        {
            var teststring = "Foo".AddWhitespaceRight(5);

            Assert.AreEqual("Foo     ", teststring);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void AddWhitespaceRightArgumentOutOfRangeExceptionTest1()
        {
            "Foo".AddWhitespaceRight(0);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void AddWhitespaceRightArgumentOutOfRangeExceptionTest2()
        {
            "Foo".AddWhitespaceRight(-1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddWhitespaceRightArgumentExceptionTest1()
        {
            "".AddWhitespaceRight(5);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddWhitespaceRightArgumentExceptionTest2()
        {
            string testString = null;
            testString.AddWhitespaceRight(5);
        }

        [TestMethod]
        public void TestToTrimmedStringOnString()
        {
            var testString = "    Foo Bar    ".ToTrimmedString();
            var expectedString = "Foo Bar";
            Assert.AreEqual(expectedString, testString);
        }

        [TestMethod]
        public void TestToTrimmedStringOnObject()
        {
            var testString = new object().ToTrimmedString();
            var expectedString = "System.Object";
            Assert.AreEqual(expectedString, testString);
        }
    }
}

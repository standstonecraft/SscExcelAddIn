using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SscExcelAddIn.Logic.Tests
{
    [TestClass()]
    public class NumStrTests
    {
        [TestMethod()]
        public void NumStrTest()
        {
            Assert.ThrowsException<NotSupportedException>(() => new NumStr("あ"));

            NumStr numStr = new NumStr("0");
            Assert.AreEqual(NumStrType.NN, numStr.StrType);
            Assert.AreEqual(0, numStr.IntValue);

            numStr = new NumStr("1");
            Assert.AreEqual(NumStrType.NN, numStr.StrType);
            Assert.AreEqual(1, numStr.IntValue);
            numStr.Add(12);
            Assert.AreEqual("13", numStr.ToString());
            Assert.AreEqual(13, numStr.IntValue);
            numStr.Set(4);
            Assert.AreEqual("4", numStr.ToString());

            numStr = new NumStr("０");
            Assert.AreEqual(NumStrType.NW, numStr.StrType);
            Assert.AreEqual(0, numStr.IntValue);

            numStr = new NumStr("１");
            numStr.Add(12);
            Assert.AreEqual("１３", numStr.ToString());
            Assert.AreEqual(13, numStr.IntValue);
            numStr.Set(4);
            Assert.AreEqual("４", numStr.ToString());

            numStr = new NumStr("２５");
            numStr.Add(12);
            Assert.AreEqual("３７", numStr.ToString());
            Assert.AreEqual(37, numStr.IntValue);

            numStr = new NumStr("①");
            Assert.AreEqual(NumStrType.M, numStr.StrType);
            numStr.Add(12);
            Assert.AreEqual("⑬", numStr.ToString());
            Assert.AreEqual(13, numStr.IntValue);
            Assert.ThrowsException<IndexOutOfRangeException>(() => numStr.Set(62).ToString());

            numStr = new NumStr("Ⅱ");
            Assert.AreEqual(NumStrType.RU, numStr.StrType);
            numStr.Add(6);
            Assert.AreEqual("Ⅷ", numStr.ToString());
            Assert.AreEqual(8, numStr.IntValue);

            numStr = new NumStr("d");
            Assert.AreEqual(NumStrType.ALN, numStr.StrType);
            numStr.Add(4);
            Assert.AreEqual("h", numStr.ToString());
            Assert.AreEqual(8, numStr.IntValue);

            numStr = new NumStr("D");
            Assert.AreEqual(NumStrType.AUN, numStr.StrType);
            numStr.Add(6);
            Assert.AreEqual("J", numStr.ToString());
            Assert.AreEqual(10, numStr.IntValue);

            numStr = new NumStr("ア");
            Assert.AreEqual(NumStrType.KW, numStr.StrType);
            numStr.Add(6);
            Assert.AreEqual("キ", numStr.ToString());
            Assert.AreEqual(7, numStr.IntValue);

            numStr = new NumStr("ｱ");
            Assert.AreEqual(NumStrType.KN, numStr.StrType);
            numStr.Add(38);
            Assert.AreEqual("ﾗ", numStr.ToString());
            Assert.AreEqual(39, numStr.IntValue);
        }
    }
}

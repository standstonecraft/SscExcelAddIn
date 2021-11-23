using Microsoft.VisualStudio.TestTools.UnitTesting;
using SscExcelAddIn.Logic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SscExcelAddIn.Logic.Tests
{
    [TestClass()]
    public class RegexLogicTests
    {
        [TestMethod()]
        public void ReplaceTextTest()
        {
            string replaced = ReplaceLogic.ReplaceText("asd123dfg", @"\d+", "XXX");
            Assert.AreEqual("asdXXXdfg", replaced);

            replaced = ReplaceLogic.ReplaceText("asd123dfg", @"asd(\d+)", "_$1_");
            Assert.AreEqual("_123_dfg", replaced);

            replaced = ReplaceLogic.ReplaceText("(1) content", @"(\d+)", "_INC($1,2)");
            Assert.AreEqual("(3) content", replaced);

            replaced = ReplaceLogic.ReplaceText("(１) content", @"([０-９]+)", "_NAR($1_NAR)");
            Assert.AreEqual("(1) content", replaced);

            int hit = 0;
            replaced = ReplaceLogic.ReplaceText("(１) content", @"([０-９]+)", "_SEQ($1,1)", ref hit);
            Assert.AreEqual("(１) content", replaced);
            replaced = ReplaceLogic.ReplaceText("(４) content", @"([０-９]+)", "_SEQ($1,1)", ref hit);
            Assert.AreEqual("(２) content", replaced);
            replaced = ReplaceLogic.ReplaceText("(７) content", @"([０-９]+)", "_SEQ($1,1)", ref hit);
            Assert.AreEqual("(３) content", replaced);

            replaced = ReplaceLogic.ReplaceText("(3) content", @"\((\d+)\)", "_CAS($1,M)");
            Assert.AreEqual("③ content", replaced);

            hit = 0;
            replaced = ReplaceLogic.ReplaceText("(１) content", $"({RegexPattern.NUM.Key})", "_CAS(_SEQ($1,1),RU)", ref hit);
            Assert.AreEqual("(Ⅰ) content", replaced);
            replaced = ReplaceLogic.ReplaceText("(3) content", $"({RegexPattern.NUM.Key})", "_CAS(_SEQ($1,1),RU)", ref hit);
            Assert.AreEqual("(Ⅱ) content", replaced);

            hit = 0;
            replaced = ReplaceLogic.ReplaceText("(１) content", $"({RegexPattern.NUM.Key})", "_SEQ(_CAS($1,RU),1)", ref hit);
            Assert.AreEqual("(Ⅰ) content", replaced);
            replaced = ReplaceLogic.ReplaceText("(3) content", $"({RegexPattern.NUM.Key})", "_SEQ(_CAS($1,RU),1)", ref hit);
            Assert.AreEqual("(Ⅱ) content", replaced);
        }
    }
}
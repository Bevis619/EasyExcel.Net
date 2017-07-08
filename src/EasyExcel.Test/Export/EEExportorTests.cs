using EasyExcel.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace EasyExcel.Export.Tests
{
    [TestClass()]
    public class EEExportorTests
    {
        [TestMethod()]
        public void SaveAsActionTest()
        {
            var users = User.MakeUsers();
            var sheets = new List<EESheet> { new EESheet(users), new EESheet(users,"user data") };
            using (var export = new EEExportor(sheets))
            {
                var result = export.SaveAsAction(@"d:\");
                Assert.IsTrue(result, "save as excel faild");

            }
        }
    }
}
using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;

namespace CodeCoverageBug
{
    [TestClass]
    public class ExcelTestMicrosoftUnitTesting
    {
        [TestMethod]
        //A text that executes just fine without code coverage, but hangs with.
        public void ExcelReliantTest()
        {
            var excel = new Excel.Application();
            Console.WriteLine("Hangs on above line when ran with Code Coverage, thus this is never printed.");

            //Quit and release
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }
    }
}

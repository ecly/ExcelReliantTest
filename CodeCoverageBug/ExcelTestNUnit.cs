using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace CodeCoverageBug
{
    [TestFixture]
    public class ExcelTestNUnit
    {
        [Test]
        //A text that executes just fine without code coverage, but hangs with.
        //Using NUnit 2.6.4 and NUnit 2 Test Adapter V. 2.1.1
        public void ExcelReliantTestNUnit()
        {
            var excel = new Excel.Application();
            Console.WriteLine("Hangs on above line when ran with Code Coverage, thus this is never printed.");

            //Quit and release
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }
    }
}

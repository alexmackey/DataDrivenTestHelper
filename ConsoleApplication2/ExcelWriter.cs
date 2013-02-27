
using System;
using System.IO;
using OfficeOpenXml;
//http://zeeshanumardotnet.blogspot.com.au/2011/06/creating-reports-in-excel-2007-using.html

namespace ConsoleApplication2
{
    public class ExcelWriter
    {

        public void DumpPropertiesToExcel(string fileName, string workSheetName, object inputToTest)
        {

            var excelPackage = new ExcelPackage();
            excelPackage.Workbook.Worksheets.Add(workSheetName);
            var ws = excelPackage.Workbook.Worksheets[1];
            var myObjOriginalType = inputToTest.GetType();

            ws.Cells[1, 1].Value = "Created ";
            ws.Cells[1, 2].Value = " " + DateTime.Now;
         
            int rowIndex = 2;

            foreach (var property in myObjOriginalType.GetProperties())
            {
                var propertyInfo = myObjOriginalType.GetProperty(property.Name);
                object actualValue;

                if (propertyInfo != null)
                {
                    actualValue = propertyInfo.GetValue(inputToTest);
                }
                else
                {
                    var fieldInfo = myObjOriginalType.GetField(property.Name);
                    actualValue = fieldInfo.GetValue(inputToTest);
                }

                //to 2dp
                if (actualValue is decimal )
                {
                    actualValue = ((decimal) actualValue).ToString("F2");
                }

                ws.Cells[rowIndex, 1].Value = property.Name;
                ws.Cells[rowIndex, 2].Value = actualValue;

                rowIndex++;

            }

            var bin = excelPackage.GetAsByteArray();
            File.WriteAllBytes(fileName, bin);
        }

       
    }
}

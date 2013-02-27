using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using NSubstitute;
using NSubstitute.Core;

namespace ConsoleApplication2
{
    class Program
    {
        private static string fileName ="test.xlsx";
        private static string worksheetName = "Sheet1";

        private static string fullFileName =
            @"C:\Users\alexj_000\Documents\visual studio 2012\Projects\ConsoleApplication2\ConsoleApplication2\test.xlsx";

        static void Main(string[] args)
        {
            var person = new Person()
                {
                    Age = 32,
                    FirstName = "Alex"
                };

            string fullPath = String.Format("{0}\\{1}",  Environment.CurrentDirectory, fileName);

            var testData = new ExcelDataReader(fullFileName, worksheetName);
            var asserter = new DataLoadAsserter();
            var result= asserter.DoesObjectContainValuesMatchingInput(person, testData);

            Console.WriteLine(result);
            Console.ReadKey();
        }

       
        public class DataLoadAsserter
        {
            public bool DoesObjectContainValuesMatchingInput(object inputToTest, ITestData testData, Func<string, string, bool> compareFunc=null)
            {
                var myObjOriginalType = inputToTest.GetType();
                var testDataValues = testData.GetTestDataAndValues();

                if (compareFunc == null)
                {
                    compareFunc=(actual, expected) => actual != expected;
                }

                foreach (KeyValuePair<String, String> expectedData in testDataValues)
                {
                    var propertyInfo = myObjOriginalType.GetProperty(expectedData.Key);
                    object actualValue;

                    if (propertyInfo != null)
                    {
                       actualValue = propertyInfo.GetValue(inputToTest);
                    }
                    else
                    {
                       var fieldInfo = myObjOriginalType.GetField(expectedData.Key);

                       if (fieldInfo == null)
                       {
                           throw new InvalidOperationException(
                               String.Format("No property {0} found on object of type {1}",
                              expectedData.Key, myObjOriginalType.FullName));
                       }
                       
                       actualValue = fieldInfo.GetValue(inputToTest);
                    }

                    if (compareFunc(actualValue.ToString(), expectedData.Value))
                    {
                        string message = "Unexpected value for {0}. Expected {1} but actually {2}.";
                        Debug.WriteLine(message, expectedData.Key, expectedData.Value, actualValue);
                        return false;
                    }

                }

                return true;
            }

        }

      
       

        
       
    }
}

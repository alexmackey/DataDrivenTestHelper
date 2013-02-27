using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    public class MockDataLoader : ITestData
    {
        private readonly Dictionary<string, string> _expectedValues = new Dictionary<string, string>();

        public MockDataLoader()
        {
            _expectedValues.Add("FirstName", "Alex");
            _expectedValues.Add("Age", "33");
            _expectedValues.Add("blah", "dfdfg");
        }

        public Dictionary<string, string> GetTestDataAndValues()
        {
            return _expectedValues;
        }
    }

}

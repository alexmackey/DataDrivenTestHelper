using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace ConsoleApplication2
{
    public class ExcelDataReader:ITestData
    {
        private readonly Dictionary<string, string> _expectedValues = new Dictionary<string, string>();
        private readonly DataSet _ds = new DataSet();

        //requires http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
        public readonly string _connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;";

        public ExcelDataReader(string fileName, string testDataWorksheet)
        {
            testDataWorksheet += "$";

            var connectionString = string.Format(_connStr, fileName);
            
            var query = String.Format("SELECT * FROM [{0}]", testDataWorksheet);
            var adapter = new OleDbDataAdapter(query, connectionString);

            adapter.Fill(_ds, testDataWorksheet);

            foreach (DataRow row in _ds.Tables[testDataWorksheet].Rows)
            {
                _expectedValues.Add(row["key"].ToString(), row["value"].ToString());
            }
        }
        
        public Dictionary<string, string> GetTestDataAndValues()
        {
            return _expectedValues;
        }
    }
}

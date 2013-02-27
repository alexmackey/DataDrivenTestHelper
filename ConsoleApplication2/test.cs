using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    public class test:ITest
    {
        public List<String> doSomething(int x)
        {
            return new List<string>();
        }
    }

    public interface ITest
    {
        List<String> doSomething(int x);
    }
}

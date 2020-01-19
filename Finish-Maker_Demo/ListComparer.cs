using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Finish_Maker_Demo
{
    public class ListComparer : IEqualityComparer<List<string>>
    {
        public bool Equals(List<string> x, List<string> y)
        {
            for (int i = 0; i < y.Count; i++)
            {
                if (!x[i].Equals(y[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        public int GetHashCode(List<string> obj)
        {
            int hash = 0;
            for (int i = 0; i < obj.Count; i++)
            {
                hash += obj[i].GetHashCode() * i;
            }
            return hash;
        }
    }
}

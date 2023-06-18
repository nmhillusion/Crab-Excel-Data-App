using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrabExcelDataApp.Util
{
    abstract class CollectionUtil
    {
        public static bool IsNullOrEmpty<T>(IEnumerable<T> lists_)
        {
            return null == lists_ || 0 == lists_.Count();
        }
    }
}

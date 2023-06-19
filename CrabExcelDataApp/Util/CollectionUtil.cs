using System.Collections.Generic;
using System.Linq;

namespace CrabExcelDataApp.Util
{
    public abstract class CollectionUtil
    {
        public static bool IsNullOrEmpty<T>(IEnumerable<T> lists_)
        {
            return null == lists_ || 0 == lists_.Count();
        }
    }
}

using System.Collections.Generic;

namespace SeperateDataApp.Service
{
    class DifferenceService
    {
        public ISet<string> distinctListObject(IList<object> allValues)
        {
            SortedSet<string> sortedSet = new();

            foreach (object item in allValues)
            {
                sortedSet.Add(item.ToString());
            }

            return sortedSet;
        }
    }
}

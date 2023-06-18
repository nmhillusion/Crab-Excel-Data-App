namespace CrabExcelDataApp.Util
{
    abstract class StringUtil
    {
        public static string ToString(object obj)
        {
            if (null == obj)
            {
                return "";
            }
            else
            {
                return obj.ToString();
            }
        }
    }
}

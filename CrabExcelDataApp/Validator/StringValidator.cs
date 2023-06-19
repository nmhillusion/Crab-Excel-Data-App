namespace CrabExcelDataApp.Validator
{
    public class StringValidator
    {
        public static bool IsBlank(string input)
        {
            return null == input || 0 == input.Trim().Length;
        }
    }
}

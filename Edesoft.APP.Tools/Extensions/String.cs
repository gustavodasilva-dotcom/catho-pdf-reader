using System.Text.RegularExpressions;

namespace Edesoft.APP.Tools.Extensions
{
    public static class String
    {
        public static string RemoveSpecialCharacters(this string text)
        {
            return Regex.Replace(text, "[^0-9a-zA-Z]+", "");
        }

        public static string FormatRealToExcelForm(this string text)
        {
            return text.Split(",")[0].Replace(".", "");
        }
    }
}

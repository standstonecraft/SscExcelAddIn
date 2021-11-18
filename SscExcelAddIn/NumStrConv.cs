using Microsoft.VisualBasic;
using System;

namespace SscExcelAddIn
{
    internal static class NumStrConv
    {
        static string ZenKanas = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
        internal static int FromStr(string key, string str)
        {
            switch (key)
            {
                case "num":
                    return int.Parse(str);
                case "zenNum":
                    return int.Parse(Strings.StrConv(str, VbStrConv.Narrow));
                case "maruNum":
                    return str.ToCharArray()[0] - '①' + 1;
                case "upperRomeNum":
                    return str.ToCharArray()[0] - 'Ⅰ' + 1;
                case "upper":
                    return str.ToCharArray()[0] - 'A' + 1;
                case "lower":
                    return str.ToCharArray()[0] - 'a' + 1;
                case "zenKana":
                    return ZenKanas.IndexOf(str.ToCharArray()[0]) + 1;
                default:
                    throw new ArgumentException();
            }
        }

        internal static string FromInt(string key, int num)
        {
            switch (key)
            {
                case "num":
                    return num.ToString();
                case "zenNum":
                    return Strings.StrConv(num.ToString(), VbStrConv.Wide);
                case "maruNum":
                    return ((char)(num - 1 + '①')).ToString();
                case "upperRomeNum":
                    return ((char)(num - 1 + 'Ⅰ')).ToString();
                case "upper":
                    return ((char)(num - 1 + 'A')).ToString();
                case "lower":
                    return ((char)(num - 1 + 'a')).ToString();
                case "zenKana":
                    return ZenKanas.ToCharArray()[num - 1].ToString();
                default:
                    throw new ArgumentException();
            }
        }

        internal static string Convert(string keyFrom, string keyTo, string str)
        {
            return FromInt(keyTo, FromStr(keyFrom, str));
        }

        internal static string AddNum(string key, string str, int add)
        {
            return FromInt(key, FromStr(key, str) + add);
        }
        internal static string AddNum(string key, string str, string add)
        {
            return FromInt(key, FromStr(key, str) + int.Parse(add));
        }
    }
}

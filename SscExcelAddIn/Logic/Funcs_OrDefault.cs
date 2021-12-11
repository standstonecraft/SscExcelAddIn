using System;

namespace SscExcelAddIn.Logic
{
    internal static partial class Funcs
    {
        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <typeparam name="TP"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static TP OrDefault<TO, TP>(TO obj, Func<TO, TP> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return default;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static bool? OrDefault<TO>(TO obj, Func<TO, bool> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static byte? OrDefault<TO>(TO obj, Func<TO, byte> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static sbyte? OrDefault<TO>(TO obj, Func<TO, sbyte> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static short? OrDefault<TO>(TO obj, Func<TO, short> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static ushort? OrDefault<TO>(TO obj, Func<TO, ushort> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static int? OrDefault<TO>(TO obj, Func<TO, int> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static uint? OrDefault<TO>(TO obj, Func<TO, uint> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static long? OrDefault<TO>(TO obj, Func<TO, long> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static ulong? OrDefault<TO>(TO obj, Func<TO, ulong> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static float? OrDefault<TO>(TO obj, Func<TO, float> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static double? OrDefault<TO>(TO obj, Func<TO, double> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static decimal? OrDefault<TO>(TO obj, Func<TO, decimal> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// オブジェクトのプロパティを安全に取り出す。
        /// </summary>
        /// <typeparam name="TO"></typeparam>
        /// <param name="obj">オブジェクト</param>
        /// <param name="f">値を取り出す関数</param>
        /// <returns>null許容型</returns>
        public static char? OrDefault<TO>(TO obj, Func<TO, char> f)
        {
            try
            {
                return f.Invoke(obj);
            }
            catch
            {
                return null;
            }
        }
    }
}

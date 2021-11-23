using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;

namespace SscExcelAddIn.Logic
{
    /// <summary>
    /// 番号を表す文字列の演算や文字種変換の機能を提供する。
    /// </summary>
    public class NumStr
    {
        /// <summary>文字列表現</summary>
        private string Value;
        /// <summary>整数値</summary>
        public int IntValue { get; private set; }
        /// <summary>文字列の種類</summary>
        public NumStrType StrType { get; set; }

        private static readonly string AllMaruNum = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿";
        private static readonly string AllZenKana = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨララリルレロワヲン";
        private static readonly string AllHanKana = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾗﾘﾙﾚﾛﾜｦﾝ";

        /// <summary>
        /// 文字列の種類と整数値を判定して保持する。
        /// </summary>
        /// <param name="str"></param>
        public NumStr(string str)
        {
            Value = str;
            Parse();
        }

        private void Parse()
        {
            if (Regex.IsMatch(Value, "^-?[0-9]+$"))
            {
                StrType = NumStrType.NN;
                IntValue = int.Parse(Value);
                return;
            }
            else if (Regex.IsMatch(Value, "^[-－]?[０-９]+$"))
            {
                StrType = NumStrType.NW;
                IntValue = int.Parse(Strings.StrConv(Value, VbStrConv.Narrow));
                return;
            }
            char cValue = Value[0];
            int found;
            if ((found = AllMaruNum.IndexOf(cValue)) > -1)
            {
                StrType = NumStrType.M;
                IntValue = found + 1;
            }
            else if ('Ⅰ' <= cValue && cValue <= 'Ⅻ')
            {
                StrType = NumStrType.RU;
                IntValue = Value[0] - 'Ⅰ' + 1;
            }
            else if ('a' <= cValue && cValue <= 'z')
            {
                StrType = NumStrType.ALN;
                IntValue = Value[0] - 'a' + 1;
            }
            else if ('A' <= cValue && cValue <= 'Z')
            {
                StrType = NumStrType.AUN;
                IntValue = Value[0] - 'A' + 1;
            }
            else if ((found = AllZenKana.IndexOf(cValue)) > -1)
            {
                StrType = NumStrType.KW;
                IntValue = found + 1;
            }
            else if ((found = AllHanKana.IndexOf(cValue)) > -1)
            {
                StrType = NumStrType.KN;
                IntValue = found + 1;
            }
            else
            {
                throw new NotSupportedException();
            }
        }

        /// <summary>
        /// 文字列表現を返す。
        /// </summary>
        /// <returns>文字列表現</returns>
        /// <exception cref="NotSupportedException">文字種類がサポートされていない場合に発生</exception>
        /// <exception cref="IndexOutOfRangeException">文字で表現できない大きさの整数値の場合に発生</exception>
        public override string ToString()
        {
            switch (StrType)
            {
                case NumStrType.U:
                    throw new NotSupportedException();
                case NumStrType.NN:
                    return IntValue.ToString();
                case NumStrType.NW:
                    return Strings.StrConv(IntValue.ToString(), VbStrConv.Wide);
                case NumStrType.M:
                    return AllMaruNum[IntValue - 1].ToString();
                case NumStrType.RU:
                    return ((char)(('Ⅰ' + IntValue - 1))).ToString();
                case NumStrType.ALN:
                    return ((char)(('a' + IntValue - 1))).ToString();
                case NumStrType.AUN:
                    return ((char)(('A' + IntValue - 1))).ToString();
                case NumStrType.KW:
                    return AllZenKana[IntValue - 1].ToString();
                case NumStrType.KN:
                    return AllHanKana[IntValue - 1].ToString();
                default:
                    throw new NotSupportedException();
            }
        }

        /// <summary>
        /// 加算する
        /// </summary>
        /// <param name="num"></param>
        /// <returns>自身</returns>
        public NumStr Add(int num)
        {
            IntValue += num;
            return this;
        }

        /// <summary>
        /// 加算する
        /// </summary>
        /// <param name="num">intに変換可能な文字列</param>
        /// <returns>自身</returns>
        public NumStr Add(string num)
        {
            return Add(int.Parse(num));
        }

        /// <summary>
        /// 整数値をセットする
        /// </summary>
        /// <param name="num"></param>
        /// <returns>自身</returns>
        public NumStr Set(int num)
        {
            IntValue = num;
            return this;
        }

        /// <summary>
        /// 整数値をセットする
        /// </summary>
        /// <param name="num">intに変換可能な文字列</param>
        /// <returns>自身</returns>
        public NumStr Set(string num)
        {
            return Set(int.Parse(num));
        }

        /// <summary>
        /// 文字種類をセットする。文字種類を変換したい場合に用いる。
        /// </summary>
        /// <param name="newType">文字種類</param>
        /// <returns>自身</returns>
        public NumStr SetType(string newType)
        {
            StrType = (NumStrType)Enum.Parse(typeof(NumStrType), newType);
            return this;
        }
    }

}

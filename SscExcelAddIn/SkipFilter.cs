using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SscExcelAddIn
{
    /// <summary>
    /// <see cref="IEnumerable{T}"/> を指定のセレクタに従って飛び飛びに選択して返す機能
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class SkipFilter<T> : IEnumerable<T>
    {
        /// <summary>
        /// 対象のリスト
        /// </summary>
        private readonly IEnumerable<T> Subject;
        /// <summary>
        /// [選択する要素数, 選択しない要素数, ...]
        /// </summary>
        private readonly List<int> SkipSelector;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject">対象のリスト</param>
        /// <param name="skipSelector">[選択する要素数, 選択しない要素数, ...]</param>
        public SkipFilter(IEnumerable<T> subject, IEnumerable<int> skipSelector)
        {
            Subject = subject;
            SkipSelector = skipSelector.ToList();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerable().GetEnumerator();
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return GetEnumerable().GetEnumerator();
        }

        private IEnumerable<T> GetEnumerable()
        {
            if (SkipSelector.Sum() == 0)
            {
                yield break;
            }
            IEnumerator<bool> isNeeded = IsNeeded(SkipSelector).GetEnumerator();
            // 要素の数だけ繰り返す
            foreach (T item in Subject)
            {
                // 本来は判定すべきだが無限に列挙されるので行わない。
                isNeeded.MoveNext();
                if (isNeeded.Current)
                {
                    yield return item;
                }

            }
        }

        /// <summary>
        /// セレクターの要素をイテレートする。最後の要素の次は最初に戻る。無限に列挙する。
        /// </summary>
        /// <param name="SkipSelector"></param>
        /// <returns></returns>
        private IEnumerable<int> Selector(List<int> SkipSelector)
        {
            int elemIndex = 0;
            if (SkipSelector.Count > 0)
            {
                while (true)
                {
                    yield return SkipSelector[elemIndex];
                    elemIndex++;
                    if (elemIndex == SkipSelector.Count)
                    {
                        elemIndex = 0;
                    }
                }
            }
        }

        /// <summary>
        /// 選択対象の場合は真を返す。無限に列挙する。
        /// </summary>
        /// <param name="SkipSelector"></param>
        /// <returns></returns>
        private IEnumerable<bool> IsNeeded(List<int> SkipSelector)
        {
            IEnumerable<int> selector = Selector(SkipSelector);
            bool isNeeded = true;
            foreach (int sel in selector)
            {
                for (int i = 0; i < sel; i++)
                {
                    yield return isNeeded;
                }
                isNeeded = !isNeeded;
            }
        }

    }
}

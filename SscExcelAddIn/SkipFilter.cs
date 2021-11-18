using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SscExcelAddIn
{
    public class SkipFilter<T> : IEnumerable<T>
    {
        private readonly IEnumerable<T> Subject;
        private readonly List<int> SkipSelector;
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
            foreach (T item in Subject)
            {
                if (isNeeded.MoveNext())
                {
                    if (isNeeded.Current)
                    {
                        yield return item;
                    }
                }
                else
                {
                    break;
                }
            }
        }
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

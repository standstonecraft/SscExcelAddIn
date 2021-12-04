using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using Reactive.Bindings;
using Reactive.Bindings.Extensions;

namespace SscExcelAddIn
{
    /// <summary>
    /// <see cref="ShapeEditControl"/> のビューモデル
    /// </summary>
    public class ShapeEditViewModel : INotifyPropertyChanged, IDisposable
    {
        #region IDisposable
        /*****************************
         * IDisposable
         *****************************/
        /// <summary>avoid memory leak</summary>
#pragma warning disable CS0067
        public event PropertyChangedEventHandler PropertyChanged;
#pragma warning restore CS0067

        private readonly CompositeDisposable cd = new CompositeDisposable();
        ///<inheritdoc/>
        public void Dispose() => cd.Dispose();
        #endregion
        #region Properties
        /*****************************
         * Consts
         *****************************/
        #endregion
        /*****************************
         * Fields
         *****************************/
        #region Properties
        /*****************************
         * Properties
         *****************************/
        /// <summary>セル情報リスト</summary>
        public ReactiveCollection<CellContentModel> CellContents { get; set; }
        /// <summary>シェイプ情報リスト</summary>
        public ReactiveCollection<ShapeContentModel> ShapeContents { get; set; }
        /// <summary>検索文字列</summary>
        public ReactiveProperty<string> SearchText { get; set; }
        /// <summary></summary>
        public ReactiveCollection<ShapeContentModel> SearchResults;
        /// <summary>前へ/次へボタンで移動するための現在位置</summary>
        private readonly ReactiveProperty<int> SearchResultPointer;
        #endregion
        #region Commands
        /*****************************
         * Commands
         *****************************/
        /// <summary>検索ボタン</summary>
        public ReactiveCommand SearchCommand { get; set; }
        /// <summary>次へボタン</summary>
        public ReactiveCommand SearchNextCommand { get; set; }
        /// <summary>前へボタン</summary>
        public ReactiveCommand SearchPrevCommand { get; set; }
        #endregion
        /// <summary>ctor</summary>
        public ShapeEditViewModel()
        {
            CellContents = new ReactiveCollection<CellContentModel>();
            ShapeContents = new ReactiveCollection<ShapeContentModel>();
            SearchText = new ReactiveProperty<string>("");
            SearchResults = new ReactiveCollection<ShapeContentModel>();
            SearchResultPointer = new ReactiveProperty<int>(-1);
            SearchCommand = SearchText.Select(x => x.Length > 0)
                .ToReactiveCommand();
            SearchNextCommand = SearchResults.ObserveProperty(x => x.Count)
                .Select(size => size != 0 && SearchResultPointer.Value < size)
                .ToReactiveCommand();
            SearchPrevCommand = SearchResultPointer.Select(index => 0 < index)
                .ToReactiveCommand();

            SearchText.Subscribe(x =>
            {
                SearchResults.Clear();
                SearchResultPointer.Value = -1;
            });
        }

        /// <summary>検索</summary>
        public void Search()
        {
            SearchResults.Clear();
            SearchResultPointer.Value = -1;
            IEnumerable<ShapeContentModel> enumerable = ShapeContents.Where(s => s.Value.IndexOf(SearchText.Value) > -1);
            foreach (ShapeContentModel row in enumerable)
            {
                SearchResults.Add(row);
            }
        }

        /// <summary>次へ</summary>
        public ShapeContentModel SearchNext()
        {
            return SearchResults[++SearchResultPointer.Value];
        }

        /// <summary>前へ</summary>
        public ShapeContentModel SearchPrev()
        {
            return SearchResults[--SearchResultPointer.Value];
        }
    }
}

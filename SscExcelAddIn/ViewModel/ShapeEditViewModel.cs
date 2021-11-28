using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;
using Reactive.Bindings;
using Reactive.Bindings.Extensions;

namespace SscExcelAddIn
{
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
        public ReactiveCollection<CellContentModel> CellContents { get; set; }
        public ReactiveCollection<ShapeContentModel> ShapeContents { get; set; }
        public ReactiveProperty<string> SearchText { get; set; }
        public ReactiveCollection<ShapeContentModel> SearchResults;
        private ReactiveProperty<int> SearchResultPointer;
        #endregion
        #region Commands
        /*****************************
         * Commands
         *****************************/
        public ReactiveCommand SearchCommand { get; set; }
        public ReactiveCommand SearchNextCommand { get; set; }
        public ReactiveCommand SearchPrevCommand { get; set; }
        #endregion
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

        public ShapeContentModel SearchNext()
        {
            return SearchResults[++SearchResultPointer.Value];
        }
        public ShapeContentModel SearchPrev()
        {
            return SearchResults[--SearchResultPointer.Value];
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using Reactive.Bindings;
using Reactive.Bindings.Extensions;
using SscExcelAddIn.Logic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SscExcelAddIn
{
    /// <summary>
    /// ReplaceControl の ViewModel
    /// </summary>
    public class ReplaceViewModel : INotifyPropertyChanged, IDisposable
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
        #region Consts
        /*****************************
         * Consts
         *****************************/
        /// <summary>プレビュー最大件数</summary>
        public const int SampleSize = 30;
        #endregion
        #region Properties
        /*****************************
         * Properties
         *****************************/
        /// <summary>検索文字列</summary>
        public ReactiveProperty<string> PatternText { get; }
        /// <summary>検索文字列 値が入っているか</summary>
        public ReadOnlyReactiveProperty<bool> HasPatternText { get; }
        /// <summary>置換文字列</summary>
        public ReactiveProperty<string> ReplacementText { get; }
        /// <summary>プレビューリスト</summary>
        public ObservableCollection<PreviewModel> PreviewList { get; }
        /// <summary>連続置換ビューモデル</summary>
        public ReplaceBatchViewModel Batch { get; }
        /// <summary>プレビュースライダー 選択値</summary>
        public ReactiveProperty<int> PreviewSliderValue { get; }
        /// <summary>プレビュースライダー 最大値</summary>
        public ReactiveProperty<int> PreviewSliderMax { get; }
        /// <summary>実行ボタン文字列</summary>
        public ReadOnlyReactiveProperty<string> GoCommandContent { get; }
        /// <summary>エラーフラグ</summary>
        public ReactiveProperty<bool> IsError { get; }
        /// <summary>エラーラベル文字列</summary>
        public ReadOnlyReactiveProperty<string> ErrorLabelText { get; }
        /// <summary>デバッグ表示用</summary>
        public ReactiveProperty<string> DebugLabelText { get; }
        #endregion
        #region Commands
        /*****************************
         * Commands
         *****************************/
        /// <summary>更新ボタン</summary>
        public ReactiveCommand RefreshCommand { get; }
        #endregion
        #region Logic
        /*****************************
         * Logic
         *****************************/
        /// <summary>ctor</summary>
        public ReplaceViewModel()
        {
            // property
            PatternText = new ReactiveProperty<string>("").AddTo(cd);
            HasPatternText = PatternText.Select(x => x.Length > 0)
                .ToReadOnlyReactiveProperty().AddTo(cd);
            ReplacementText = new ReactiveProperty<string>("").AddTo(cd);
            Batch = new ReplaceBatchViewModel(this).AddTo(cd);
            PreviewList = new ObservableCollection<PreviewModel>();
            PreviewSliderMax = Batch.Data.ObserveProperty(x => x.Count)
                .ToReactiveProperty().AddTo(cd);
            PreviewSliderValue = new ReactiveProperty<int>().AddTo(cd);
            GoCommandContent = Batch.IsBatchMode.Select(x => x ? "連続" : "置換")
                .ToReadOnlyReactiveProperty().AddTo(cd);
            IsError = new ReactiveProperty<bool>().AddTo(cd);
            ErrorLabelText = IsError.Select(x => x ? "エラー" : "　")
                .ToReadOnlyReactiveProperty().AddTo(cd);
            //DebugLabelText = Batch.HasData.Select(x => x.ToString()).ToReadOnlyReactiveProperty();
            DebugLabelText = new ReactiveProperty<string>("");

            // command
            Batch.Command();
            RefreshCommand = new ReactiveCommand().AddTo(cd);

            // subscribe
            Observable.Merge(PatternText, ReplacementText)
                .Subscribe(x => RefreshPreview(Funcs.GetSample(SampleSize)));
            Batch.Subscribe();
            Batch.Data.CollectionChangedAsObservable()
                .Subscribe(x => RefreshPreview(Funcs.GetSample(SampleSize)));
            PreviewSliderMax.Subscribe(x => PreviewSliderValue.Value = x);
            PreviewSliderValue.Subscribe(x => RefreshPreview(Funcs.GetSample(SampleSize)));
            RefreshCommand.Subscribe(() => RefreshPreview(Funcs.GetSample(SampleSize)));
        }
        /// <summary>
        /// プレビュー更新
        /// </summary>
        /// <param name="sample"></param>
        public void RefreshPreview(IEnumerable<Excel.Range> sample)
        {
            PreviewList.Clear();
            if (sample == null)
            {
                return;
            }
            else
            {
                List<string> sampleTexts = sample.Select(x => (string)x.Formula.ToString()).ToList();
                if (Batch.IsBatchMode.Value)
                {
                    for (int batchIdx = 0; batchIdx < Batch.Data.Count && batchIdx < PreviewSliderValue.Value; batchIdx++)
                    {
                        BatchModel model = Batch.Data[batchIdx];
                        if (!string.IsNullOrEmpty(model.PatternText))
                        {
                            int seqNum = 0;
                            for (int i = 0; i < sampleTexts.Count; i++)
                            {
                                sampleTexts[i] = ReplaceLogic.ReplaceText(sampleTexts[i], model.PatternText, model.ReplacementText, ref seqNum);
                            }
                        }
                    }
                }
                else
                {
                    int seqNum = 0;
                    for (int i = 0; i < sampleTexts.Count; i++)
                    {
                        sampleTexts[i] = ReplaceLogic.ReplaceText(sampleTexts[i], PatternText.Value ?? "", ReplacementText.Value ?? "", ref seqNum);
                    }
                }

                IEnumerable<PreviewModel> result = Enumerable.Zip(sample, sampleTexts, (x, y) => new PreviewModel
                {
                    BeforeText = (string)x.Formula.ToString(),
                    AfterText = y
                });
                foreach (PreviewModel res in result)
                {
                    PreviewList.Add(res);
                }
            }
        }
        #endregion
    }
}

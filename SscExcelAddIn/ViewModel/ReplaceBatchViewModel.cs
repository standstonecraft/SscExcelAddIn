using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Windows;
using Reactive.Bindings;
using Reactive.Bindings.Extensions;

namespace SscExcelAddIn
{
    /// <summary>
    /// 高度な置換コントール 連続置換ビューモデル
    /// </summary>
    public class ReplaceBatchViewModel : INotifyPropertyChanged, IDisposable
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

        private ReplaceViewModel Parent { get; }

        /// <summary>連続置換リスト</summary>
        public ObservableCollection<BatchModel> Data { get; }
        /// <summary>連続置換リスト 選択行</summary>
        public ReactiveProperty<int> Selection { get; }
        /// <summary>連続置換モード</summary>
        public ReadOnlyReactiveProperty<bool> IsBatchMode { get; }
        /// <summary>連続置換データあり</summary>
        public ReadOnlyReactiveProperty<bool> HasData { get; }
        /// <summary>連続置換データ複数あり</summary>
        public ReadOnlyReactiveProperty<bool> HasMultiple { get; }


        /// <summary>🡿</summary>
        public ReactiveCommand ImportCommand { get; set; }
        /// <summary>🡽</summary>
        public ReactiveCommand ExportCommand { get; set; }
        /// <summary>🡹</summary>
        public ReactiveCommand UpCommand { get; set; }
        /// <summary>🡻</summary>
        public ReactiveCommand DownCommand { get; set; }
        /// <summary>＋</summary>
        public ReactiveCommand AddCommand { get; set; }
        /// <summary>－</summary>
        public ReactiveCommand RemoveCommand { get; set; }
        /// <summary>コピー</summary>
        public ReactiveCommand CopyCommand { get; set; }
        /// <summary>貼付</summary>
        public ReactiveCommand PasteCommand { get; set; }

        /// <summary>ctor</summary>
        public ReplaceBatchViewModel(ReplaceViewModel replaceViewModel)
        {
            Parent = replaceViewModel;
            Data = new ObservableCollection<BatchModel>();
            Selection = new ReactiveProperty<int>(-1).AddTo(cd);

            IsBatchMode = Data.ObserveProperty(x => x.Count)
                .Select(x => x > 0)
                .ToReadOnlyReactiveProperty().AddTo(cd);
            HasData = IsBatchMode.ToReadOnlyReactiveProperty().AddTo(cd);
            HasMultiple = Data.ObserveProperty(x => x.Count)
                .Select(x => x > 1)
                .ToReadOnlyReactiveProperty().AddTo(cd);


        }
        /// <summary>
        /// 連続置換関連コマンドの連携
        /// </summary>
        public void Command()
        {
            ImportCommand = Parent.HasPatternText.ToReactiveCommand().AddTo(cd);
            ExportCommand = HasData.ToReactiveCommand().AddTo(cd);
            UpCommand = HasMultiple.ToReactiveCommand().AddTo(cd);
            DownCommand = HasMultiple.ToReactiveCommand().AddTo(cd);
            AddCommand = new ReactiveCommand().AddTo(cd);
            RemoveCommand = HasData.ToReactiveCommand().AddTo(cd);
            CopyCommand = HasData.ToReactiveCommand().AddTo(cd);
            PasteCommand = new ReactiveCommand().AddTo(cd);
        }
        /// <summary>
        /// 連続置換関連コマンドの購読
        /// </summary>
        public void Subscribe()
        {
            ImportCommand.Subscribe(Import);
            ExportCommand.Subscribe(Export);
            UpCommand.Subscribe(Up);
            DownCommand.Subscribe(Down);
            AddCommand.Subscribe(Add);
            RemoveCommand.Subscribe(Remove);
            CopyCommand.Subscribe(Copy);
            PasteCommand.Subscribe(Paste);
        }

        private void Import(object obj)
        {
            int selectedIndex = Selection.Value;
            if (!string.IsNullOrEmpty(Parent.PatternText.Value))
            {
                BatchModel batchModel = new BatchModel
                {
                    PatternText = Parent.PatternText.Value,
                    ReplacementText = Parent.ReplacementText.Value
                };

                if (selectedIndex > -1)
                {
                    Data[selectedIndex] = batchModel;
                }
                else
                {
                    Data.Add(batchModel);
                    Selection.Value = Data.Count - 1;
                    Parent.PreviewSliderValue.Value = Data.Count;
                }
                Parent.PatternText.Value = "";
                Parent.ReplacementText.Value = "";
            }

        }

        private void Export(object obj)
        {
            int selectedIndex = Selection.Value;
            if (selectedIndex > -1)
            {
                Parent.PatternText.Value = Data[selectedIndex].PatternText;
                Parent.ReplacementText.Value = Data[selectedIndex].ReplacementText;
            }
        }

        private void Up(object obj)
        {
            int selectedIndex = Selection.Value;
            if (selectedIndex > 0)
            {
                Data.Move(selectedIndex, selectedIndex - 1);
                Selection.Value = selectedIndex - 1;
            }
        }

        private void Down(object obj)
        {
            int selectedIndex = Selection.Value;
            if (selectedIndex > -1 && selectedIndex < Data.Count - 1)
            {
                Data.Move(selectedIndex, selectedIndex + 1);
                Selection.Value = selectedIndex + 1;
            }
        }

        private void Add(object obj)
        {
            int selectedIndex = Selection.Value;
            int insert = selectedIndex < 0 || selectedIndex == Data.Count - 1
                         ? Data.Count : selectedIndex;

            Data.Insert(insert, new BatchModel
            {
                PatternText = Parent.PatternText.Value,
                ReplacementText = Parent.ReplacementText.Value
            });
            Selection.Value = insert;
            Parent.PreviewSliderValue.Value = Data.Count;
        }

        private void Remove(object obj)
        {
            int selectedIndex = Selection.Value;
            if (selectedIndex > -1)
            {
                // 要素があれば行う
                Data.RemoveAt(Selection.Value);
                // 行が残っていれば、
                if (IsBatchMode.Value)
                {
                    // 元の選択行か残っている行数の小さい方を選択する
                    Selection.Value = Math.Min(selectedIndex, Data.Count - 1);
                }
            }
        }

        private void Copy(object obj)
        {
            string tsv = string.Join("\r\n", Data.Select(row => row.PatternText + "\t" + row.ReplacementText));
            Clipboard.SetText(tsv);
        }

        private void Paste(object obj)
        {
            if (Clipboard.ContainsText())
            {
                try
                {
                    IEnumerable<string[]> data = Clipboard.GetText().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(row => row.Split('\t'));
                    foreach (string[] row in data)
                    {
                        if (row.Length != 2)
                        {
                            throw new ArgumentException();
                        }
                        if (row[0] != "")
                        {
                            Data.Add(new BatchModel
                            {
                                PatternText = row[0],
                                ReplacementText = row[1]
                            });
                        }
                    }
                }
                catch (Exception)
                {

                    throw;
                }
            }
        }
    }
}

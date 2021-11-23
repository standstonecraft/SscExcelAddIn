using System.ComponentModel;

namespace SscExcelAddIn
{
    /// <summary>
    /// ReplaceControl の ViewModel
    /// </summary>
    public class ReplaceViewModel : INotifyPropertyChanged
    {

        //private bool isBatchRowSelectedVal;
        //public bool IsBatchDataGridSelected
        //{
        //    get => isBatchRowSelectedVal;
        //    set
        //    {
        //        isBatchRowSelectedVal = value;
        //        NotifyPropertyChanged("IsBatchDataGridSelected");
        //    }
        //}

        /// <summary>
        /// 
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        private void NotifyPropertyChanged(string info)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
        }
    }

}

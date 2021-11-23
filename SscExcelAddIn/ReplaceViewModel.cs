using System.ComponentModel;

namespace SscExcelAddIn
{
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

        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged(string info)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
        }
    }

}

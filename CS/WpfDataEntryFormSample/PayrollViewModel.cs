using System;
using System.ComponentModel;
using System.Linq;

namespace WpfDataEntryFormSample
{
    public class PayrollViewModel : INotifyPropertyChanged
    {

        public PayrollViewModel(ICollectionView payroll)
        {
            Payroll = payroll ?? throw new ArgumentNullException("payroll");
            Payroll.CurrentChanged += Payroll_CurrentChanged;
        }

        public ICollectionView Payroll { get; }

        public int Count => Payroll.Cast<object>().Count();

        public void MoveFirst() => Payroll.MoveCurrentToFirst();

        public void MovePrevious() => Payroll.MoveCurrentToPrevious();

        public void MoveNext() => Payroll.MoveCurrentToNext();

        public void MoveLast() => Payroll.MoveCurrentToLast();

        public bool CanMovePrevious() => Payroll.CurrentPosition > 0;

        public bool CanMoveNext() => Payroll.CurrentPosition < Count - 1;

        public string DisplayText => !Payroll.IsEmpty ? $"Record {Payroll.CurrentPosition + 1} of {Count}" : string.Empty;

        public event PropertyChangedEventHandler PropertyChanged;

        private void Payroll_CurrentChanged(object sender, EventArgs e) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("DisplayText"));
    }
}

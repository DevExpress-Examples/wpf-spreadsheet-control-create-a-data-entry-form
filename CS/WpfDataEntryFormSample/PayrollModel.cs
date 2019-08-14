using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WpfDataEntryFormSample
{
    /// <summary>
    /// An entity class that exposes basic properties for a payroll. 
    /// </summary>
    public class PayrollModel : INotifyPropertyChanged
    {
        private string employeeName;
        private double hourlyWages;
        private double regularHoursWorked;
        private double vacationHours;
        private double sickHours;
        private double overtimeHours;
        private double overtimeRate;
        private double otherDeduction;
        private int taxStatus;
        private int federalAllowance;
        private double stateTax;
        private double federalIncomeTax;
        private double socialSecurityTax;
        private double medicareTax;
        private double insuranceDeduction;
        private double otherRegularDeduction;

        public string EmployeeName
        {
            get => employeeName;
            set
            {
                if (employeeName != value)
                {
                    employeeName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double HourlyWages
        {
            get => hourlyWages;
            set
            {
                if (hourlyWages != value)
                {
                    hourlyWages = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double RegularHoursWorked
        {
            get => regularHoursWorked;
            set
            {
                if (regularHoursWorked != value)
                {
                    regularHoursWorked = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double VacationHours
        {
            get => vacationHours;
            set
            {
                if (vacationHours != value)
                {
                    vacationHours = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double SickHours
        {
            get => sickHours;
            set
            {
                if (sickHours != value)
                {
                    sickHours = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double OvertimeHours
        {
            get => overtimeHours;
            set
            {
                if (overtimeHours != value)
                {
                    overtimeHours = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double OvertimeRate
        {
            get => overtimeRate;
            set
            {
                if (overtimeRate != value)
                {
                    overtimeRate = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double OtherDeduction
        {
            get => otherDeduction;
            set
            {
                if (otherDeduction != value)
                {
                    otherDeduction = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public int TaxStatus
        {
            get => taxStatus;
            set
            {
                if (taxStatus != value)
                {
                    taxStatus = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public int FederalAllowance
        {
            get => federalAllowance;
            set
            {
                if (federalAllowance != value)
                {
                    federalAllowance = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double StateTax
        {
            get => stateTax;
            set
            {
                if (stateTax != value)
                {
                    stateTax = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double FederalIncomeTax
        {
            get => federalIncomeTax;
            set
            {
                if (federalIncomeTax != value)
                {
                    federalIncomeTax = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double SocialSecurityTax
        {
            get => socialSecurityTax;
            set
            {
                if (socialSecurityTax != value)
                {
                    socialSecurityTax = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double MedicareTax
        {
            get => medicareTax;
            set
            {
                if (medicareTax != value)
                {
                    medicareTax = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double InsuranceDeduction
        {
            get => insuranceDeduction;
            set
            {
                if (insuranceDeduction != value)
                {
                    insuranceDeduction = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double OtherRegularDeduction
        {
            get => otherRegularDeduction;
            set
            {
                if (otherRegularDeduction != value)
                {
                    otherRegularDeduction = value;
                    NotifyPropertyChanged();
                }
            }
        }

        #region INotifyPropertyChanged members
        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property.  
        // The CallerMemberName attribute applied to the optional propertyName parameter  
        // causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}

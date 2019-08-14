Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Namespace WpfDataEntryFormSample

    'An entity class that exposes basic properties for a payroll. 
    Public Class PayrollModel
        Implements INotifyPropertyChanged

        Private _employeeName As String
        Private _hourlyWages As Double
        Private _regularHoursWorked As Double
        Private _vacationHours As Double
        Private _sickHours As Double
        Private _overtimeHours As Double
        Private _overtimeRate As Double
        Private _otherDeduction As Double
        Private _taxStatus As Integer
        Private _federalAllowance As Integer
        Private _stateTax As Double
        Private _federalIncomeTax As Double
        Private _socialSecurityTax As Double
        Private _medicareTax As Double
        Private _insuranceDeduction As Double
        Private _otherRegularDeduction As Double

        Public Property EmployeeName() As String
            Get
                Return _employeeName
            End Get
            Set(ByVal value As String)
                If _employeeName <> value Then
                    _employeeName = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property HourlyWages() As Double
            Get
                Return _hourlyWages
            End Get
            Set(ByVal value As Double)
                If _hourlyWages <> value Then
                    _hourlyWages = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property RegularHoursWorked() As Double
            Get
                Return _regularHoursWorked
            End Get
            Set(ByVal value As Double)
                If _regularHoursWorked <> value Then
                    _regularHoursWorked = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property VacationHours() As Double
            Get
                Return _vacationHours
            End Get
            Set(ByVal value As Double)
                If _vacationHours <> value Then
                    _vacationHours = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property SickHours() As Double
            Get
                Return _sickHours
            End Get
            Set(ByVal value As Double)
                If _sickHours <> value Then
                    _sickHours = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property OvertimeHours() As Double
            Get
                Return _overtimeHours
            End Get
            Set(ByVal value As Double)
                If _overtimeHours <> value Then
                    _overtimeHours = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property OvertimeRate() As Double
            Get
                Return _overtimeRate
            End Get
            Set(ByVal value As Double)
                If _overtimeRate <> value Then
                    _overtimeRate = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property OtherDeduction() As Double
            Get
                Return _otherDeduction
            End Get
            Set(ByVal value As Double)
                If _otherDeduction <> value Then
                    _otherDeduction = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property TaxStatus() As Integer
            Get
                Return _taxStatus
            End Get
            Set(ByVal value As Integer)
                If _taxStatus <> value Then
                    _taxStatus = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property FederalAllowance() As Integer
            Get
                Return _federalAllowance
            End Get
            Set(ByVal value As Integer)
                If _federalAllowance <> value Then
                    _federalAllowance = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property StateTax() As Double
            Get
                Return _stateTax
            End Get
            Set(ByVal value As Double)
                If _stateTax <> value Then
                    _stateTax = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property FederalIncomeTax() As Double
            Get
                Return _federalIncomeTax
            End Get
            Set(ByVal value As Double)
                If _federalIncomeTax <> value Then
                    _federalIncomeTax = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property SocialSecurityTax() As Double
            Get
                Return _socialSecurityTax
            End Get
            Set(ByVal value As Double)
                If _socialSecurityTax <> value Then
                    _socialSecurityTax = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property MedicareTax() As Double
            Get
                Return _medicareTax
            End Get
            Set(ByVal value As Double)
                If _medicareTax <> value Then
                    _medicareTax = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property InsuranceDeduction() As Double
            Get
                Return _insuranceDeduction
            End Get
            Set(ByVal value As Double)
                If _insuranceDeduction <> value Then
                    _insuranceDeduction = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

        Public Property OtherRegularDeduction() As Double
            Get
                Return _otherRegularDeduction
            End Get
            Set(ByVal value As Double)
                If _otherRegularDeduction <> value Then
                    _otherRegularDeduction = value
                    NotifyPropertyChanged()
                End If
            End Set
        End Property

#Region "INotifyPropertyChanged members"
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        ' This method is called by the Set accessor of each property.  
        ' The CallerMemberName attribute applied to the optional propertyName parameter  
        ' causes the property name of the caller to be substituted as an argument.
        Private Sub NotifyPropertyChanged(<CallerMemberName> ByVal Optional propertyName As String = "")
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub
#End Region
    End Class
End Namespace

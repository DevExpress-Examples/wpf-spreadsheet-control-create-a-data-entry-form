Imports System
Imports System.ComponentModel
Imports System.Linq

Namespace WpfDataEntryFormSample
    'A view model for the Data Navigator
    Public Class PayrollViewModel
        Implements INotifyPropertyChanged

        Public Sub New(ByVal payroll As ICollectionView)
            If payroll Is Nothing Then
                Throw New ArgumentNullException("payroll")
            End If
            Me.Payroll = payroll
            AddHandler Me.Payroll.CurrentChanged, AddressOf Payroll_CurrentChanged
        End Sub

        Private privatePayroll As ICollectionView
        Public Property Payroll() As ICollectionView
            Get
                Return privatePayroll
            End Get
            Private Set(ByVal value As ICollectionView)
                privatePayroll = value
            End Set
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return Payroll.Cast(Of Object)().Count()
            End Get
        End Property

        Public Sub MoveFirst()
            Payroll.MoveCurrentToFirst()
        End Sub

        Public Sub MovePrevious()
            Payroll.MoveCurrentToPrevious()
        End Sub

        Public Sub MoveNext()
            Payroll.MoveCurrentToNext()
        End Sub

        Public Sub MoveLast()
            Payroll.MoveCurrentToLast()
        End Sub

        Public Function CanMovePrevious() As Boolean
            Return Payroll.CurrentPosition > 0
        End Function

        Public Function CanMoveNext() As Boolean
            Return Payroll.CurrentPosition < Count - 1
        End Function

        Public ReadOnly Property DisplayText() As String
            Get
                Return If((Not Payroll.IsEmpty), $"Record {Payroll.CurrentPosition + 1} of {Count}", String.Empty)
            End Get
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub Payroll_CurrentChanged(ByVal sender As Object, ByVal e As EventArgs)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("DisplayText"))
        End Sub
    End Class
End Namespace

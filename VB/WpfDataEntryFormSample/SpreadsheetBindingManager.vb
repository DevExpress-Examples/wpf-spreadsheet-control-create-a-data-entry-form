Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports DevExpress.Spreadsheet
Imports DevExpress.Xpf.Spreadsheet

Namespace WpfDataEntryFormSample

    'This class is used to bind the data source object's properties to spreadsheet cells. 
    Public Class SpreadsheetBindingManager
        Private _control As SpreadsheetControl
        Private _dataSource As Object
        Private currentItem As Object
        Private collectionView As ICollectionView
        Private ReadOnly cellBindings As New Dictionary(Of String, String)()
        Private ReadOnly propertyDescriptors As New PropertyDescriptorCollection(Nothing)

        Public Sub New()
        End Sub

        Public Property Control() As SpreadsheetControl
            Get
                Return _control
            End Get
            Set(ByVal value As SpreadsheetControl)
                If Not ReferenceEquals(_control, value) Then
                    If _control IsNot Nothing Then
                        RemoveHandler _control.CellValueChanged, AddressOf SpreadsheetControl_CellValueChanged
                    End If
                    _control = value
                    If _control IsNot Nothing Then
                        AddHandler _control.CellValueChanged, AddressOf SpreadsheetControl_CellValueChanged
                    End If
                End If
            End Set
        End Property

        Public Property SheetName() As String

        Public Property DataSource() As Object
            Get
                Return _dataSource
            End Get
            Set(ByVal value As Object)
                If Not ReferenceEquals(_dataSource, value) Then
                    Detach()
                    _dataSource = value
                    Attach()
                End If
            End Set
        End Property

        ''' <summary>
        ''' Creates a binding between the data source object's property and a cell.
        ''' </summary>
        ''' <param name="propertyName">Specifies the data source property name.</param>
        ''' <param name="cellReference">Specifies a cell reference in the A1 reference style.</param>
        Public Sub AddBinding(ByVal propertyName As String, ByVal cellReference As String)
            If cellBindings.ContainsKey(propertyName) Then Throw New ArgumentException($"Already has binding to {propertyName} property")
            Dim provider As IItemProperties = TryCast(_dataSource, IItemProperties)

            If provider IsNot Nothing Then
                Dim itemProperties = provider.ItemProperties
                Dim propertyInfo As ItemPropertyInfo = itemProperties.SingleOrDefault(Function(p) p.Name = propertyName)
                Dim propertyDescriptor As PropertyDescriptor = If(propertyInfo IsNot Nothing, TryCast(propertyInfo.Descriptor, PropertyDescriptor), Nothing)
                If propertyDescriptor Is Nothing Then Throw New InvalidOperationException($"Unknown {propertyName} property")
                If currentItem IsNot Nothing Then propertyDescriptor.AddValueChanged(currentItem, AddressOf OnPropertyChanged)
                propertyDescriptors.Add(propertyDescriptor)
            End If

            cellBindings.Add(propertyName, cellReference)
        End Sub

        ''' <summary>
        ''' Removes a binding for a data source property.
        ''' </summary>
        ''' <param name="propertyName">Specifies the data source property name.</param>
        Public Sub RemoveBinding(ByVal propertyName As String)
            If cellBindings.ContainsKey(propertyName) Then
                Dim propertyDescriptor As PropertyDescriptor = propertyDescriptors(propertyName)
                If currentItem IsNot Nothing Then
                    propertyDescriptor.RemoveValueChanged(currentItem, AddressOf OnPropertyChanged)
                End If
                propertyDescriptors.Remove(propertyDescriptor)
                cellBindings.Remove(propertyName)
            End If
        End Sub

        ''' <summary>
        ''' Removes all bindings.
        ''' </summary>
        Public Sub ClearBindings()
            UnsubscribePropertyChanged()
            propertyDescriptors.Clear()
            cellBindings.Clear()
        End Sub

        ''' <summary>
        ''' Retrieves the binding manager and property descriptors, and subscribes to the data source and data members events.
        ''' </summary>
        Private Sub Attach()
            collectionView = TryCast(_dataSource, ICollectionView)
            If collectionView IsNot Nothing Then
                AddHandler collectionView.CurrentChanged, AddressOf CollectionView_CurrentChanged
                currentItem = collectionView.CurrentItem
            End If
            Dim provider As IItemProperties = TryCast(_dataSource, IItemProperties)
            If provider IsNot Nothing Then
                Dim itemProperties = provider.ItemProperties
                For Each propertyName As String In cellBindings.Keys
                    Dim propertyInfo As ItemPropertyInfo = itemProperties.SingleOrDefault(Function(p) p.Name = propertyName)
                    Dim propertyDescriptor As PropertyDescriptor = If(propertyInfo IsNot Nothing, TryCast(propertyInfo.Descriptor, PropertyDescriptor), Nothing)
                    If propertyDescriptor Is Nothing Then Throw New InvalidOperationException($"Unable to get property descriptor for {propertyName} property")
                    propertyDescriptors.Add(propertyDescriptor)
                Next
            End If
            PullData()
            SubscribePropertyChanged()
        End Sub

        ''' <summary>
        ''' Unsubscribes from the data source and data members events, and clears property descriptors. 
        ''' </summary>
        Private Sub Detach()
            If _dataSource IsNot Nothing Then
                UnsubscribePropertyChanged()
                If collectionView IsNot Nothing Then
                    RemoveHandler collectionView.CurrentChanged, AddressOf CollectionView_CurrentChanged
                    collectionView = Nothing
                    currentItem = Nothing
                End If
                propertyDescriptors.Clear()
            End If
        End Sub

        Private Sub CollectionView_CurrentChanged(ByVal sender As Object, ByVal e As EventArgs)

            'Update the data entry form when the current item changes.
            _control?.BeginUpdate()
            Try
                DeactivateCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.ActiveCell)
                UnsubscribePropertyChanged()
                currentItem = collectionView.CurrentItem
                PullData()
                SubscribePropertyChanged()
                ActivateCellEditor()
            Finally
                _control?.EndUpdate()
            End Try
        End Sub

        Private Sub UnsubscribePropertyChanged()
            If currentItem IsNot Nothing Then
                For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
                    propertyDescriptor.RemoveValueChanged(currentItem, AddressOf OnPropertyChanged)
                Next propertyDescriptor
            End If
        End Sub

        Private Sub SubscribePropertyChanged()
            If currentItem IsNot Nothing Then
                For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
                    propertyDescriptor.AddValueChanged(currentItem, AddressOf OnPropertyChanged)
                Next propertyDescriptor
            End If
        End Sub

        Private Sub OnPropertyChanged(ByVal sender As Object, ByVal eventArgs As EventArgs)

            'Update the bound cell's value when the corresponding data source property's value changes.
            Dim propertyDescriptor As PropertyDescriptor = TryCast(sender, PropertyDescriptor)
            If propertyDescriptor IsNot Nothing AndAlso currentItem IsNot Nothing Then
                Dim reference As String = Nothing
                If cellBindings.TryGetValue(propertyDescriptor.Name, reference) Then
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)))
                End If
            End If
        End Sub

        ' Pull data from the data source (update values of all bound cells).
        Private Sub PullData()
            If currentItem IsNot Nothing Then
                For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
                    Dim reference As String = cellBindings(propertyDescriptor.Name)
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)))
                Next propertyDescriptor
            End If
        End Sub

        Private Sub SpreadsheetControl_CellValueChanged(ByVal sender As Object, ByVal e As DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs)

            'Update the data source property's value when the bound cell's value changes.
            If e.SheetName = SheetName Then
                Dim reference As String = e.Cell.GetReferenceA1()
                Dim propertyName As String = cellBindings.SingleOrDefault(Function(p) p.Value = reference).Key
                If Not String.IsNullOrEmpty(propertyName) Then
                    Dim propertyDescriptor As PropertyDescriptor = propertyDescriptors(propertyName)
                    If propertyDescriptor IsNot Nothing AndAlso currentItem IsNot Nothing Then
                        propertyDescriptor.SetValue(currentItem, e.Value.ToObject())
                    End If
                End If
            End If
        End Sub

        Private ReadOnly Property Sheet() As Worksheet
            Get
                Return If(_control IsNot Nothing AndAlso _control.Document.Worksheets.Contains(SheetName), _control.Document.Worksheets(SheetName), Nothing)
            End Get
        End Property

        Private Sub SetCellValue(ByVal reference As String, ByVal value As CellValue)
            If Sheet IsNot Nothing Then
                Dim reactivateEditor As Boolean = IsCellEditorActive AndAlso reference = Sheet.Selection.GetReferenceA1()
                If reactivateEditor Then
                    DeactivateCellEditor()
                End If
                Sheet(reference).Value = value
                If reactivateEditor Then
                    ActivateCellEditor()
                End If
            End If
        End Sub

        Private ReadOnly Property IsCellEditorActive() As Boolean
            Get
                Return _control IsNot Nothing AndAlso _control.IsCellEditorActive
            End Get
        End Property

        Private Sub ActivateCellEditor()
            Dim _sheet = Sheet
            If _sheet IsNot Nothing Then
                Dim editors = _sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(_sheet.Selection)
                If editors.Count = 1 Then
                    _control.OpenCellEditor(DevExpress.XtraSpreadsheet.CellEditorMode.Edit)
                End If
            End If
        End Sub

        Private Sub DeactivateCellEditor(Optional ByVal mode As DevExpress.XtraSpreadsheet.CellEditorEnterValueMode = DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Cancel)
            If IsCellEditorActive Then
                _control.CloseCellEditor(mode)
            End If
        End Sub
    End Class
End Namespace


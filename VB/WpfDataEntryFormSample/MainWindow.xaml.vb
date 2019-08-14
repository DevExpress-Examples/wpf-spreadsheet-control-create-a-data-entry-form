Imports System
Imports System.Windows
Imports System.Windows.Data
Imports DevExpress.Spreadsheet
Imports DevExpress.Xpf.Core
Imports DevExpress.Xpf.Editors.Settings

Namespace WpfDataEntryFormSample
	Partial Public Class MainWindow
		Inherits ThemedWindow

		Private ReadOnly payrollViewModel As PayrollViewModel
		Private ReadOnly bindingManager As New SpreadsheetBindingManager()

		Public Sub New()
			InitializeComponent()
			LoadDocumentTemplate()
			BindCustomEditors()
			payrollViewModel = New PayrollViewModel(CType(FindResource("PayrollViewSource"), CollectionViewSource).View)
			DataContext = payrollViewModel
		End Sub

		Private Sub LoadDocumentTemplate()
			spreadsheetControl1.LoadDocument("PayrollCalculatorTemplate.xlsx")
			spreadsheetControl1.Document.History.IsEnabled = False
		End Sub

		Private Sub BindCustomEditors()
            Dim sheet = spreadsheetControl1.ActiveWorksheet

            'Assign custom editors to worksheet cells.
            sheet.CustomCellInplaceEditors.Add(sheet("D8"), CustomCellInplaceEditorType.Custom, "RegularHoursWorked")
			sheet.CustomCellInplaceEditors.Add(sheet("D10"), CustomCellInplaceEditorType.Custom, "VacationHours")
			sheet.CustomCellInplaceEditors.Add(sheet("D12"), CustomCellInplaceEditorType.Custom, "SickHours")
			sheet.CustomCellInplaceEditors.Add(sheet("D14"), CustomCellInplaceEditorType.Custom, "OvertimeHours")
			sheet.CustomCellInplaceEditors.Add(sheet("D16"), CustomCellInplaceEditorType.Custom, "OvertimeRate")
			sheet.CustomCellInplaceEditors.Add(sheet("D22"), CustomCellInplaceEditorType.Custom, "OtherDeduction")
		End Sub


        'Activate a cell editor when a user selects an editable cell.
        Private Function CreateCustomEditorSettings(ByVal tag As String) As BaseEditSettings
			Select Case tag
				Case "RegularHoursWorked"
					Return CreateSpinEditSettings(0, 184, 1)
				Case "VacationHours"
					Return CreateSpinEditSettings(0, 184, 1)
				Case "SickHours"
					Return CreateSpinEditSettings(0, 184, 1)
				Case "OvertimeHours"
					Return CreateSpinEditSettings(0, 100, 1)
				Case "OvertimeRate"
					Return CreateSpinEditSettings(0, 50, 1)
				Case "OtherDeduction"
					Return CreateSpinEditSettings(0, 100, 1)
				Case Else
					Return Nothing
			End Select
		End Function

		Private Function CreateSpinEditSettings(ByVal minValue As Integer, ByVal maxValue As Integer, ByVal increment As Integer) As SpinEditSettings
			Return New SpinEditSettings With {.HorizontalContentAlignment = EditSettingsHorizontalAlignment.Right, .MinValue = minValue, .MaxValue = maxValue, .Increment = increment, .IsFloatValue = False}
		End Function


        'Display a custom in-place editor (a spin editor) for a cell.
        Private Sub SpreadsheetControl1_CustomCellEdit(ByVal sender As Object, ByVal e As DevExpress.Xpf.Spreadsheet.SpreadsheetCustomCellEditEventArgs)
			If e.ValueObject.IsText Then
				e.EditSettings = CreateCustomEditorSettings(e.ValueObject.TextValue)
			End If

		End Sub

        'Suppress the protection warning message.
        Private Sub SpreadsheetControl1_ProtectionWarning(ByVal sender As Object, ByVal e As System.ComponentModel.HandledEventArgs)
			e.Handled = True
		End Sub

		Private Sub SpreadsheetControl1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
			Dim sheet = spreadsheetControl1.ActiveWorksheet
			If sheet IsNot Nothing Then
				Dim editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection)
				If editors.Count = 1 Then
					spreadsheetControl1.OpenCellEditor(DevExpress.XtraSpreadsheet.CellEditorMode.Edit)
				End If
			End If
		End Sub

		Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			bindingManager.Control = spreadsheetControl1
			bindingManager.SheetName = "Payroll Calculator"

            'Bind data source properties to spreadsheet cells. 
            bindingManager.AddBinding("EmployeeName", "C3")
			bindingManager.AddBinding("HourlyWages", "D6")
			bindingManager.AddBinding("RegularHoursWorked", "D8")
			bindingManager.AddBinding("VacationHours", "D10")
			bindingManager.AddBinding("SickHours", "D12")
			bindingManager.AddBinding("OvertimeHours", "D14")
			bindingManager.AddBinding("OvertimeRate", "D16")
			bindingManager.AddBinding("OtherDeduction", "D22")
			bindingManager.AddBinding("TaxStatus", "I4")
			bindingManager.AddBinding("FederalAllowance", "I6")
			bindingManager.AddBinding("StateTax", "I8")
			bindingManager.AddBinding("FederalIncomeTax", "I10")
			bindingManager.AddBinding("SocialSecurityTax", "I12")
			bindingManager.AddBinding("MedicareTax", "I14")
			bindingManager.AddBinding("InsuranceDeduction", "I20")
			bindingManager.AddBinding("OtherRegularDeduction", "I22")


            'Bind the list of PayrollModel objects to the BindingSource. 
            bindingManager.DataSource = payrollViewModel.Payroll
		End Sub
	End Class
End Namespace

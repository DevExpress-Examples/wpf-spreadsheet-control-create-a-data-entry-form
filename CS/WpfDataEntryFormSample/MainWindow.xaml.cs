using System;
using System.Windows;
using System.Windows.Data;
using DevExpress.Spreadsheet;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Editors.Settings;

namespace WpfDataEntryFormSample
{
    public partial class MainWindow : ThemedWindow
    {
        private readonly PayrollViewModel payrollViewModel;
        private readonly SpreadsheetBindingManager bindingManager = new SpreadsheetBindingManager();

        public MainWindow()
        {
            InitializeComponent();
            LoadDocumentTemplate();
            BindCustomEditors();
            payrollViewModel = new PayrollViewModel(((CollectionViewSource)FindResource("PayrollViewSource")).View);
            DataContext = payrollViewModel;
        }

        private void LoadDocumentTemplate()
        {
            spreadsheetControl1.LoadDocument("PayrollCalculatorTemplate.xlsx");
            spreadsheetControl1.Document.History.IsEnabled = false;
        }

        private void BindCustomEditors()
        {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            //Assign custom editors to worksheet cells.
            sheet.CustomCellInplaceEditors.Add(sheet["D8"], CustomCellInplaceEditorType.Custom, "RegularHoursWorked");
            sheet.CustomCellInplaceEditors.Add(sheet["D10"], CustomCellInplaceEditorType.Custom, "VacationHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D12"], CustomCellInplaceEditorType.Custom, "SickHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D14"], CustomCellInplaceEditorType.Custom, "OvertimeHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D16"], CustomCellInplaceEditorType.Custom, "OvertimeRate");
            sheet.CustomCellInplaceEditors.Add(sheet["D22"], CustomCellInplaceEditorType.Custom, "OtherDeduction");
        }

        private BaseEditSettings CreateCustomEditorSettings(string tag)
        {
            switch (tag)
            {
                case "RegularHoursWorked": return CreateSpinEditSettings(0, 184, 1);
                case "VacationHours": return CreateSpinEditSettings(0, 184, 1);
                case "SickHours": return CreateSpinEditSettings(0, 184, 1);
                case "OvertimeHours": return CreateSpinEditSettings(0, 100, 1);
                case "OvertimeRate": return CreateSpinEditSettings(0, 50, 1);
                case "OtherDeduction": return CreateSpinEditSettings(0, 100, 1);
                default: return null;
            }
        }

        private SpinEditSettings CreateSpinEditSettings(int minValue, int maxValue, int increment) => new SpinEditSettings
        {
            HorizontalContentAlignment = EditSettingsHorizontalAlignment.Right,
            MinValue = minValue,
            MaxValue = maxValue,
            Increment = increment,
            IsFloatValue = false
        };

        // Display a custom in-place editor (a spin editor) for a cell.
        private void SpreadsheetControl1_CustomCellEdit(object sender, DevExpress.Xpf.Spreadsheet.SpreadsheetCustomCellEditEventArgs e)
        {
            if (e.ValueObject.IsText)
                e.EditSettings = CreateCustomEditorSettings(e.ValueObject.TextValue);

        }

        // Suppress the protection warning message.
        private void SpreadsheetControl1_ProtectionWarning(object sender, System.ComponentModel.HandledEventArgs e) => e.Handled = true;


        // Activate a cell editor when a user selects an editable cell.
        private void SpreadsheetControl1_SelectionChanged(object sender, EventArgs e)
        {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            if (sheet != null)
            {
                var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
                if (editors.Count == 1)
                    spreadsheetControl1.OpenCellEditor(DevExpress.XtraSpreadsheet.CellEditorMode.Edit);
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Bind data source properties to spreadsheet cells. 
            bindingManager.Control = spreadsheetControl1;
            bindingManager.SheetName = "Payroll Calculator";

            bindingManager.AddBinding("EmployeeName", "C3");
            bindingManager.AddBinding("HourlyWages", "D6");
            bindingManager.AddBinding("RegularHoursWorked", "D8");
            bindingManager.AddBinding("VacationHours", "D10");
            bindingManager.AddBinding("SickHours", "D12");
            bindingManager.AddBinding("OvertimeHours", "D14");
            bindingManager.AddBinding("OvertimeRate", "D16");
            bindingManager.AddBinding("OtherDeduction", "D22");
            bindingManager.AddBinding("TaxStatus", "I4");
            bindingManager.AddBinding("FederalAllowance", "I6");
            bindingManager.AddBinding("StateTax", "I8");
            bindingManager.AddBinding("FederalIncomeTax", "I10");
            bindingManager.AddBinding("SocialSecurityTax", "I12");
            bindingManager.AddBinding("MedicareTax", "I14");
            bindingManager.AddBinding("InsuranceDeduction", "I20");
            bindingManager.AddBinding("OtherRegularDeduction", "I22");

            // Bind the list of PayrollModel objects to the BindingSource.
            bindingManager.DataSource = payrollViewModel.Payroll;
        }
    }
}

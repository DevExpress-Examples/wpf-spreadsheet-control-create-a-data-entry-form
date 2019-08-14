Imports System
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows

Namespace WpfDataEntryFormSample
	''' <summary>
	''' Interaction logic for App.xaml
	''' </summary>
	Partial Public Class App
		Inherits Application

		Protected Overrides Sub OnStartup(ByVal e As StartupEventArgs)
			DevExpress.Xpf.Core.ApplicationThemeHelper.ApplicationThemeName = DevExpress.Xpf.Core.Theme.Office2019ColorfulName
			MyBase.OnStartup(e)
		End Sub
	End Class
End Namespace

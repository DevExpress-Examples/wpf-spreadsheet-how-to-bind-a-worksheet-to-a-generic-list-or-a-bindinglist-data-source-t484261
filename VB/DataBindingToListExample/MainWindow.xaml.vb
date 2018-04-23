Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation
Imports System.Windows.Shapes
Imports DevExpress.Xpf.Core
Imports DevExpress.Spreadsheet
Imports System.ComponentModel

Namespace DataBindingToListExample
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits DevExpress.Xpf.Ribbon.DXRibbonWindow

        Private weatherDataBinding As WorksheetDataBinding
        Private fishesDataBinding As WorksheetDataBinding

        Public Sub New()
            InitializeComponent()
            ribbonControl1.SelectedPage = barPageExample
'            #Region "#ErrorSubscribe"
            AddHandler spreadsheetControl1.Document.Worksheets(0).DataBindings.Error, AddressOf DataBindings_Error
'            #End Region ' #ErrorSubscribe
        End Sub
        #Region "#ErrorHandler"
        Private Sub DataBindings_Error(ByVal sender As Object, ByVal e As DataBindingErrorEventArgs)
            MessageBox.Show(String.Format("Error at worksheet.Rows[{0}]." & ControlChars.Lf & " The error is : {1}", e.RowIndex, e.ErrorType.ToString()), "Binding Error")
        End Sub
        #End Region ' #ErrorHandler

        Private Sub barBtnBindWeather_ItemClick(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            BindWeatherReport(MyWeatherReportSource.Data)
        End Sub
        Private Sub barBtnWeatherBindingList_ItemClick(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            BindWeatherReport(MyWeatherReportSource.DataAsBindingList)
        End Sub

        Private Sub BindWeatherReport(ByVal weatherDatasource As Object)
            If Me.weatherDataBinding IsNot Nothing Then
                spreadsheetControl1.Document.Worksheets(0).DataBindings.Remove(Me.weatherDataBinding)
            End If
'            #Region "#BindTheList"
            ' Specify the binding options.
            Dim dsOptions As New ExternalDataSourceOptions()
            dsOptions.ImportHeaders = True
            dsOptions.CellValueConverter = New MyWeatherConverter()
            dsOptions.SkipHiddenRows = True
            ' Bind the data source to the worksheet range.
            Dim sheet As Worksheet = spreadsheetControl1.Document.Worksheets(0)
            Dim sheetDataBinding As WorksheetDataBinding = sheet.DataBindings.BindToDataSource(weatherDatasource, 2, 1, dsOptions)
'            #End Region ' #BindTheList
            Me.weatherDataBinding = sheetDataBinding
            ' Highlight the binding range.
            Me.weatherDataBinding.Range.FillColor = System.Drawing.Color.Lavender
            ' Adjust column width.
            spreadsheetControl1.Document.Worksheets(0).Range.FromLTRB(1, 1, Me.weatherDataBinding.Range.RightColumnIndex, Me.weatherDataBinding.Range.BottomRowIndex).AutoFitColumns()
        End Sub

        Private Sub barBtnAddWeatherReport_ItemClick(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            MyWeatherReportSource.DataAsBindingList.Insert(1, New WeatherReport() With {.Date = New Date(1776, 2, 29), .Weather = Weather.Sunny, .HourlyReport = MyWeatherReportSource.GenerateRandomHourlyReport()})
        End Sub
        #Region "My Fishes"
        Private Sub barBtnBindMyFishes_ItemClick(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            ' Specify the binding options.
            Dim dsOptions As New ExternalDataSourceOptions()
            dsOptions.ImportHeaders = True
            ' Bind the data source to the worksheet range.
            Me.fishesDataBinding = spreadsheetControl1.Document.Worksheets(0).DataBindings.BindToDataSource(MyFishesSource.Data, 2, 5, dsOptions)
            ' Highlight the binding range.
            Me.fishesDataBinding.Range.FillColor = System.Drawing.Color.LightCyan
            ' Adjust column width.
            spreadsheetControl1.Document.Worksheets(0).Range.FromLTRB(5, 2, Me.fishesDataBinding.Range.RightColumnIndex, Me.fishesDataBinding.Range.BottomRowIndex).AutoFitColumns()

        End Sub
        #End Region ' My Fishes

        Private Sub barBtnImport_ItemClick(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            If Me.fishesDataBinding IsNot Nothing Then
                Me.fishesDataBinding.Range.FillColor = System.Drawing.Color.Empty
                spreadsheetControl1.Document.Worksheets(0).DataBindings.Remove(Me.fishesDataBinding)
            End If
            Dim columnCount As Integer = DisplayColumnHeaders(MyFishesSource.Data, 2, 5)
            spreadsheetControl1.Document.Worksheets(0).Import(MyFishesSource.Data, 3, 5)
            spreadsheetControl1.Document.Worksheets(0).Range.FromLTRB(5, 2, 5 + columnCount, 2 + MyFishesSource.Data.Count).AutoFitColumns()

        End Sub

        Private Sub barBtnUnbind_ItemClick(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            For Each wdb As WorksheetDataBinding In spreadsheetControl1.Document.Worksheets(0).DataBindings
                wdb.Range.FillColor = System.Drawing.Color.Empty
            Next wdb
            weatherDataBinding = Nothing
            fishesDataBinding = Nothing
            spreadsheetControl1.Document.Worksheets(0).DataBindings.Clear()
        End Sub

        Private Function DisplayColumnHeaders(ByVal dataSource As Object, ByVal topRow As Integer, ByVal leftColumn As Integer) As Integer
            ' Get column headers from the data source  
            Dim pdc As PropertyDescriptorCollection = DataSourceHelper.GetSourceProperties(dataSource)
            For i As Integer = 0 To pdc.Count - 1
                Dim pd As PropertyDescriptor = pdc(i)
                spreadsheetControl1.ActiveWorksheet(topRow, i + leftColumn).Value = pd.DisplayName
            Next i
            Return pdc.Count
        End Function
    End Class
End Namespace

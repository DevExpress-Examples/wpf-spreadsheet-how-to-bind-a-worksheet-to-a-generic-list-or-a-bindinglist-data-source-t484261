using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DevExpress.Xpf.Core;
using DevExpress.Spreadsheet;
using System.ComponentModel;

namespace DataBindingToListExample {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DevExpress.Xpf.Ribbon.DXRibbonWindow {
        WorksheetDataBinding weatherDataBinding;
        WorksheetDataBinding fishesDataBinding;

        public MainWindow() {
            InitializeComponent();
            ribbonControl1.SelectedPage = barPageExample;
            #region #ErrorSubscribe
            spreadsheetControl1.Document.Worksheets[0].DataBindings.Error += DataBindings_Error;
            #endregion #ErrorSubscribe
        }
        #region #ErrorHandler
        private void DataBindings_Error(object sender, DataBindingErrorEventArgs e) {
            MessageBox.Show(String.Format("Error at worksheet.Rows[{0}].\n The error is : {1}", e.RowIndex, e.ErrorType.ToString()), "Binding Error");
        }
        #endregion #ErrorHandler

        private void barBtnBindWeather_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e) {
            BindWeatherReport(MyWeatherReportSource.Data);
        }
        private void barBtnWeatherBindingList_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e) {
            BindWeatherReport(MyWeatherReportSource.DataAsBindingList);
        }

        private void BindWeatherReport(object weatherDatasource) {
            if (this.weatherDataBinding != null)
                spreadsheetControl1.Document.Worksheets[0].DataBindings.Remove(this.weatherDataBinding);
            #region #BindTheList
            // Specify the binding options.
            ExternalDataSourceOptions dsOptions = new ExternalDataSourceOptions();
            dsOptions.ImportHeaders = true;
            dsOptions.CellValueConverter = new MyWeatherConverter();
            dsOptions.SkipHiddenRows = true;
            // Bind the data source to the worksheet range.
            Worksheet sheet = spreadsheetControl1.Document.Worksheets[0];
            WorksheetDataBinding sheetDataBinding = sheet.DataBindings.BindToDataSource(weatherDatasource, 2, 1, dsOptions);
            #endregion #BindTheList
            this.weatherDataBinding = sheetDataBinding;
            // Highlight the binding range.
            this.weatherDataBinding.Range.FillColor = System.Drawing.Color.Lavender;
            // Adjust column width.
            spreadsheetControl1.Document.Worksheets[0].Range.FromLTRB(1, 1, this.weatherDataBinding.Range.RightColumnIndex, this.weatherDataBinding.Range.BottomRowIndex).AutoFitColumns();
        }

        private void barBtnAddWeatherReport_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e) {
            MyWeatherReportSource.DataAsBindingList.Insert(1, new WeatherReport() {
                Date = new DateTime(1776, 2, 29),
                Weather = Weather.Sunny,
                HourlyReport = MyWeatherReportSource.GenerateRandomHourlyReport()
            });
        }
        #region My Fishes
        private void barBtnBindMyFishes_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e) {
            // Specify the binding options.
            ExternalDataSourceOptions dsOptions = new ExternalDataSourceOptions();
            dsOptions.ImportHeaders = true;
            // Bind the data source to the worksheet range.
            this.fishesDataBinding = spreadsheetControl1.Document.Worksheets[0].DataBindings.BindToDataSource(MyFishesSource.Data, 2, 5, dsOptions);
            // Highlight the binding range.
            this.fishesDataBinding.Range.FillColor = System.Drawing.Color.LightCyan;
            // Adjust column width.
            spreadsheetControl1.Document.Worksheets[0].Range.FromLTRB(5, 2, this.fishesDataBinding.Range.RightColumnIndex, this.fishesDataBinding.Range.BottomRowIndex).AutoFitColumns();

        }
        #endregion My Fishes

        private void barBtnImport_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e) {
            if (this.fishesDataBinding != null) {
                this.fishesDataBinding.Range.FillColor = System.Drawing.Color.Empty;
                spreadsheetControl1.Document.Worksheets[0].DataBindings.Remove(this.fishesDataBinding);
            }
            int columnCount = DisplayColumnHeaders(MyFishesSource.Data, 2, 5);
            spreadsheetControl1.Document.Worksheets[0].Import(MyFishesSource.Data, 3, 5);
            spreadsheetControl1.Document.Worksheets[0].Range.FromLTRB(5, 2, 5 + columnCount, 2 + MyFishesSource.Data.Count).AutoFitColumns();

        }

        private void barBtnUnbind_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e) {
            foreach (WorksheetDataBinding wdb in spreadsheetControl1.Document.Worksheets[0].DataBindings)
                wdb.Range.FillColor = System.Drawing.Color.Empty;
            weatherDataBinding = null;
            fishesDataBinding = null;
            spreadsheetControl1.Document.Worksheets[0].DataBindings.Clear();
        }

        private int DisplayColumnHeaders(object dataSource, int topRow, int leftColumn) {
            // Get column headers from the data source  
            PropertyDescriptorCollection pdc = DataSourceHelper.GetSourceProperties(dataSource);
            for (int i = 0; i < pdc.Count; i++) {
                PropertyDescriptor pd = pdc[i];
                spreadsheetControl1.ActiveWorksheet[topRow, i + leftColumn].Value = pd.DisplayName;
            }
            return pdc.Count;
        }
    }
}

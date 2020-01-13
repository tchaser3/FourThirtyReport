/* Title:           Four Thirty Report
 * Date:            1-8-20
 * Author:          Terry Holmes
 * 
 * Description:     This is used to generate the 4:30 report */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEmployeeDLL;
using NewEventLogDLL;
using VehicleMainDLL;
using VehicleAssignmentDLL;
using Excel = Microsoft.Office.Interop.Excel;
using VehiclesInShopDLL;
using VehicleInYardDLL;
using DateSearchDLL;
using Microsoft.Win32;

namespace FourThirtyReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        VehicleMainClass TheVehicleMainClass = new VehicleMainClass();
        VehicleAssignmentClass TheVehicleAssignmentClass = new VehicleAssignmentClass();
        VehiclesInShopClass TheVehiclesInShopClass = new VehiclesInShopClass();
        VehicleInYardClass TheVehicleInYardClass = new VehicleInYardClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();

        //setting up the data
        FindActiveVehicleMainDataSet TheFindActiveVehicleMainDataSet = new FindActiveVehicleMainDataSet();
        FindCurrentAssignedVehicleByVehicleIDDataSet TheFindCurrentVehicleAssignedVehicleByVehicleIDDataSet = new FindCurrentAssignedVehicleByVehicleIDDataSet();
        AutomileImportDataSet TheAutomileImportDataSet = new AutomileImportDataSet();
        VehicleCurrentStatusDataSet TheVehicleCurrentStatusDataSet = new VehicleCurrentStatusDataSet();
        FindCurrentAssignedVehicleMainByVehicleIDDataSet TheFindCurrentAssignedVehicleMainByVehicleIDDataSet = new FindCurrentAssignedVehicleMainByVehicleIDDataSet();
        FindVehicleMainInShopByVehicleIDDataSet TheFindVehicleInShopByVehicleIDDataSet = new FindVehicleMainInShopByVehicleIDDataSet();
        FindVehiclesInYardByVehicleIDAndDateRangeDataSet TheFindVehicleInYardByVehicleIDAndDateRangeDataSet = new FindVehiclesInYardByVehicleIDAndDateRangeDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEnmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        ShortedCurrentVehicleDataSet TheShortedCurrentVehicleDataSet = new ShortedCurrentVehicleDataSet();

        //setting up global on it
        int gintThirdCounter;
        int gintThirdNumberOfRecords;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void expClose_Expanded(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void expImportExcel_Expanded(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strGeoFence;
            string strIsInside;
            string strDriver;
            string strVehicleNumber;

            try
            {
                TheAutomileImportDataSet.automileimport.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 5; intCounter <= intNumberOfRecords; intCounter++)
                {

                    strGeoFence = Convert.ToString((range.Cells[intCounter, 7] as Excel.Range).Value2).ToUpper();
                    strIsInside = Convert.ToString((range.Cells[intCounter, 9] as Excel.Range).Value2).ToUpper();
                    strDriver = Convert.ToString((range.Cells[intCounter, 10] as Excel.Range).Value2).ToUpper();
                    strVehicleNumber = Convert.ToString((range.Cells[intCounter, 11] as Excel.Range).Value2).ToUpper();

                    AutomileImportDataSet.automileimportRow NewVehicleRow = TheAutomileImportDataSet.automileimport.NewautomileimportRow();

                    NewVehicleRow.Driver = strDriver;
                    NewVehicleRow.GeoFence = strGeoFence;
                    NewVehicleRow.IsInside = strIsInside;
                    NewVehicleRow.VehicleNumber = strVehicleNumber;

                    TheAutomileImportDataSet.automileimport.Rows.Add(NewVehicleRow);
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheAutomileImportDataSet.automileimport;
                
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Four Thirty Report // Main Window // Import Excel  " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void expProcess_Expanded(object sender, RoutedEventArgs e)
        {
            
        }

        private void expProcess_Expanded_1(object sender, RoutedEventArgs e)
        {
            int intVehicleCounter;
            int intVehicleNumberOfRecords;
            int intAutomileCounter;
            int intAutomileNumberOfRecords;
            int intVehicleID;
            string strVehicleNumber;
            string strAutomileVehicleNumber;
            bool blnItemFound;
            bool blnItemEntered;
            int intThirdCounter;
            string strInOrOut;
            string strManager;
            string strFullName;
            string strAssignedOffice;
            DateTime datStartDate;
            DateTime datEndDate;
            int intRecordsReturned;
            int intManagerID;

            try
            {
                //loading the vehicle data set
                TheFindActiveVehicleMainDataSet = TheVehicleMainClass.FindActiveVehicleMain();
                TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.Rows.Clear();
                TheShortedCurrentVehicleDataSet.shortenlist.Rows.Clear();

                datStartDate = DateTime.Now;
                datStartDate = TheDateSearchClass.RemoveTime(datStartDate);
                datEndDate = TheDateSearchClass.AddingDays(datStartDate, 1);

                intVehicleNumberOfRecords = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain.Rows.Count - 1;
                intAutomileNumberOfRecords = TheAutomileImportDataSet.automileimport.Rows.Count - 1;
                gintThirdCounter = 0;
                gintThirdNumberOfRecords = 0;

                for (intVehicleCounter = 0; intVehicleCounter <= intVehicleNumberOfRecords; intVehicleCounter++)
                {
                    intVehicleID = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intVehicleCounter].VehicleID;
                    strVehicleNumber = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intVehicleCounter].VehicleNumber;

                    strAssignedOffice = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intVehicleCounter].AssignedOffice;

                    TheFindCurrentAssignedVehicleMainByVehicleIDDataSet = TheVehicleAssignmentClass.FindCurrentAssignedVehicleMainByVehicleID(intVehicleID);

                    strFullName = TheFindCurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].FirstName + " ";
                    strFullName += TheFindCurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].LastName;
                    intManagerID = TheFindCurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].ManagerID;

                    TheFindEnmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);

                    strManager = TheFindEnmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                    strManager += TheFindEnmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;
                    strInOrOut = "UNKNOWN";

                    if(TheFindCurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].LastName == "WAREHOUSE")
                    {
                        strManager = "FLEET MANAGER";
                    }

                    TheFindVehicleInShopByVehicleIDDataSet = TheVehiclesInShopClass.FindVehicleMainInShopByVehicleID(intVehicleID);
                    
                    TheFindVehicleInYardByVehicleIDAndDateRangeDataSet = TheVehicleInYardClass.FindVehiclesInYardByVehicleIDAndDateRange(intVehicleID, datStartDate, datEndDate);

                    intRecordsReturned = TheFindVehicleInYardByVehicleIDAndDateRangeDataSet.FindVehiclesInYardByVehicleIDAndDateRange.Rows.Count;

                    if (intRecordsReturned > 0)
                    {
                        strInOrOut = "IN YARD";
                    }

                    blnItemEntered = false;

                    for (intAutomileCounter = 0; intAutomileCounter <= intAutomileNumberOfRecords; intAutomileCounter++)
                    {
                        strAutomileVehicleNumber = TheAutomileImportDataSet.automileimport[intAutomileCounter].VehicleNumber;   

                        blnItemFound = strAutomileVehicleNumber.Contains(strVehicleNumber);

                        if (blnItemFound == true)
                        {
                            strInOrOut = TheAutomileImportDataSet.automileimport[intAutomileCounter].IsInside;

                            if (gintThirdCounter > 0)
                            {
                                for (intThirdCounter = 0; gintThirdCounter <= gintThirdNumberOfRecords; intThirdCounter++)
                                {
                                    if (strVehicleNumber == TheVehicleCurrentStatusDataSet.vehiclecurrentstatus[intThirdCounter].VehicleNumber)
                                    {
                                        if(strInOrOut == "YES")
                                        {
                                            TheVehicleCurrentStatusDataSet.vehiclecurrentstatus[intThirdCounter].InOrOut = "IN YARD";
                                        }
                                        else if(strInOrOut == "NO")
                                        {
                                            TheVehicleCurrentStatusDataSet.vehiclecurrentstatus[intThirdCounter].InOrOut = strInOrOut;
                                        }                                        

                                        blnItemEntered = true;
                                    }
                                }
                            }
                        }
                        
                    }


                    intRecordsReturned = TheFindVehicleInShopByVehicleIDDataSet.FindVehicleMainInShopByVehicleID.Rows.Count;

                    if (intRecordsReturned > 0)
                    {
                        strInOrOut = "IN SHOP";
                    }

                    if (blnItemEntered == false)
                    {
                        VehicleCurrentStatusDataSet.vehiclecurrentstatusRow NewVehicleRow = TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.NewvehiclecurrentstatusRow();

                        NewVehicleRow.AssignedOffice = strAssignedOffice;
                        NewVehicleRow.Manager = strManager;
                        NewVehicleRow.Driver = strFullName;
                        NewVehicleRow.InOrOut = strInOrOut;
                        NewVehicleRow.VehicleNumber = strVehicleNumber;

                        TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.Rows.Add(NewVehicleRow);

                        ShortedCurrentVehicleDataSet.shortenlistRow SecondVehicleRow = TheShortedCurrentVehicleDataSet.shortenlist.NewshortenlistRow();

                        SecondVehicleRow.Manager = strManager;
                        SecondVehicleRow.VehicleNumber = strVehicleNumber;
                        SecondVehicleRow.InOrOut = strInOrOut;

                        TheShortedCurrentVehicleDataSet.shortenlist.Rows.Add(SecondVehicleRow);
                    }
                }

                dgrResults.ItemsSource = TheVehicleCurrentStatusDataSet.vehiclecurrentstatus;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Four Thirty Report // Main Window // Process Expander " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void ExportFirstSheet()
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.Rows.Count;
                intColumnNumberOfRecords = TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheVehicleCurrentStatusDataSet.vehiclecurrentstatus.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Four Thirty Report // Main Window // Export First Sheet " + ex.Message);

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
        private void ExportSecondSheet()
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheShortedCurrentVehicleDataSet.shortenlist.Rows.Count;
                intColumnNumberOfRecords = TheShortedCurrentVehicleDataSet.shortenlist.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheShortedCurrentVehicleDataSet.shortenlist.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheShortedCurrentVehicleDataSet.shortenlist.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Four Thirty Report // Main Window // Export Second Sheet " + ex.Message);

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        private void expExportExcel_Expanded(object sender, RoutedEventArgs e)
        {
            ExportFirstSheet();

            ExportSecondSheet();
        }
    }
}

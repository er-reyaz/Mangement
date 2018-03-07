using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using ex = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ExceptionManagement.Controllers
{
   // [Authorize]
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //try
            //{
            //    System.Data.OleDb.OleDbConnection MyConnection;
            //    System.Data.DataSet DtSet;
            //    System.Data.OleDb.OleDbDataAdapter MyCommand;
            //    MyConnection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Reyaz\Desktop\IPMS.xlsx';Extended Properties=Excel 8.0;");
            //    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [EmployeeBilling$]", MyConnection);
            //    MyCommand.TableMappings.Add("Table", "TestTable");
            //    DtSet = new System.Data.DataSet();
            //    MyCommand.Fill(DtSet);
            //   // dataGridView1.DataSource = DtSet.Tables[0];
            //    MyConnection.Close();
            //}
            //catch (Exception ex)
            //{
            //   // MessageBox.Show(ex.ToString());
            //}
            string path = @"C:\Users\Reyaz\Desktop\IPMS.xlsx";
            exceldata(path);

            return View();
        }
        public static System.Data.DataTable exceldata(string filePath)
        {
            DataTable dtexcel = new DataTable("EmployeeBilling");
            bool hasHeaders = true;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                new object[] { null, null, null, "TABLE" });
            //Looping Total Sheet of Xl File
            /*foreach (DataRow schemaRow in schemaTable.Rows)
            {
            }*/
            //Looping a first Sheet of Xl File
            DataRow schemaRow = schemaTable.Rows[0];
            string sheet = schemaRow["TABLE_NAME"].ToString();
            if (!sheet.EndsWith("_"))
            {
                string query = "SELECT  * FROM [EmployeeBilling$]";
                OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                dtexcel.Locale = CultureInfo.CurrentCulture;
                daexcel.Fill(dtexcel);
            }

            conn.Close();
            DataTable actualTable = new DataTable();
            DataTable myTable = new DataTable();
            myTable = dtexcel.Copy();
            if (dtexcel.Columns.Count >= 5)
            {
                for (int i = 5; i < dtexcel.Columns.Count; i++)
                {
                    // DataColumn dc = new DataColumn(((System.Data.DataColumn)dtexcel.Columns[i]).Caption);
                    // actualTable.Columns.Add(dc);
                    myTable.Columns.RemoveAt(5);
                }
                
            }
            var projects = myTable.DefaultView.ToTable(true, ((DataColumn)dtexcel.Columns[2]).Caption);
            List<Sheet> list = new List<Sheet>();
            // var cols = dtexcel.Columns.Cast<System.Data.DataColumn>().Take(5);
            foreach (DataRow dr in projects.Rows)
            {
                Sheet exSheet = new Sheet();
                string columnProject = myTable.Columns[2].Caption;
                string columnToSum = myTable.Columns[4].Caption;
                string projectName = Convert.ToString(dr[0]);
                string  expression = columnProject + "='"+ projectName + "'";
                var rowArray = myTable.Select(expression);
                DataTable sheetTable = myTable.Clone();
                foreach (DataRow row in rowArray)
                {
                    sheetTable.ImportRow(row);
                }
                object sum= sheetTable.Compute("Sum(["+ columnToSum + "])", "");
                exSheet.SheetName = projectName;
                exSheet.SheetTable = sheetTable;
                exSheet.SheetSum = sum;
                list.Add(exSheet);
            }
            //--------------------- generte excel
            ex.Application ExcelApp = new ex.Application();

            ex.Workbook ExcelWorkBook = null;
            ex.Worksheet ExcelWorkSheet = null;

            ExcelApp.Visible = true;
            ExcelWorkBook = ExcelApp.Workbooks.Add(ex.XlWBATemplate.xlWBATWorksheet);

            try
            {
                for (int i = 1; i < list.Count; i++)
                    ExcelWorkBook.Worksheets.Add(); //Adding New sheet in Excel Workbook

                for (int i = 0; i < list.Count; i++)
                {
                    int r = 1; // Initialize Excel Row Start Position  = 1

                    ExcelWorkSheet = ExcelWorkBook.Worksheets[i + 1];

                    //Writing Columns Name in Excel Sheet

                    for (int col = 1; col <= list[i].SheetTable.Columns.Count; col++)
                        ExcelWorkSheet.Cells[r, col] = list[i].SheetTable.Columns[col - 1].ColumnName;
                    r++;

                    //Writing Rows into Excel Sheet
                    for (int row = 0; row < list[i].SheetTable.Rows.Count; row++) //r stands for ExcelRow and col for ExcelColumn
                    {
                        // Excel row and column start positions for writing Row=1 and Col=1
                        for (int col = 1; col <= list[i].SheetTable.Columns.Count; col++)
                            ExcelWorkSheet.Cells[r, col] = list[i].SheetTable.Rows[row][col - 1].ToString();
                        r++;
                    }
                    ExcelWorkSheet.Name = "sheet"+i;//Renaming the ExcelSheets

                }
              
                ExcelWorkBook.SaveAs(@"C:\Users\Reyaz\Desktop\IPMS_test.xlsx");
                ExcelWorkBook.Close();
                ExcelApp.Quit();
                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);
            }
            catch (Exception exHandle)
            {

                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
            finally
            {

                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }
            return dtexcel;
           
        }

        private class Sheet
        {
           public string SheetName { get; set; }
           public DataTable SheetTable { get; set; }
           public object SheetSum { get; set; }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
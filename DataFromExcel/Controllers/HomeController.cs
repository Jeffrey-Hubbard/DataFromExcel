using ClosedXML.Excel;
using DataFromExcel.Models;
using DataFromExcel.ViewModels;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DataFromExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult DownloadFile(string fileName)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Content/";
            byte[] fileBytes = System.IO.File.ReadAllBytes(path + fileName);
            //string fileName = "SalesTemplate.xlsx";
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        [HttpPost]
        public ActionResult ImportFile(HttpPostedFileBase file)
        {

                string path = Server.MapPath("~/Content/Upload/" + file.FileName);
                file.SaveAs(path);

                string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;IMEX=1'";

                using (SalesAnalysisEntities db = new SalesAnalysisEntities())
                {
                    DataSet ds = new DataSet(); // for storage of imported excep data
                    var sales = new List<Sale>();

                //using (OleDbConnection excelConnection = new OleDbConnection(excelConnectionString))
                using (var workBook = new XLWorkbook(path))
                {

                    //excelConnection.Open();

                    //    OleDbCommand cmd = new OleDbCommand("Select * from [Sales]", excelConnection);

                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    var dt = new DataTable();

                    //loop through the worksheet rows
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        //use the first row to define columns in DataTable
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            // add rows to datatable
                            dt.Rows.Add();
                            int i = 0;
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }                   
                    }

                    ds.Tables.Add(dt);

                    }

                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        if (dr[0] == null || dr[0] == "")
                        {
                            break;
                        }
                        Sale sale = new Sale();
                        sale.Date =  DateTime.Parse( dr.Field<string>("Date"));
                        sale.StoreNumber = Convert.ToInt32( dr.Field<string>("Store Number"));
                        sale.Person = dr.Field<string>("Person");
                        sale.Item = dr.Field<string>("Item");
                        sale.Units = Convert.ToInt32( dr.Field<string>("Units"));
                        sale.UnitCost = Convert.ToInt32( dr.Field<string>("Unit Cost"));
                        sale.Total = Convert.ToInt32( dr.Field<string>("Total"));
                        db.Sales.Add(sale);
                        
                    } 


                    db.SaveChanges();
                     
                }

                ViewBag.Result = "Successfully Imported";

            return RedirectToAction("Index", "Analysis");
        }
    }
}
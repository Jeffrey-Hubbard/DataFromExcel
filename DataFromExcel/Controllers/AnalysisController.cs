using DataFromExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DataFromExcel.ViewModels;

namespace DataFromExcel.Controllers
{
    public class AnalysisController : Controller
    {
        // GET: Analysis
        public ActionResult Index()
        {
            var analysisResultVm = new AnalysisResultsVm();
            DateTime firstOfCurrentMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime firstOfLastMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, 1);

            // List of all products (no duplicates)
            using (var context = new SalesAnalysisEntities())
            {
                var result = context.Sales.GroupBy(x => x.Item).Select(y => y.FirstOrDefault().Item).ToList();
                analysisResultVm.ProductList = result;
            }


            // Total sales
            using (var context = new SalesAnalysisEntities())
            {
                var result = context.Sales.Sum(x => x.Total);
                analysisResultVm.SalesTotal = result;
            }


            // current month's sales
            using (var context = new SalesAnalysisEntities())
            {
                
                var result = context.Sales.Where(y => y.Date >= firstOfCurrentMonth ).Sum(x => x.Total);
                analysisResultVm.CurrentMoSalesTotal = result;
                analysisResultVm.CurrentMoSalesPace = (result / DateTime.Now.Day) * DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);

            }

            // Last month's sales total
            using (var context = new SalesAnalysisEntities())
            {
                
                var result = context.Sales.Where(y => y.Date >= firstOfLastMonth && y.Date < firstOfCurrentMonth).Sum(x => x.Total);
                analysisResultVm.LastMoSalesTotal = result;
            }

            return View(analysisResultVm);
        }




    }
}
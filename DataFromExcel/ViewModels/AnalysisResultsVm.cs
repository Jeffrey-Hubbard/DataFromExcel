using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataFromExcel.ViewModels
{
    public class AnalysisResultsVm
    {
        public List<string> ProductList { get; set; }
        public decimal? SalesTotal { get; set; }
        public decimal? CurrentMoSalesTotal { get; set; }
        public decimal? CurrentMoSalesPace { get; set; }
        public decimal? LastMoSalesTotal { get; set; }


    }
}
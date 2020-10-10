using LinqToExcel;
using LinqToExcel.Attributes;
using LinqToExcel.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ReadExcel.Controllers
{
    public class HomeController : Controller
    {
        public class Sheet1
        {
            [ExcelColumn("Number1")]
            public decimal? Number1 { get; set; }
            [ExcelColumn("Number2")]
            public decimal? Number2 { get; set; }
            [ExcelColumn("Total")]
            public decimal? Total { get; set; }
        }


        public ActionResult Index()
        {
            var excel = new ExcelQueryFactory(@"D:\Test01.xlsx");
            var test1 = (from c in excel.Worksheet<Sheet1>()
                         select c).ToList();

            var test2 = (from c in excel.WorksheetRangeNoHeader("C3", "C3") select c).ToList();


            return View();
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
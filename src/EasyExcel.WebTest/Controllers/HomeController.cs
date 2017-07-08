using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EasyExcel.Export;

namespace EasyExcel.WebTest.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
          
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            var users = EasyExcel.Test.Models.User.MakeUsers();
            var sheets = new List<EESheet> { new EESheet(users), new EESheet(users, "user data") };
            using (var export = new EEExportor(sheets))
            {
                export.StreamAction();
            }

            return View();
        }
    }
}
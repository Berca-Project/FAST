using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class TaxStampController : Controller
    {
        // GET: TaxStamp
        public ActionResult Index()
        {
            return View();
        }

        // GET: TaxStamp/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: TaxStamp/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: TaxStamp/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: TaxStamp/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: TaxStamp/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: TaxStamp/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: TaxStamp/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}

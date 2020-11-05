using Fast.Web.Models;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class CRRController : BaseController<MppModel>
    {
        // GET: CRR
        [CustomAuthorize("reportcrr")]
        public ActionResult Index()
        {
            return View();
        }

        // GET: CRR/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: CRR/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: CRR/Create
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

        // GET: CRR/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: CRR/Edit/5
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

        // GET: CRR/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: CRR/Delete/5
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

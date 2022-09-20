using BuisnessLogicLayer;
using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebDoc.Controllers
{
    public class HomeController : Controller
    {

        BuisnessController bc = new BuisnessController();

        public ActionResult Index()
        {
            return View(bc.fileList);
        }

        public ActionResult GetTemplate()
        {
            return File(bc.GetTemplate(), "doc", "Example.docx");
        }

        public ActionResult GetTemplateById(Guid docTemplateId)
        {
            return File(bc.GetTemplateById(docTemplateId), "doc", "Template.docx");
        }

        public ActionResult CreateTemplate()
        {
            DocTemplate temp = new DocTemplate();
            return View(temp);
        }

        [HttpPost]
        public ActionResult CreateTemplate(DocTemplate input)
        {
            byte[] file = Helper.StreamToByteArray(input.FileInput.InputStream);
            bc.CreateTemplate(input.Title, file);

            return RedirectToAction("Index");
        }

        public ActionResult UseTemplate(Guid docTemplateId)
        {
            DocTemplate respont = bc.fileList.FirstOrDefault(x => x.DocTemplateId == docTemplateId);
            return View(respont);
        }

        [HttpPost]
        public ActionResult UseTemplate(DocTemplate input)
        {
            byte[] file = bc.CreateDocument(input);
            string filetype = (input.Pdf == true) ? ".pdf" : ".docx";

            return File(file, "doc", input.Title + filetype);
        }
    }
}
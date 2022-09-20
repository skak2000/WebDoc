using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Model
{
    public class DocTemplate
    {
        public DocTemplate()
        {
            DocTemplateId = Guid.NewGuid();
            CreateDate = DateTime.Now;
        }

        public Guid DocTemplateId { get; set; }
        public DateTime CreateDate { get; set; }
        public string Title { get; set; }
        public string Word1 { get; set; }
        public string Word2 { get; set; }
        public string Word3 { get; set; }
        public string Word4 { get; set; }
        public string Word5 { get; set; }
        public HttpPostedFileBase FileInput { get; set; }

        public bool Pdf { get; set; }

        public string FileTemplatePath
        {
            get 
            { 
                return (System.Web.Hosting.HostingEnvironment.MapPath("~") + "Templates\\" + DocTemplateId.ToString() + ".docx");
            }
        }

        public string FileTempPath
        {
            get
            {
                return (System.Web.Hosting.HostingEnvironment.MapPath("~") + "Temp\\" + DocTemplateId.ToString() + ".docx");
            }
        }

        // File to return to user
        public string OutputPath
        {
            get
            {
                string fileType = (Pdf == true) ? ".pdf" : ".docx";
                return (System.Web.Hosting.HostingEnvironment.MapPath("~") + "Temp\\" + DocTemplateId.ToString() + fileType);
            }
        }
    }
}

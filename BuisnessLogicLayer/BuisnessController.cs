using Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuisnessLogicLayer
{
    public class BuisnessController
    {
        public List<DocTemplate> fileList = new List<DocTemplate>();

        public BuisnessController()
        {
            if(File.Exists(System.Web.Hosting.HostingEnvironment.MapPath("~") + "Templates\\MyData.json") == false)
            {
                fileList = new List<DocTemplate>();
                CreateTemplate("Basic", GetTemplate());
            }

            string json = File.ReadAllText(System.Web.Hosting.HostingEnvironment.MapPath("~") + "Templates\\MyData.json");
            fileList = JsonConvert.DeserializeObject<List<DocTemplate>>(json);
        }

        public void CreateTemplate(string title, byte[] file)
        {
            DocTemplate template = new DocTemplate()
            {
                Title = title
            };

            // Creating the template in the template folder.
            File.WriteAllBytes(template.FileTemplatePath, file);
            fileList.Add(template);
            SaveData();
        }

        public byte[] CreateDocument(DocTemplate input)
        {
            // Creating a copy of the template, becurs SaveAs PDF will override the original docx.
            File.Copy(input.FileTemplatePath, input.FileTempPath);
            WordTool.GenerateLetter(input);
            
            // Getting the new created file
            byte[] respont = File.ReadAllBytes(input.OutputPath);

            // Clean up
            if(File.Exists(input.FileTempPath))
            {
                File.Delete(input.FileTempPath);                
            }
            if (File.Exists(input.OutputPath))
            {
                File.Delete(input.OutputPath);
            }
            return respont;
        }

        public byte[] GetTemplate()
        {
            byte[] respont = File.ReadAllBytes(System.Web.Hosting.HostingEnvironment.MapPath("~") + "Templates\\Basic.docx");
            return respont;
        }

        public byte[] GetTemplateById(Guid templateId)
        {
            DocTemplate temp = fileList.FirstOrDefault(x => x.DocTemplateId == templateId);
            byte[] respont = File.ReadAllBytes(temp.FileTemplatePath);
            return respont;
        }

        private void SaveData()
        {
            string json = JsonConvert.SerializeObject(fileList);
            File.WriteAllText(System.Web.Hosting.HostingEnvironment.MapPath("~") + "Templates\\MyData.json", json);
        }
    }
}

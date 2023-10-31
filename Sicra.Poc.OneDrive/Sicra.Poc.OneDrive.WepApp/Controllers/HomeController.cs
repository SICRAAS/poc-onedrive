using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Sicra.Poc.OneDrive.WepApp.BusinessLogic;
using Sicra.Poc.OneDrive.WepApp.Models;

namespace Sicra.Poc.OneDrive.WepApp.Controllers
{
    public class HomeController : Controller
    {
        private const string ExcelFileName = "mock.xlsx";
        private const string SiteId = "e0e0cf55-f46e-407a-9efa-b655129ef6b6";
        private const string FolderName = "_poc-onedrive";
        
        public ActionResult Index()
        {
            return View();
        }

        public async Task<ActionResult> NewFile()
        {
            
            var file = FileHandler.ReadMockFile(ExcelFileName);
            var url = await FileHandler.StoreOnSharepointSite(file, SiteId, FolderName);
            
            
            var model = new NewFileModel
            {
                Url = url
            };
            return new RedirectResult(model.Url);
        }
    }
}
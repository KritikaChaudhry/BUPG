using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using bupg_final.Models;
using Microsoft.AspNetCore.Http;
using System.IO;
using OfficeOpenXml;
using B_I.Models;

namespace B_I.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public async Task Import(IFormFile file)
        {
            var list = new List<AllocationSummary>();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var pkg = new ExcelPackage(stream))
                {
                    ExcelWorksheet excelWorksheet = pkg.Workbook.Worksheets["Base"];
                    var rows = excelWorksheet.Dimension.Rows;

                    for (int row = 2; row <= rows; row++)
                    {
                        list.Add(new AllocationSummary
                        {
                            Id = Convert.ToString(excelWorksheet.Cells[row, 1].Value),
                            GGID = int.Parse(Convert.ToString(excelWorksheet.Cells[row, 2].Value)),
                            BankId = Convert.ToString(excelWorksheet.Cells[row, 3].Value),
                            EmployeeName = Convert.ToString(excelWorksheet.Cells[row, 4].Value),
                            Source = Convert.ToString(excelWorksheet.Cells[row, 5].Value),
                            ProjectName = Convert.ToString(excelWorksheet.Cells[row, 6].Value),
                            ProjectNumber = Convert.ToString(excelWorksheet.Cells[row, 7].Value),
                            ProjectStartDate = Convert.ToString(excelWorksheet.Cells[row, 8].Value),
                            NTLoginId = Convert.ToString(excelWorksheet.Cells[row, 10].Value),
                            GlobalDateOfJoining = Convert.ToString(excelWorksheet.Cells[row, 11].Value),
                            Designation = Convert.ToString(excelWorksheet.Cells[row, 12].Value),
                            LocalGrade = Convert.ToString(excelWorksheet.Cells[row, 13].Value),
                            Location = Convert.ToString(excelWorksheet.Cells[row, 14].Value),
                            CurrentOU = Convert.ToString(excelWorksheet.Cells[row, 15].Value),
                            Practice = Convert.ToString(excelWorksheet.Cells[row, 16].Value),
                            SubPractice = Convert.ToString(excelWorksheet.Cells[row, 17].Value),
                            SupervisorFullName = Convert.ToString(excelWorksheet.Cells[row, 18].Value),
                            TowerHead = Convert.ToString(excelWorksheet.Cells[row, 19].Value),
                            ManualEntry = Convert.ToString(excelWorksheet.Cells[row, 20].Value),
                            EngagementType = Convert.ToString(excelWorksheet.Cells[row, 21].Value),
                            BGV = Convert.ToString(excelWorksheet.Cells[row, 22].Value),
                            BGVCompletionDate = Convert.ToString(excelWorksheet.Cells[row, 23].Value),
                            PrimarySkill = Convert.ToString(excelWorksheet.Cells[row, 24].Value),
                            MobileNo = Convert.ToString(excelWorksheet.Cells[row, 25].Value),
                            SCBEmailId = Convert.ToString(excelWorksheet.Cells[row, 26].Value),
                            EmailId = Convert.ToString(excelWorksheet.Cells[row, 27].Value),
                            BookingId = Convert.ToString(excelWorksheet.Cells[row, 28].Value),
                            On = Convert.ToString(excelWorksheet.Cells[row, 29].Value),
                            By = Convert.ToString(excelWorksheet.Cells[row, 30].Value),
                            Tower = Convert.ToString(excelWorksheet.Cells[row, 31].Value)
                        });
                    }
                }
            }
            var context = new SCBContext();
            context.AllocationSummary.AddRange(list);
            context.SaveChanges();
        }

        public IActionResult SOW_eRRFStatus()
        {
            return View();
        }

        public IActionResult UpdateSOW()
        {
            return View();
        }

        public IActionResult eRRF()
        {
            return View();
        }

        public IActionResult AssociateDetail()
        {
            return View();
        }

        public IActionResult OfferUpdate()
        {
            return View();
        }
        public IActionResult Exit()
        {
            return View();
        }
        public IActionResult Requirement()
        {
            return View();
        }
             public IActionResult UploadAllocation()
        {
            return View();
        }
           public IActionResult AllocationReport()
        {
            return View();
        }
       
            public IActionResult AllocationDetail()
        {
            return View();
        }
           public IActionResult VST()
        {
            return View();
        }

        public IActionResult Upload_eRRF()
        {
            return View();
        } 

        public IActionResult Upload()
        {
            return View();
        } 
       public IActionResult ELearning()
        {
            return View();
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

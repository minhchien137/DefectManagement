using System.Text;
using DefectManagement.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;

namespace DefectManagement.Controllers
{
    public class DefectController : Controller
    {
        private readonly ApplicationDbContext _context;

        public DefectController(ApplicationDbContext context, HttpClient httpClient)
        {
            _context = context;
        }

        // Get danh sách operation
        public IActionResult Create()
        {
            var operation = _context.SVN_quality_reason
                .Select(x => x.operation)
                .Distinct()
                .ToList();

            ViewBag.Operations = new SelectList(operation);
            return View();
        }

        // API lấy Code theo Operation
        [HttpGet]
        public JsonResult GetCodesByOperation(string operation)
        {
            var codes = _context.SVN_quality_reason
                .Where(x => x.operation == operation)
                .Select(x => new { x.code, x.name })
                .ToList();

            return Json(codes);
        }

        [HttpPost]
        public IActionResult Create([FromBody] SVN_Defect_Record_History model)
        {
            try
            {

                if (string.IsNullOrEmpty(model.Employer_code) ||
                    string.IsNullOrEmpty(model.Work_order) ||
                    string.IsNullOrEmpty(model.Item_code) ||
                    string.IsNullOrEmpty(model.Operation) ||
                    string.IsNullOrEmpty(model.Defect_Code) ||
                    model.Qty_NG <= 0)
                {
                    return Json(new { success = false, message = "Vui lòng điền đầy đủ thông tin bắt buộc!" });
                }

                var historyRecord = new SVN_Defect_Record_History
                {
                    Work_order = model.Work_order,
                    Item_code = model.Item_code,
                    Defect_Code = model.Defect_Code,
                    Qty_NG = model.Qty_NG,
                    INSDatetime = DateTime.Now.ToString("yyyyMMdd"),
                    Operation = model.Operation,
                    Employer_code = model.Employer_code,
                    Employer_name = model.Employer_name ?? "",
                    Note = model.Note
                };

                // lưu vào bảng Defect_History
                _context.SVN_Defect_Record_History.Add(historyRecord);
                _context.SaveChanges();
                Console.WriteLine("History saved successfully");

                var insertSql = @"INSERT INTO SVN_Defect_Record_Copy 
                                (Item_code, Defect_Code, Qty_NG, INSDatetime, Operation, Employer_code, Employer_name) 
                                VALUES ({0}, {1}, {2}, {3}, {4}, {5}, {6})";

                // lưu vào bảng Defect_Record
                _context.Database.ExecuteSqlRaw(insertSql,
                    historyRecord.Item_code,
                    historyRecord.Defect_Code,
                    historyRecord.Qty_NG,
                    historyRecord.INSDatetime,
                    historyRecord.Operation,
                    historyRecord.Employer_code,
                    historyRecord.Employer_name ?? "");

                Console.WriteLine("Copy saved successfully");

                return Json(new { success = true, message = "Lưu thông tin thành công!", data = model });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Có lỗi xảy ra: {ex.Message}" });
            }
        }


    }
}


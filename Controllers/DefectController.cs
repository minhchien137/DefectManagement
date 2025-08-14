using System.Text;
using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;
using DefectManagement.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;

namespace DefectManagement.Controllers
{
    public class DefectController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public DefectController(ApplicationDbContext context, IWebHostEnvironment webHostEnvironment)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
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


        // Xuất File Excel
        public async Task<IActionResult> ExportToExcel(string workOrder = "", string itemCode = "", string employerCode = "", string operation = "", string fromInsDateTime = "", string toInsDateTime = "")
        {
            var query = _context.SVN_Defect_Record_History.AsQueryable();

            if (!string.IsNullOrEmpty(workOrder))
                query = query.Where(x => x.Work_order.Contains(workOrder));

            if (!string.IsNullOrEmpty(itemCode))
                query = query.Where(x => x.Item_code.Contains(itemCode));

            if (!string.IsNullOrEmpty(employerCode))
                query = query.Where(x => x.Employer_code.Contains(employerCode));

            if (!string.IsNullOrEmpty(operation))
                query = query.Where(x => x.Operation.Contains(operation));

            if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
            {
                var formattedDate = fromDate.ToString("yyyyMMdd");
                query = query.Where(x => x.INSDatetime.CompareTo(formattedDate) >= 0);
            }
            if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
            {
                var formattedDate = toDate.ToString("yyyyMMdd");
                query = query.Where(x => x.INSDatetime.CompareTo(formattedDate) <= 0);
            }


            var data = await query.OrderByDescending(x => x.Time_line).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("DefectHistory");
                var currentRow = 1;

                // Font mặc định
                ws.Style.Font.FontName = "Arial";
                ws.Style.Font.FontSize = 11;

                // Header
                string[] headers = { "ID", "Work Order", "Item Code", "Defect Code", "Qty NG", "INS DateTime", "Operation", "Employer Code", "Employer Name", "Note", "Image Error", "Time Line" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(currentRow, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.Green;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }

                // Data
                foreach (var item in data)
                {
                    currentRow++;
                    ws.Cell(currentRow, 1).Value = item.Id;
                    ws.Cell(currentRow, 2).Value = item.Work_order;
                    ws.Cell(currentRow, 3).Value = item.Item_code;
                    ws.Cell(currentRow, 4).Value = item.Defect_Code;
                    ws.Cell(currentRow, 5).Value = item.Qty_NG;
                    ws.Cell(currentRow, 6).Value = item.INSDatetime;
                    ws.Cell(currentRow, 7).Value = item.Operation;
                    ws.Cell(currentRow, 8).Value = item.Employer_code;
                    ws.Cell(currentRow, 9).Value = item.Employer_name;
                    ws.Cell(currentRow, 10).Value = item.Note;
                    ws.Cell(currentRow, 11).Value = item.Image_error;
                    ws.Cell(currentRow, 12).Value = item.Time_line?.ToString("yyyy-MM-dd HH:mm:ss");
                }

                // Canh giữa các cột số và ngày
                ws.Columns(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // ID
                ws.Columns(5, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Qty NG
                ws.Columns(6, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // INS DateTime
                ws.Columns(12, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Time Line

                // Auto fit column width
                ws.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "DefectHistory.xlsx");
                }
            }
        }


        public async Task<IActionResult> Result(string workOrder = "", string itemCode = "", string employerCode = "", string operation = "", string fromInsDateTime = "", string toInsDateTime = "")
        {
            try
            {
                var sql = "SELECT * FROM SVN_Defect_Record_History WHERE 1=1";
                var parameters = new List<object>();

                if (!string.IsNullOrEmpty(workOrder))
                {
                    sql += " AND (Work_order LIKE {" + parameters.Count + "} OR Work_order IS NULL)";
                    parameters.Add($"%{workOrder}%");
                }

                if (!string.IsNullOrEmpty(itemCode))
                {
                    sql += " AND (Item_code LIKE {" + parameters.Count + "} OR Item_code IS NULL)";
                    parameters.Add($"%{itemCode}%");
                }

                if (!string.IsNullOrEmpty(employerCode))
                {
                    sql += " AND (Employer_code LIKE {" + parameters.Count + "} OR Employer_code IS NULL)";
                    parameters.Add($"%{employerCode}%");
                }

                if (!string.IsNullOrEmpty(operation))
                {
                    sql += " AND (Operation LIKE {" + parameters.Count + "} OR Operation IS NULL)";
                    parameters.Add($"%{operation}%");
                }

                // Logic lọc ngày tháng cũ
                // ĐÃ SỬA ĐỔI: Logic lọc ngày tháng theo cột INSDatetime (chuỗi)
                if (!string.IsNullOrEmpty(fromInsDateTime) && DateTime.TryParse(fromInsDateTime, out var fromDate))
                {
                    var formattedDate = fromDate.ToString("yyyyMMdd");
                    sql += " AND (INSDatetime >= {" + parameters.Count + "})";
                    parameters.Add(formattedDate);
                }

                if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                {
                    var formattedDate = toDate.ToString("yyyyMMdd");
                    sql += " AND (INSDatetime <= {" + parameters.Count + "})";
                    parameters.Add(formattedDate);
                }

                sql += " ORDER BY Time_line DESC";

                var results = await _context.SVN_Defect_Record_History
                    .FromSqlRaw(sql, parameters.ToArray())
                    .AsNoTracking()
                    .ToListAsync();

                // Truyền giá trị filter ra View
                ViewBag.WorkOrder = workOrder ?? "";
                ViewBag.ItemCode = itemCode ?? "";
                ViewBag.EmployerCode = employerCode ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.FromInsDateTime = fromInsDateTime ?? "";
                ViewBag.ToInsDateTime = toInsDateTime ?? "";
                return View(results);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi: {ex.Message}";
                ViewBag.WorkOrder = workOrder ?? "";
                ViewBag.ItemCode = itemCode ?? "";
                ViewBag.EmployerCode = employerCode ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.FromInsDateTime = fromInsDateTime ?? "";
                ViewBag.ToInsDateTime = toInsDateTime ?? "";

                return View(new List<SVN_Defect_Record_History>());
            }
        }


        [HttpPost]
        public async Task<IActionResult> Create(SVN_Defect_Record_History model, IFormFile imageFile)
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

                string imagePath = null;

                // Xử lý upload ảnh nếu có
                if (imageFile != null && imageFile.Length > 0)
                {
                    // Kiểm tra định dạng file
                    var allowedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
                    var fileExtension = Path.GetExtension(imageFile.FileName).ToLower();

                    if (!allowedExtensions.Contains(fileExtension))
                    {
                        return Json(new { success = false, message = "Chỉ cho phép upload ảnh với định dạng: jpg, jpeg, png, gif, bmp" });
                    }

                    // Kiểm tra kích thước file (5MB)
                    if (imageFile.Length > 5 * 1024 * 1024)
                    {
                        return Json(new { success = false, message = "Kích thước ảnh không được vượt quá 5MB" });
                    }

                    // Tạo thư mục uploads/defect-images nếu chưa có
                    var uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "uploads", "defect-images");
                    if (!Directory.Exists(uploadsFolder))
                    {
                        Directory.CreateDirectory(uploadsFolder);
                    }

                    // Tạo tên file unique
                    var fileName = $"{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}{fileExtension}";
                    var filePath = Path.Combine(uploadsFolder, fileName);

                    // Lưu file
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await imageFile.CopyToAsync(stream);
                    }

                    // Lưu đường dẫn relative để hiển thị trên web
                    imagePath = $"/uploads/defect-images/{fileName}";
                }

                // Gọi stored procedure với parameters
                await _context.Database.ExecuteSqlRawAsync(
                    "EXEC [dbo].[SVN_InsertDefectReport] {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}",
                    model.Work_order ?? "",
                    model.Item_code ?? "",
                    model.Defect_Code ?? "",
                    model.Qty_NG,
                    DateTime.Now.ToString("yyyyMMdd"),
                    model.Operation ?? "",
                    model.Employer_code ?? "",
                    model.Employer_name ?? "",
                    model.Note ?? "",
                    imagePath ?? "",
                    DateTime.Now);

                Console.WriteLine("Stored procedure executed successfully");

                return Json(new { success = true, message = "Lưu thông tin thành công!", data = model });
            }

            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Có lỗi xảy ra: {ex.Message}" });
            }
        }


    }
}


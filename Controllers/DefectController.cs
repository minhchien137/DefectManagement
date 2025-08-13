using System.Text;
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
                    "EXEC [dbo].[SVN_InsertDefectReport] {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}",
                    model.Work_order ?? "",
                    model.Item_code ?? "",
                    model.Defect_Code ?? "",
                    model.Qty_NG,
                    DateTime.Now.ToString("yyyyMMdd"),
                    model.Operation ?? "",
                    model.Employer_code ?? "",
                    model.Employer_name ?? "",
                    model.Note ?? "",
                    imagePath ?? "");

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


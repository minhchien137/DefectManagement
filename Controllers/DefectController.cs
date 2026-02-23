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
        public async Task<IActionResult> ExportToExcel(string workOrder = "", string defectCode = "", string defectName = "", string employerCode = "", string operation = "", string fromInsDateTime = "", string toInsDateTime = "")
        {
            var query = _context.SVN_Defect_Record_History.AsQueryable();

            if (!string.IsNullOrEmpty(workOrder))
                query = query.Where(x => x.Work_order.Contains(workOrder));

            if (!string.IsNullOrEmpty(defectCode))
                query = query.Where(x => x.Defect_Code.Contains(defectCode));

            if (!string.IsNullOrEmpty(defectName))
                query = query.Where(x => x.Defect_Name.Contains(defectName));

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

            // Sắp xếp bản ghi theo thời gian ASC
            var data = await query.OrderBy(x => x.Time_line).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("DefectHistory");
                var currentRow = 1;

                // Font mặc định
                ws.Style.Font.FontName = "Times New Roman";
                ws.Style.Font.FontSize = 11;

                // Header
                string[] headers = { "ID", "Work Order", "Item Code", "Defect Code", "Defect Name", "Qty NG", "INS DateTime", "Operation", "Employer Code", "Employer Name", "Note", "Image Error", "Time Line" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(currentRow, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.5);
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
                    ws.Cell(currentRow, 5).Value = item.Defect_Name;
                    ws.Cell(currentRow, 6).Value = item.Qty_NG;
                    ws.Cell(currentRow, 7).Value = item.INSDatetime;
                    ws.Cell(currentRow, 8).Value = item.Operation;
                    ws.Cell(currentRow, 9).Value = item.Employer_code;
                    ws.Cell(currentRow, 10).Value = item.Employer_name;
                    ws.Cell(currentRow, 11).Value = item.Note;
                    ws.Cell(currentRow, 12).Value = item.Image_error;
                    ws.Cell(currentRow, 13).Value = item.Time_line?.ToString("yyyy-MM-dd HH:mm:ss");
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


        // Updated Result method with server-side pagination
        public async Task<IActionResult> Result(string workOrder = "", string defectCode = "", string defectName = "",
            string employerCode = "", string operation = "", string fromInsDateTime = "", string toInsDateTime = "",
            int page = 1, int pageSize = 25)
        {
            try
            {
                var query = _context.SVN_Defect_Record_History.AsQueryable();

                // Apply filters
                if (!string.IsNullOrEmpty(workOrder))
                    query = query.Where(x => x.Work_order.Contains(workOrder));

                if (!string.IsNullOrEmpty(defectCode))
                    query = query.Where(x => x.Defect_Code.Contains(defectCode));

                if (!string.IsNullOrEmpty(defectName))
                    query = query.Where(x => x.Defect_Name.Contains(defectName));

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

                // Get total count for pagination
                var totalRecords = await query.CountAsync();

                // Apply sorting and pagination
                var results = await query
                    .OrderByDescending(x => x.Time_line)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .AsNoTracking()
                    .ToListAsync();

                // Calculate pagination info
                var totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);

                // Pass pagination data to view
                ViewBag.CurrentPage = page;
                ViewBag.TotalPages = totalPages;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalRecords = totalRecords;
                ViewBag.HasPreviousPage = page > 1;
                ViewBag.HasNextPage = page < totalPages;

                // Truyền giá trị filter ra View
                ViewBag.WorkOrder = workOrder ?? "";
                ViewBag.DefectCode = defectCode ?? "";
                ViewBag.DefectName = defectName ?? "";
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
                ViewBag.DefectCode = defectCode ?? "";
                ViewBag.DefectName = defectName ?? "";
                ViewBag.EmployerCode = employerCode ?? "";
                ViewBag.Operation = operation ?? "";
                ViewBag.FromInsDateTime = fromInsDateTime ?? "";
                ViewBag.ToInsDateTime = toInsDateTime ?? "";

                // Set default pagination values for error case
                ViewBag.CurrentPage = 1;
                ViewBag.TotalPages = 0;
                ViewBag.PageSize = pageSize;
                ViewBag.TotalRecords = 0;
                ViewBag.HasPreviousPage = false;
                ViewBag.HasNextPage = false;

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
                    "EXEC [dbo].[SVN_InsertDefectReport] {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}",
                    model.Work_order ?? "",
                    model.Item_code ?? "",
                    model.Defect_Code ?? "",
                    model.Defect_Name ?? "",
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


        /************************** Báo cáo Defect **************************/

        [HttpGet]
        public async Task<IActionResult> Report(string fromDate = "",string toDate = "",string workOrder = "",string operation = "",string defectCode = "",string itemCode = "",int page = 1,int pageSize = 10)
        {
            try
            {
                // ─── 1. Lấy toàn bộ dữ liệu không lọc để populate dropdown ───
                var allRaw = await _context.SVN_Defect_Record_History
                    .AsNoTracking()
                    .ToListAsync();

                // Dropdown lists (distinct, sorted)
                ViewBag.AllWorkOrders = allRaw.Select(x => x.Work_order ?? "").Distinct().OrderBy(x => x).Where(x => x != "").ToList();
                ViewBag.AllOperations = allRaw.Select(x => x.Operation ?? "").Distinct().OrderBy(x => x).Where(x => x != "").ToList();
                ViewBag.AllItemCodes = allRaw.Select(x => x.Item_code ?? "").Distinct().OrderBy(x => x).Where(x => x != "").ToList();

                // Defect Code list: distinct by code, kèm Defect Name để hiện "R14 - Lost roughness"
                var defectCodeDict = allRaw
                    .Where(x => !string.IsNullOrEmpty(x.Defect_Code))
                    .GroupBy(x => x.Defect_Code)
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(x => x.Defect_Name ?? "").FirstOrDefault(n => n != "") ?? ""
                    );
                ViewBag.AllDefectCodes = defectCodeDict.Keys.OrderBy(x => x).ToList();
                ViewBag.AllDefectNames = defectCodeDict;  // Dictionary<string,string> code → name

                // ─── 2. Query có filter ───
                var query = _context.SVN_Defect_Record_History.AsQueryable();

                if (!string.IsNullOrEmpty(workOrder))
                    query = query.Where(x => x.Work_order == workOrder);

                if (!string.IsNullOrEmpty(operation))
                    query = query.Where(x => x.Operation == operation);

                if (!string.IsNullOrEmpty(defectCode))
                    query = query.Where(x => x.Defect_Code == defectCode);

                if (!string.IsNullOrEmpty(itemCode))
                    query = query.Where(x => x.Item_code == itemCode);

                if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParse(fromDate, out var parsedFrom))
                {
                    var formatted = parsedFrom.ToString("yyyyMMdd");
                    query = query.Where(x => x.INSDatetime.CompareTo(formatted) >= 0);
                }

                if (!string.IsNullOrEmpty(toDate) && DateTime.TryParse(toDate, out var parsedTo))
                {
                    var formatted = parsedTo.ToString("yyyyMMdd");
                    query = query.Where(x => x.INSDatetime.CompareTo(formatted) <= 0);
                }

                // ─── 3. Lấy toàn bộ dữ liệu filtered cho KPI và Charts ───
                var allData = await query.AsNoTracking().ToListAsync();

                // KPI
                var totalOperations = allData.Select(x => x.Operation).Distinct().Count();
                var totalRecords = allData.Count;
                var totalQtyNG = allData.Sum(x => x.Qty_NG ?? 0);

                // Chart 1 – Defect by Operation (raw, gộp ở JS)
                var byOperation = allData
                    .GroupBy(x => x.Operation ?? "N/A")
                    .Select(g => new
                    {
                        Operation = g.Key,
                        TotalRecords = g.Count(),
                        TotalQtyNG = g.Sum(x => x.Qty_NG ?? 0)
                    })
                    .OrderByDescending(x => x.TotalQtyNG)
                    .ToList();

                // Chart 2 – Pareto by DefectCode + DefectName (raw, gộp ở JS)
                var byDefectName = allData
                    .GroupBy(x => new { x.Defect_Code, x.Defect_Name })
                    .Select(g => new
                    {
                        DefectCode = g.Key.Defect_Code ?? "",
                        DefectName = g.Key.Defect_Name ?? "N/A",
                        Count = g.Count(),
                        TotalQtyNG = g.Sum(x => x.Qty_NG ?? 0)
                    })
                    .OrderByDescending(x => x.TotalQtyNG)
                    .ToList();

                // Cumulative % Pareto
                var totalNG = byDefectName.Sum(x => x.TotalQtyNG);
                double cumulative = 0;
                var paretoData = byDefectName.Select(x =>
                {
                    cumulative += totalNG > 0 ? (double)x.TotalQtyNG / totalNG * 100 : 0;
                    return new
                    {
                        x.DefectCode,
                        x.DefectName,
                        x.Count,
                        x.TotalQtyNG,
                        CumulativePct = Math.Round(cumulative, 1)
                    };
                }).ToList();

                // Chart 3 – Pie: Top 10 Operations (dùng byOperation, gộp ở JS)
                // Dữ liệu pie sẽ dùng lại opLabels/opQtyNG từ ChartOperationLabels/ChartOperationQtyNG trong JS

                // Chart 4 – Daily trend
                var dailyTrend = allData
                    .Where(x => x.Time_line.HasValue)
                    .GroupBy(x => x.Time_line.Value.Date)
                    .Select(g => new
                    {
                        Date = g.Key.ToString("dd/MM"),
                        DateFull = g.Key,
                        QtyNG = g.Sum(x => x.Qty_NG ?? 0),
                        Count = g.Count()
                    })
                    .OrderBy(x => x.DateFull)
                    .ToList();

                // Moving average 7 ngày
                var qtyNGList = dailyTrend.Select(x => (double)x.QtyNG).ToList();
                var movingAvg = new List<double?>();
                for (int i = 0; i < qtyNGList.Count; i++)
                {
                    if (i < 6) movingAvg.Add(null);
                    else movingAvg.Add(Math.Round(qtyNGList.Skip(i - 6).Take(7).Average(), 1));
                }

                // ─── 4. Serialize chart data ───
                ViewBag.ChartOperationLabels = System.Text.Json.JsonSerializer.Serialize(byOperation.Select(x => x.Operation));
                ViewBag.ChartOperationQtyNG = System.Text.Json.JsonSerializer.Serialize(byOperation.Select(x => x.TotalQtyNG));
                ViewBag.ChartOperationCounts = System.Text.Json.JsonSerializer.Serialize(byOperation.Select(x => x.TotalRecords));

                ViewBag.ParetoLabels = System.Text.Json.JsonSerializer.Serialize(paretoData.Select(x => $"{x.DefectCode} - {x.DefectName}"));
                ViewBag.ParetoQtyNG = System.Text.Json.JsonSerializer.Serialize(paretoData.Select(x => x.TotalQtyNG));
                ViewBag.ParetoCumulative = System.Text.Json.JsonSerializer.Serialize(paretoData.Select(x => x.CumulativePct));

                // PieLabels/PieData không còn dùng ở View (dùng opMerged trong JS), nhưng giữ để tránh lỗi
                ViewBag.PieLabels = System.Text.Json.JsonSerializer.Serialize(byOperation.Take(10).Select(x => x.Operation));
                ViewBag.PieData = System.Text.Json.JsonSerializer.Serialize(byOperation.Take(10).Select(x => x.TotalQtyNG));

                ViewBag.DailyDates = System.Text.Json.JsonSerializer.Serialize(dailyTrend.Select(x => x.Date));
                ViewBag.DailyQtyNG = System.Text.Json.JsonSerializer.Serialize(dailyTrend.Select(x => x.QtyNG));
                ViewBag.DailyCounts = System.Text.Json.JsonSerializer.Serialize(dailyTrend.Select(x => x.Count));
                ViewBag.DailyMovingAvg = System.Text.Json.JsonSerializer.Serialize(movingAvg);

                // KPI
                ViewBag.TotalOperations = totalOperations;
                ViewBag.TotalRecords = totalRecords;
                ViewBag.TotalQtyNG = totalQtyNG;

                // ─── 5. Phân trang Detail Table ───
                var pagedData = await query
                    .OrderByDescending(x => x.Time_line)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .AsNoTracking()
                    .ToListAsync();

                int totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);

                ViewBag.CurrentPage = page;
                ViewBag.TotalPages = totalPages;
                ViewBag.PageSize = pageSize;
                ViewBag.HasPreviousPage = page > 1;
                ViewBag.HasNextPage = page < totalPages;

                // Filter values
                ViewBag.FromDate = fromDate;
                ViewBag.ToDate = toDate;
                ViewBag.WorkOrder = workOrder;
                ViewBag.Operation = operation;
                ViewBag.DefectCode = defectCode;
                ViewBag.ItemCode = itemCode;

                return View(pagedData);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Lỗi: {ex.Message}";
                // Set defaults
                ViewBag.TotalOperations = 0; ViewBag.TotalRecords = 0; ViewBag.TotalQtyNG = 0;
                ViewBag.CurrentPage = 1; ViewBag.TotalPages = 0; ViewBag.PageSize = pageSize;
                ViewBag.HasPreviousPage = false; ViewBag.HasNextPage = false;
                ViewBag.FromDate = fromDate; ViewBag.ToDate = toDate;
                ViewBag.WorkOrder = workOrder; ViewBag.Operation = operation;
                ViewBag.DefectCode = defectCode; ViewBag.ItemCode = itemCode;
                ViewBag.ChartOperationLabels = "[]"; ViewBag.ChartOperationQtyNG = "[]"; ViewBag.ChartOperationCounts = "[]";
                ViewBag.ParetoLabels = "[]"; ViewBag.ParetoQtyNG = "[]"; ViewBag.ParetoCumulative = "[]";
                ViewBag.PieLabels = "[]"; ViewBag.PieData = "[]";
                ViewBag.DailyDates = "[]"; ViewBag.DailyQtyNG = "[]"; ViewBag.DailyCounts = "[]"; ViewBag.DailyMovingAvg = "[]";
                ViewBag.AllWorkOrders = new List<string>();
                ViewBag.AllOperations = new List<string>();
                ViewBag.AllDefectCodes = new List<string>();
                ViewBag.AllDefectNames = new Dictionary<string, string>();
                ViewBag.AllItemCodes = new List<string>();
                return View(new List<SVN_Defect_Record_History>());
            }
        }


        // Lấy chi tiết 1 record + lịch sử

        [HttpGet]
        public async Task<IActionResult> GetDefectDetail(int id)
        {
            try
            {
                var record = await _context.SVN_Defect_Record_History
                    .FirstOrDefaultAsync(x => x.Id == id);

                if (record == null)
                    return Json(new { success = false, message = "Không tìm thấy bản ghi" });

                // Lịch sử cùng Work Order
                var sameWO = await _context.SVN_Defect_Record_History
                    .Where(x => x.Work_order == record.Work_order && x.Id != id)
                    .OrderByDescending(x => x.Time_line)
                    .Take(10)
                    .Select(x => new
                    {
                        x.Id,
                        x.Defect_Code,
                        x.Defect_Name,
                        x.Qty_NG,
                        x.Operation,
                        x.Employer_name,
                        TimeLine = x.Time_line.HasValue ? x.Time_line.Value.ToString("dd/MM/yyyy HH:mm:ss") : ""
                    })
                    .ToListAsync();

                // Lịch sử cùng Defect Code trong WO đó
                var sameDefectInWO = await _context.SVN_Defect_Record_History
                    .Where(x => x.Work_order == record.Work_order && x.Defect_Code == record.Defect_Code && x.Id != id)
                    .OrderByDescending(x => x.Time_line)
                    .Take(10)
                    .Select(x => new
                    {
                        x.Id,
                        x.Qty_NG,
                        x.Employer_name,
                        x.Note,
                        TimeLine = x.Time_line.HasValue ? x.Time_line.Value.ToString("dd/MM/yyyy HH:mm:ss") : ""
                    })
                    .ToListAsync();

                return Json(new
                {
                    success = true,
                    data = new
                    {
                        record.Id,
                        record.Work_order,
                        record.Item_code,
                        record.Defect_Code,
                        record.Defect_Name,
                        record.Qty_NG,
                        record.Operation,
                        record.Employer_code,
                        record.Employer_name,
                        record.Note,
                        record.Image_error,
                        record.INSDatetime,
                        TimeLine = record.Time_line.HasValue ? record.Time_line.Value.ToString("dd/MM/yyyy HH:mm:ss") : ""
                    },
                    sameWO,
                    sameDefectInWO
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

 
        // Xuất Excel báo cáo
        [HttpGet]
        public async Task<IActionResult> ExportDefectReport(
            string fromDate = "", string toDate = "",
            string workOrder = "", string operation = "",
            string defectCode = "", string itemCode = "")
        {
            // ─── 1. Query có filter (giống Report action) ───
            var query = _context.SVN_Defect_Record_History.AsQueryable();

            if (!string.IsNullOrEmpty(workOrder))  query = query.Where(x => x.Work_order == workOrder);
            if (!string.IsNullOrEmpty(operation))  query = query.Where(x => x.Operation == operation);
            if (!string.IsNullOrEmpty(defectCode)) query = query.Where(x => x.Defect_Code == defectCode);
            if (!string.IsNullOrEmpty(itemCode))   query = query.Where(x => x.Item_code == itemCode);
            if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParse(fromDate, out var f))
                query = query.Where(x => x.INSDatetime.CompareTo(f.ToString("yyyyMMdd")) >= 0);
            if (!string.IsNullOrEmpty(toDate) && DateTime.TryParse(toDate, out var t))
                query = query.Where(x => x.INSDatetime.CompareTo(t.ToString("yyyyMMdd")) <= 0);

            var data = await query.OrderByDescending(x => x.Time_line).AsNoTracking().ToListAsync();

            // ─── 2. Tính KPI ───
            int    kpiOperations = data.Select(x => x.Operation).Distinct().Count();
            int    kpiRecords    = data.Count;
            int    kpiQtyNG      = data.Sum(x => x.Qty_NG ?? 0);

            // ─── 3. Defect theo Operation ───
            var byOperation = data
                .GroupBy(x => x.Operation ?? "N/A")
                .Select(g => new { Operation = g.Key, QtyNG = g.Sum(x => x.Qty_NG ?? 0), Count = g.Count() })
                .OrderByDescending(x => x.QtyNG).ToList();

            // ─── 4. Pareto theo Defect Name ───
            var byDefect = data
                .GroupBy(x => new { x.Defect_Code, x.Defect_Name })
                .Select(g => new { Code = g.Key.Defect_Code ?? "", Name = g.Key.Defect_Name ?? "N/A", QtyNG = g.Sum(x => x.Qty_NG ?? 0) })
                .OrderByDescending(x => x.QtyNG).ToList();

            int    totalNG  = byDefect.Sum(x => x.QtyNG);
            double cumPct   = 0;
            var    paretoList = byDefect.Select(x =>
            {
                cumPct += totalNG > 0 ? (double)x.QtyNG / totalNG * 100 : 0;
                return new { x.Code, x.Name, x.QtyNG, CumPct = Math.Round(cumPct, 1) };
            }).ToList();

            // ─── 5. Daily Trend + Moving Avg 7 ngày ───
            var dailyTrend = data
                .Where(x => x.Time_line.HasValue)
                .GroupBy(x => x.Time_line!.Value.Date)
                .Select(g => new { Date = g.Key, QtyNG = g.Sum(x => x.Qty_NG ?? 0), Count = g.Count() })
                .OrderBy(x => x.Date).ToList();

            var qtyList   = dailyTrend.Select(x => (double)x.QtyNG).ToList();
            var movingAvg = qtyList.Select((_, i) =>
                i < 6 ? (double?)null : Math.Round(qtyList.Skip(i - 6).Take(7).Average(), 1)).ToList();

            // ─── 6. Build Excel ───
            using var wb = new XLWorkbook();

            // Helper styles
            var BLUE   = XLColor.FromHtml("#1e3c72");
            var LBLUE  = XLColor.FromHtml("#2a5298");
            var HEADER = XLColor.FromHtml("#2a5298");
            var ACCENT = XLColor.FromHtml("#ebf0fb");
            var RED    = XLColor.FromHtml("#e53e3e");
            var GOLD   = XLColor.FromHtml("#d69e2e");
            var WHITE  = XLColor.White;
            var GRAY   = XLColor.FromHtml("#718096");

            void SetBorder(IXLRange range)
            {
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder  = XLBorderStyleValues.Thin;
                range.Style.Border.OutsideBorderColor = XLColor.FromHtml("#cbd5e0");
                range.Style.Border.InsideBorderColor  = XLColor.FromHtml("#cbd5e0");
            }

            /* ═══════════════════════════════════════════════
               SHEET 1 – TỔNG QUAN
            ═══════════════════════════════════════════════ */
            var ws1 = wb.Worksheets.Add("Tổng Quan");
            ws1.Style.Font.FontName = "Arial";
            ws1.Style.Font.FontSize = 11;
            ws1.Column(1).Width = 28;
            ws1.Column(2).Width = 22;
            ws1.Column(3).Width = 22;
            ws1.Column(4).Width = 22;
            ws1.Column(5).Width = 22;
            ws1.Column(6).Width = 20;

            int r = 1;

            // ── Tiêu đề chính ──
            ws1.Cell(r, 1).Value = "BÁO CÁO DEFECT";
            var titleRange = ws1.Range(r, 1, r, 6);
            titleRange.Merge();
            ws1.Cell(r, 1).Style.Font.Bold      = true;
            ws1.Cell(r, 1).Style.Font.FontSize  = 16;
            ws1.Cell(r, 1).Style.Font.FontColor = WHITE;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor    = BLUE;
            ws1.Cell(r, 1).Style.Alignment.Horizontal    = XLAlignmentHorizontalValues.Center;
            ws1.Cell(r, 1).Style.Alignment.Vertical      = XLAlignmentVerticalValues.Center;
            ws1.Row(r).Height = 32;
            r++;

            // ── Thông tin lọc ──
            ws1.Cell(r, 1).Value = $"Ngày xuất: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
            ws1.Range(r, 1, r, 6).Merge();
            ws1.Cell(r, 1).Style.Font.Italic    = true;
            ws1.Cell(r, 1).Style.Font.FontColor = GRAY;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f4f8");
            ws1.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            r++;

            // Bộ lọc – dạng lưới 2 cột
            var filters = new[]
            {
                ("Từ ngày",    string.IsNullOrEmpty(fromDate) ? "-- Tất cả --" : DateTime.Parse(fromDate).ToString("dd/MM/yyyy")),
                ("Đến ngày",   string.IsNullOrEmpty(toDate)   ? "-- Tất cả --" : DateTime.Parse(toDate).ToString("dd/MM/yyyy")),
                ("Work Order", string.IsNullOrEmpty(workOrder)  ? "-- Tất cả --" : workOrder),
                ("Operation",  string.IsNullOrEmpty(operation)  ? "-- Tất cả --" : operation),
                ("Defect Code",string.IsNullOrEmpty(defectCode) ? "-- Tất cả --" : defectCode),
                ("Item Code",  string.IsNullOrEmpty(itemCode)   ? "-- Tất cả --" : itemCode),
            };

            ws1.Cell(r, 1).Value = "BỘ LỌC DỮ LIỆU";
            ws1.Range(r, 1, r, 6).Merge();
            ws1.Cell(r, 1).Style.Font.Bold      = true;
            ws1.Cell(r, 1).Style.Font.FontSize  = 12;
            ws1.Cell(r, 1).Style.Font.FontColor = WHITE;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor  = LBLUE;
            ws1.Cell(r, 1).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            ws1.Row(r).Height = 20;
            r++;

            for (int fi = 0; fi < filters.Length; fi += 2)
            {
                // Label 1
                ws1.Cell(r, 1).Value = filters[fi].Item1;
                ws1.Cell(r, 1).Style.Font.Bold      = true;
                ws1.Cell(r, 1).Style.Fill.BackgroundColor = ACCENT;
                // Value 1
                ws1.Cell(r, 2).Value = filters[fi].Item2;
                ws1.Range(r, 2, r, 3).Merge();
                // Label 2
                if (fi + 1 < filters.Length)
                {
                    ws1.Cell(r, 4).Value = filters[fi + 1].Item1;
                    ws1.Cell(r, 4).Style.Font.Bold      = true;
                    ws1.Cell(r, 4).Style.Fill.BackgroundColor = ACCENT;
                    ws1.Cell(r, 5).Value = filters[fi + 1].Item2;
                    ws1.Range(r, 5, r, 6).Merge();
                }
                SetBorder(ws1.Range(r, 1, r, 6));
                r++;
            }
            r++;

            // ── KPI Dashboard ──
            ws1.Cell(r, 1).Value = "DASHBOARD TỔNG QUAN";
            ws1.Range(r, 1, r, 6).Merge();
            ws1.Cell(r, 1).Style.Font.Bold      = true;
            ws1.Cell(r, 1).Style.Font.FontSize  = 12;
            ws1.Cell(r, 1).Style.Font.FontColor = WHITE;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor  = LBLUE;
            ws1.Cell(r, 1).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            ws1.Row(r).Height = 20;
            r++;

            var kpiHeaders = new[] { "Tổng số Operation", "Tổng Defect Records", "Tổng số lượng NG" };
            var kpiIntValues = new[] { kpiOperations, kpiRecords, kpiQtyNG };
            var kpiColors  = new[] { XLColor.FromHtml("#2a5298"), RED, GOLD };

            // Header KPI
            for (int k = 0; k < 3; k++)
            {
                int col = k * 2 + 1;
                ws1.Cell(r, col).Value = kpiHeaders[k];
                ws1.Range(r, col, r, col + 1).Merge();
                ws1.Cell(r, col).Style.Font.Bold      = true;
                ws1.Cell(r, col).Style.Font.FontSize  = 10;
                ws1.Cell(r, col).Style.Font.FontColor = WHITE;
                ws1.Cell(r, col).Style.Fill.BackgroundColor  = kpiColors[k];
                ws1.Cell(r, col).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            }
            ws1.Row(r).Height = 20;
            r++;

            // Value KPI
            for (int k = 0; k < 3; k++)
            {
                int col = k * 2 + 1;
                ws1.Cell(r, col).Value = kpiIntValues[k];
                ws1.Range(r, col, r, col + 1).Merge();
                ws1.Cell(r, col).Style.Font.Bold      = true;
                ws1.Cell(r, col).Style.Font.FontSize  = 20;
                ws1.Cell(r, col).Style.Font.FontColor = kpiColors[k];
                ws1.Cell(r, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#f8faff");
                ws1.Cell(r, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, col).Style.Alignment.Vertical   = XLAlignmentVerticalValues.Center;
                SetBorder(ws1.Range(r, col, r, col + 1));
            }
            ws1.Row(r).Height = 36;
            r += 2;

            // ── Chart 1: Defect theo Operation ──
            ws1.Cell(r, 1).Value = "DEFECT THEO OPERATION";
            ws1.Range(r, 1, r, 6).Merge();
            ws1.Cell(r, 1).Style.Font.Bold      = true;
            ws1.Cell(r, 1).Style.Font.FontSize  = 12;
            ws1.Cell(r, 1).Style.Font.FontColor = WHITE;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor  = LBLUE;
            ws1.Cell(r, 1).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            ws1.Row(r).Height = 20;
            r++;

            var opHeaders = new[] { "STT", "Operation", "Tổng Qty NG", "Số lần phát sinh", "% trên tổng" };
            int[] opColWidths = { 6, 0, 16, 18, 14 };
            for (int i = 0; i < opHeaders.Length; i++)
            {
                var cell = ws1.Cell(r, i + 1);
                cell.Value = opHeaders[i];
                cell.Style.Font.Bold      = true;
                cell.Style.Font.FontColor = WHITE;
                cell.Style.Fill.BackgroundColor = HEADER;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            ws1.Range(r, 5, r, 6).Merge(); // Merge % trên tổng
            SetBorder(ws1.Range(r, 1, r, 5));
            r++;

            int opStartRow = r;
            for (int i = 0; i < byOperation.Count; i++)
            {
                var op = byOperation[i];
                ws1.Cell(r, 1).Value = i + 1;
                ws1.Cell(r, 2).Value = op.Operation;
                ws1.Cell(r, 3).Value = op.QtyNG;
                ws1.Cell(r, 4).Value = op.Count;
                ws1.Cell(r, 5).Value = kpiQtyNG > 0 ? Math.Round((double)op.QtyNG / kpiQtyNG * 100, 1) : 0;
                ws1.Cell(r, 5).Style.NumberFormat.Format = "0.0\"%\"";
                ws1.Range(r, 5, r, 6).Merge();
                ws1.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 3).Style.Font.FontColor = RED;
                ws1.Cell(r, 3).Style.Font.Bold = true;
                if (i % 2 == 1) ws1.Range(r, 1, r, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#f8faff");
                SetBorder(ws1.Range(r, 1, r, 6));
                r++;
            }
            r++;

            // ── Chart 2: Pareto Defect Name ──
            ws1.Cell(r, 1).Value = "PARETO DEFECT THEO TÊN LỖI";
            ws1.Range(r, 1, r, 6).Merge();
            ws1.Cell(r, 1).Style.Font.Bold      = true;
            ws1.Cell(r, 1).Style.Font.FontSize  = 12;
            ws1.Cell(r, 1).Style.Font.FontColor = WHITE;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor  = LBLUE;
            ws1.Cell(r, 1).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            ws1.Row(r).Height = 20;
            r++;

            var paretoHeaders = new[] { "STT", "Defect Code", "Defect Name", "Qty NG", "% trên tổng", "% Tích lũy" };
            for (int i = 0; i < paretoHeaders.Length; i++)
            {
                var cell = ws1.Cell(r, i + 1);
                cell.Value = paretoHeaders[i];
                cell.Style.Font.Bold      = true;
                cell.Style.Font.FontColor = WHITE;
                cell.Style.Fill.BackgroundColor = HEADER;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            SetBorder(ws1.Range(r, 1, r, 6));
            r++;

            for (int i = 0; i < paretoList.Count; i++)
            {
                var p = paretoList[i];
                ws1.Cell(r, 1).Value = i + 1;
                ws1.Cell(r, 2).Value = p.Code;
                ws1.Cell(r, 3).Value = p.Name;
                ws1.Cell(r, 4).Value = p.QtyNG;
                ws1.Cell(r, 5).Value = totalNG > 0 ? Math.Round((double)p.QtyNG / totalNG * 100, 1) : 0;
                ws1.Cell(r, 6).Value = p.CumPct;
                ws1.Cell(r, 5).Style.NumberFormat.Format = "0.0\"%\"";
                ws1.Cell(r, 6).Style.NumberFormat.Format = "0.0\"%\"";
                ws1.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 4).Style.Font.FontColor = RED;
                ws1.Cell(r, 4).Style.Font.Bold      = true;
                // Highlight nếu tích lũy >= 80%
                if (p.CumPct <= 80) ws1.Cell(r, 6).Style.Font.FontColor = XLColor.FromHtml("#059669");
                if (i % 2 == 1) ws1.Range(r, 1, r, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#f8faff");
                SetBorder(ws1.Range(r, 1, r, 6));
                r++;
            }
            r++;

            // ── Chart 3: Daily Trend + Moving Avg ──
            ws1.Cell(r, 1).Value = "DEFECT TREND THEO NGÀY (kèm Moving Avg 7 ngày)";
            ws1.Range(r, 1, r, 6).Merge();
            ws1.Cell(r, 1).Style.Font.Bold      = true;
            ws1.Cell(r, 1).Style.Font.FontSize  = 12;
            ws1.Cell(r, 1).Style.Font.FontColor = WHITE;
            ws1.Cell(r, 1).Style.Fill.BackgroundColor  = LBLUE;
            ws1.Cell(r, 1).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            ws1.Row(r).Height = 20;
            r++;

            var trendHeaders = new[] { "STT", "Ngày", "Qty NG", "Số lần phát sinh", "Moving Avg 7 ngày", "" };
            for (int i = 0; i < 5; i++)
            {
                var cell = ws1.Cell(r, i + 1);
                cell.Value = trendHeaders[i];
                cell.Style.Font.Bold      = true;
                cell.Style.Font.FontColor = WHITE;
                cell.Style.Fill.BackgroundColor = HEADER;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            ws1.Range(r, 5, r, 6).Merge();
            SetBorder(ws1.Range(r, 1, r, 6));
            r++;

            for (int i = 0; i < dailyTrend.Count; i++)
            {
                var d = dailyTrend[i];
                ws1.Cell(r, 1).Value = i + 1;
                ws1.Cell(r, 2).Value = d.Date.ToString("dd/MM/yyyy");
                ws1.Cell(r, 3).Value = d.QtyNG;
                ws1.Cell(r, 4).Value = d.Count;
                if (movingAvg[i].HasValue)
                {
                    ws1.Cell(r, 5).Value = movingAvg[i]!.Value;
                    ws1.Range(r, 5, r, 6).Merge();
                    ws1.Cell(r, 5).Style.Font.FontColor = GOLD;
                }
                else
                {
                    ws1.Range(r, 5, r, 6).Merge();
                    ws1.Cell(r, 5).Value = "—";
                    ws1.Cell(r, 5).Style.Font.FontColor = GRAY;
                }
                ws1.Cell(r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws1.Cell(r, 3).Style.Font.FontColor = RED;
                ws1.Cell(r, 3).Style.Font.Bold = true;
                if (i % 2 == 1) ws1.Range(r, 1, r, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#f8faff");
                SetBorder(ws1.Range(r, 1, r, 6));
                r++;
            }

            /* ═══════════════════════════════════════════════
               SHEET 2 – CHI TIẾT
            ═══════════════════════════════════════════════ */
            var ws2 = wb.Worksheets.Add("Chi Tiết");
            ws2.Style.Font.FontName = "Arial";
            ws2.Style.Font.FontSize = 11;

            int r2 = 1;

            // Tiêu đề sheet 2
            ws2.Cell(r2, 1).Value = "CHI TIẾT DEFECT RECORDS";
            ws2.Range(r2, 1, r2, 10).Merge();
            ws2.Cell(r2, 1).Style.Font.Bold      = true;
            ws2.Cell(r2, 1).Style.Font.FontSize  = 14;
            ws2.Cell(r2, 1).Style.Font.FontColor = WHITE;
            ws2.Cell(r2, 1).Style.Fill.BackgroundColor   = BLUE;
            ws2.Cell(r2, 1).Style.Alignment.Horizontal   = XLAlignmentHorizontalValues.Center;
            ws2.Row(r2).Height = 28;
            r2++;

            // Filter info nhỏ
            string filterSummary = $"Lọc: Từ {(string.IsNullOrEmpty(fromDate) ? "đầu" : DateTime.Parse(fromDate).ToString("dd/MM/yyyy"))} " +
                                   $"đến {(string.IsNullOrEmpty(toDate) ? "nay" : DateTime.Parse(toDate).ToString("dd/MM/yyyy"))}  |  " +
                                   $"Tổng: {kpiRecords} records  |  Qty NG: {kpiQtyNG}";
            ws2.Cell(r2, 1).Value = filterSummary;
            ws2.Range(r2, 1, r2, 10).Merge();
            ws2.Cell(r2, 1).Style.Font.Italic    = true;
            ws2.Cell(r2, 1).Style.Font.FontColor = GRAY;
            ws2.Cell(r2, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f4f8");
            ws2.Cell(r2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            r2 += 2;

            // Header bảng
            string[] detailHeaders = { "STT", "Work Order", "Item Code", "Defect Code", "Defect Name", "Qty NG", "Operation", "Employer", "Note", "Time Line" };
            for (int i = 0; i < detailHeaders.Length; i++)
            {
                var cell = ws2.Cell(r2, i + 1);
                cell.Value = detailHeaders[i];
                cell.Style.Font.Bold      = true;
                cell.Style.Font.FontColor = WHITE;
                cell.Style.Fill.BackgroundColor = HEADER;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical   = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.WrapText   = true;
            }
            ws2.Row(r2).Height = 22;
            SetBorder(ws2.Range(r2, 1, r2, 10));
            r2++;

            int stt = 1;
            foreach (var item in data)
            {
                ws2.Cell(r2, 1).Value  = stt++;
                ws2.Cell(r2, 2).Value  = item.Work_order;
                ws2.Cell(r2, 3).Value  = item.Item_code;
                ws2.Cell(r2, 4).Value  = item.Defect_Code;
                ws2.Cell(r2, 5).Value  = item.Defect_Name;
                ws2.Cell(r2, 6).Value  = item.Qty_NG;
                ws2.Cell(r2, 7).Value  = item.Operation;
                ws2.Cell(r2, 8).Value  = $"{item.Employer_code} – {item.Employer_name}".Trim(' ', '–');
                ws2.Cell(r2, 9).Value  = item.Note;
                ws2.Cell(r2, 10).Value = item.Time_line?.ToString("dd/MM/yyyy HH:mm:ss");

                ws2.Cell(r2, 1).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
                ws2.Cell(r2, 4).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
                ws2.Cell(r2, 6).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
                ws2.Cell(r2, 6).Style.Font.FontColor         = RED;
                ws2.Cell(r2, 6).Style.Font.Bold              = true;
                ws2.Cell(r2, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                if (stt % 2 == 0) ws2.Range(r2, 1, r2, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#f8faff");
                SetBorder(ws2.Range(r2, 1, r2, 10));
                r2++;
            }

            // Auto fit
            ws2.Column(1).Width  = 6;
            ws2.Column(2).Width  = 16;
            ws2.Column(3).Width  = 20;
            ws2.Column(4).Width  = 13;
            ws2.Column(5).Width  = 26;
            ws2.Column(6).Width  = 10;
            ws2.Column(7).Width  = 28;
            ws2.Column(8).Width  = 22;
            ws2.Column(9).Width  = 22;
            ws2.Column(10).Width = 20;

            // Freeze header
            ws2.SheetView.FreezeRows(r2 - data.Count - 1);

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"BaoCaoDefect_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }


    }
}


using System.Text;
using ClosedXML.Excel;
using System.IO;
using System.Threading.Tasks;
using System.Drawing;
using DefectManagement.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

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
                query = query.Where(x => x.INSDatetime.CompareTo(fromDate.ToString("yyyyMMdd")) >= 0);
            if (!string.IsNullOrEmpty(toInsDateTime) && DateTime.TryParse(toInsDateTime, out var toDate))
                query = query.Where(x => x.INSDatetime.CompareTo(toDate.ToString("yyyyMMdd")) <= 0);

            var data = await query.OrderBy(x => x.Time_line).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("DefectHistory");
                var currentRow = 1;

                ws.Style.Font.FontName = "Times New Roman";
                ws.Style.Font.FontSize = 11;

                // ── Header ──
                string[] headers = { "ID", "Work Order", "Item Code", "Defect Code", "Defect Name",
                              "Qty NG", "INS DateTime", "Operation", "Employer Code",
                              "Employer Name", "Note", "Image Error", "Time Line" };

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
                ws.Row(currentRow).Height = 20;

                // ── Data rows ──
                const int IMAGE_COL = 12;
                const double ROW_HEIGHT = 60;
                const int IMG_WIDTH = 80;
                const int IMG_HEIGHT = 60;

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
                    ws.Cell(currentRow, 13).Value = item.Time_line?.ToString("yyyy-MM-dd HH:mm:ss");

                    // ── Xử lý ảnh ──
                    bool imageEmbedded = false;

                    if (!string.IsNullOrWhiteSpace(item.Image_error))
                    {
                        var relativePath = item.Image_error.TrimStart('/').Replace('/', Path.DirectorySeparatorChar);
                        var physicalPath = Path.Combine(_webHostEnvironment.WebRootPath, relativePath);

                        if (System.IO.File.Exists(physicalPath))
                        {
                            try
                            {
                                ws.Row(currentRow).Height = ROW_HEIGHT;

                                using var imgStream = new FileStream(physicalPath, FileMode.Open, FileAccess.Read);

                                ws.AddPicture(imgStream)
                                  .MoveTo(ws.Cell(currentRow, IMAGE_COL))
                                  .WithSize(IMG_WIDTH, IMG_HEIGHT);

                                imageEmbedded = true;
                            }
                            catch
                            {
                                // Nếu lỗi load ảnh → fallback hiện link
                            }
                        }
                    }

                    if (!imageEmbedded)
                    {
                        ws.Cell(currentRow, IMAGE_COL).Value = item.Image_error ?? "";
                    }
                }

                // ── Căn chỉnh ──
                ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Column(7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Column(13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Column(IMAGE_COL).Width = 14;
                ws.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "DefectHistory.xlsx");
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
        public async Task<IActionResult> Create(SVN_Defect_Record_History model, IFormFile imageFile, bool IsNightShift = false)
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

                // ===== Xử lý thời gian ca đêm =====
                string insDatetime;
                DateTime timeLine;

                if (IsNightShift)
                {
                    // INSDatetime = ngày hôm qua (yyyyMMdd)
                    var yesterday = DateTime.Now.AddDays(-1);
                    insDatetime = yesterday.ToString("yyyyMMdd");

                    // Time_line = random trong khoảng 20:00 - 23:59 ngày hôm qua
                    var rng = new Random();
                    int randomMinutes = rng.Next(0, 4 * 60); // 0 đến 239 phút (4 tiếng: 20h-23h59)
                    timeLine = yesterday.Date.AddHours(20).AddMinutes(randomMinutes);
                }
                else
                {
                    insDatetime = DateTime.Now.ToString("yyyyMMdd");
                    timeLine = DateTime.Now;
                }

                // Gọi stored procedure với parameters
                await _context.Database.ExecuteSqlRawAsync(
                    "EXEC [dbo].[SVN_InsertDefectReport] {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}",
                    model.Work_order ?? "",
                    model.Item_code ?? "",
                    model.Defect_Code ?? "",
                    model.Defect_Name ?? "",
                    model.Qty_NG,
                    insDatetime,
                    model.Operation ?? "",
                    model.Employer_code ?? "",
                    model.Employer_name ?? "",
                    model.Note ?? "",
                    imagePath ?? "",
                    timeLine);

                Console.WriteLine($"Stored procedure executed. IsNightShift={IsNightShift}, INSDatetime={insDatetime}, TimeLine={timeLine}");

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

    /* ──────────────────────────────────────────────────────────────
       ExportDefectExcel  –  EPPlus multi-tab workbook (5 sheets)
    ────────────────────────────────────────────────────────────── */
    [AcceptVerbs("GET", "POST")]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> ExportDefectExcel([FromBody] ExportDefectDto? dto)
    {
        try
        {
        ExcelPackage.License.SetNonCommercialOrganization("DefectManagement");

        // Supports both POST (with JSON body + chart images) and GET (query-string fallback)
        var fromDate   = dto?.DateFrom   ?? Request.Query["fromDate"].FirstOrDefault()   ?? "";
        var toDate     = dto?.DateTo     ?? Request.Query["toDate"].FirstOrDefault()     ?? "";
        var workOrder  = dto?.WorkOrder  ?? Request.Query["workOrder"].FirstOrDefault()  ?? "";
        var operation  = dto?.Operation  ?? Request.Query["operation"].FirstOrDefault()  ?? "";
        var defectCode = dto?.DefectCode ?? Request.Query["defectCode"].FirstOrDefault() ?? "";
        var itemCode   = dto?.ItemCode   ?? Request.Query["itemCode"].FirstOrDefault()   ?? "";

        // ── 1. Query (identical filter logic to Report action) ──
        var query = _context.SVN_Defect_Record_History.AsQueryable();
        if (!string.IsNullOrEmpty(workOrder))  query = query.Where(x => x.Work_order  == workOrder);
        if (!string.IsNullOrEmpty(operation))  query = query.Where(x => x.Operation   == operation);
        if (!string.IsNullOrEmpty(defectCode)) query = query.Where(x => x.Defect_Code == defectCode);
        if (!string.IsNullOrEmpty(itemCode))   query = query.Where(x => x.Item_code   == itemCode);
        if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParse(fromDate, out var fd))
            query = query.Where(x => x.INSDatetime.CompareTo(fd.ToString("yyyyMMdd")) >= 0);
        if (!string.IsNullOrEmpty(toDate) && DateTime.TryParse(toDate, out var td))
            query = query.Where(x => x.INSDatetime.CompareTo(td.ToString("yyyyMMdd")) <= 0);

        var data = await query.OrderByDescending(x => x.Time_line).AsNoTracking().ToListAsync();

        // ── 2. Aggregate data for all tabs ──
        int kpiOps     = data.Select(x => x.Operation).Distinct().Count();
        int kpiRecords = data.Count;
        int kpiQtyNG   = data.Sum(x => x.Qty_NG ?? 0);

        var byOperation = data
            .GroupBy(x => x.Operation ?? "N/A")
            .Select(g => new { Operation = g.Key, QtyNG = g.Sum(x => x.Qty_NG ?? 0), Count = g.Count() })
            .OrderByDescending(x => x.QtyNG).ToList();

        var byDefect = data
            .GroupBy(x => new { x.Defect_Code, x.Defect_Name })
            .Select(g => new { Code = g.Key.Defect_Code ?? "", Name = g.Key.Defect_Name ?? "N/A", QtyNG = g.Sum(x => x.Qty_NG ?? 0) })
            .OrderByDescending(x => x.QtyNG).ToList();
        int totalDefectNG = byDefect.Sum(x => x.QtyNG);
        double cumAcc = 0;
        var paretoList = byDefect.Select(x => {
            cumAcc += totalDefectNG > 0 ? (double)x.QtyNG / totalDefectNG * 100 : 0;
            return new {
                x.Code, x.Name, x.QtyNG,
                PctOfTotal = totalDefectNG > 0 ? Math.Round((double)x.QtyNG / totalDefectNG * 100, 1) : 0.0,
                CumPct = Math.Round(cumAcc, 1)
            };
        }).ToList();

        var top10Ops = byOperation.Take(10).ToList();

        var dailyTrend = data
            .Where(x => x.Time_line.HasValue)
            .GroupBy(x => x.Time_line!.Value.Date)
            .Select(g => new { Date = g.Key, QtyNG = g.Sum(x => x.Qty_NG ?? 0), Count = g.Count() })
            .OrderBy(x => x.Date).ToList();
        var qtyList   = dailyTrend.Select(x => (double)x.QtyNG).ToList();
        var movingAvg = qtyList.Select((_, i) =>
            i < 6 ? (double?)null : Math.Round(qtyList.Skip(i - 6).Take(7).Average(), 1)).ToList();

        // ── 3. Colors ──
        var cBlueDark  = Color.FromArgb(0x1B, 0x4F, 0x8A);
        var cBlueMid   = Color.FromArgb(0x2A, 0x52, 0x98);
        var cRed       = Color.FromArgb(0xE5, 0x3E, 0x3E);
        var cAmber     = Color.FromArgb(0xD6, 0x9E, 0x2E);
        var cWhite     = Color.White;
        var cLightBlue = Color.FromArgb(0xEF, 0xF6, 0xFF);
        var cLightGray = Color.FromArgb(0xF0, 0xF4, 0xF8);
        var cGray      = Color.FromArgb(0x71, 0x80, 0x96);
        var cGreen     = Color.FromArgb(0x05, 0x96, 0x69);

        using var pkg = new ExcelPackage();
        var openStreams = new List<MemoryStream>();

        // ── Shared style helpers ──
        void StyleHeader(ExcelRange range, Color bg, Color fg, int fs = 11)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(bg);
            range.Style.Font.Color.SetColor(fg);
            range.Style.Font.Bold = true;
            range.Style.Font.Size = fs;
            range.Style.Font.Name = "Calibri";
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
        }
        void ThinBorder(ExcelRange range)
        {
            var bc = Color.FromArgb(0xCB, 0xD5, 0xE0);
            range.Style.Border.Top.Style    = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style   = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style  = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Color.SetColor(bc);
            range.Style.Border.Bottom.Color.SetColor(bc);
            range.Style.Border.Left.Color.SetColor(bc);
            range.Style.Border.Right.Color.SetColor(bc);
        }
        void SetFill(ExcelRange range, Color bg)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(bg);
        }

        /* ═══════════════════════════════════════════════════
           TAB 1 – Summary
        ═══════════════════════════════════════════════════ */
        var ws1 = pkg.Workbook.Worksheets.Add("Summary");
        ws1.View.ShowGridLines = false;
        ws1.Cells.Style.Font.Name = "Calibri";
        ws1.Cells.Style.Font.Size = 11;

        int row = 1;

        // Main title
        ws1.Cells[row, 1, row, 11].Merge = true;
        ws1.Cells[row, 1].Value = "BÁO CÁO DEFECT";
        StyleHeader(ws1.Cells[row, 1, row, 11], cBlueDark, cWhite, 16);
        ws1.Row(row).Height = 38;
        row++;

        // Export timestamp
        ws1.Cells[row, 1, row, 11].Merge = true;
        ws1.Cells[row, 1].Value = $"Exported at: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
        ws1.Cells[row, 1].Style.Font.Italic = true;
        ws1.Cells[row, 1].Style.Font.Color.SetColor(cGray);
        SetFill(ws1.Cells[row, 1, row, 11], cLightGray);
        ws1.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        row++;

        // Filter info rows
        var filters = new (string Label, string Value)[]
        {
            ("Từ ngày",     string.IsNullOrEmpty(fromDate) ? "Tất cả" : DateTime.Parse(fromDate).ToString("dd/MM/yyyy")),
            ("Đến ngày",    string.IsNullOrEmpty(toDate)   ? "Tất cả" : DateTime.Parse(toDate).ToString("dd/MM/yyyy")),
            ("Work Order",  string.IsNullOrEmpty(workOrder)  ? "Tất cả" : workOrder),
            ("Operation",   string.IsNullOrEmpty(operation)  ? "Tất cả" : operation),
            ("Defect Code", string.IsNullOrEmpty(defectCode) ? "Tất cả" : defectCode),
            ("Item Code",   string.IsNullOrEmpty(itemCode)   ? "Tất cả" : itemCode),
        };
        foreach (var (label, value) in filters)
        {
            ws1.Cells[row, 1].Value = label;
            ws1.Cells[row, 1].Style.Font.Bold = true;
            SetFill(ws1.Cells[row, 1], cLightBlue);
            ws1.Cells[row, 2, row, 4].Merge = true;
            ws1.Cells[row, 2].Value = value;
            ThinBorder(ws1.Cells[row, 1, row, 4]);
            row++;
        }
        row++; // blank row

        // KPI section title
        ws1.Cells[row, 1, row, 11].Merge = true;
        ws1.Cells[row, 1].Value = "DASHBOARD TỔNG QUAN";
        StyleHeader(ws1.Cells[row, 1, row, 11], cBlueMid, cWhite, 12);
        ws1.Row(row).Height = 22;
        row++;

        // KPI header labels
        var kpiLabels = new[] { "Tổng số Operation", "Tổng Defect Records", "Tổng số lượng NG" };
        var kpiVals   = new[] { kpiOps, kpiRecords, kpiQtyNG };
        var kpiColors = new[] { cBlueMid, cRed, cAmber };
        for (int k = 0; k < 3; k++)
        {
            int c = k * 4 + 1;
            ws1.Cells[row, c, row, c + 3].Merge = true;
            ws1.Cells[row, c].Value = kpiLabels[k];
            StyleHeader(ws1.Cells[row, c, row, c + 3], kpiColors[k], cWhite, 10);
        }
        ws1.Row(row).Height = 22;
        row++;

        // KPI values
        for (int k = 0; k < 3; k++)
        {
            int c = k * 4 + 1;
            ws1.Cells[row, c, row, c + 3].Merge = true;
            ws1.Cells[row, c].Value = kpiVals[k];
            SetFill(ws1.Cells[row, c, row, c + 3], Color.FromArgb(0xF8, 0xFA, 0xFF));
            ws1.Cells[row, c].Style.Font.Bold = true;
            ws1.Cells[row, c].Style.Font.Size = 22;
            ws1.Cells[row, c].Style.Font.Color.SetColor(kpiColors[k]);
            ws1.Cells[row, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws1.Cells[row, c].Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
            ws1.Cells[row, c].Style.Numberformat.Format = "#,##0";
            ThinBorder(ws1.Cells[row, c, row, c + 3]);
        }
        ws1.Row(row).Height = 42;
        row += 2;

        // Detail records section title
        ws1.Cells[row, 1, row, 11].Merge = true;
        ws1.Cells[row, 1].Value = "CHI TIẾT DEFECT RECORDS";
        StyleHeader(ws1.Cells[row, 1, row, 11], cBlueMid, cWhite, 12);
        ws1.Row(row).Height = 22;
        row++;

        // Detail table header
        string[] detailHeaders = { "#", "Work Order", "Item Code", "Defect Code", "Defect Name", "Qty NG", "Operation", "Employer", "Note", "Image", "Time Line" };
        for (int c = 0; c < detailHeaders.Length; c++)
        {
            ws1.Cells[row, c + 1].Value = detailHeaders[c];
            StyleHeader(ws1.Cells[row, c + 1], cBlueMid, cWhite);
            ThinBorder(ws1.Cells[row, c + 1]);
        }
        ws1.Row(row).Height = 22;
        ws1.View.FreezePanes(row + 1, 1);
        ws1.Cells[row, 1, row, detailHeaders.Length].AutoFilter = true;
        row++;

        int stt = 1;
        foreach (var item in data)
        {
            bool even  = stt % 2 == 0;
            var  rowBg = even ? cLightBlue : cWhite;

            ws1.Cells[row, 1].Value  = stt;
            ws1.Cells[row, 2].Value  = item.Work_order;
            ws1.Cells[row, 3].Value  = item.Item_code;
            ws1.Cells[row, 4].Value  = item.Defect_Code;
            ws1.Cells[row, 5].Value  = item.Defect_Name;
            ws1.Cells[row, 6].Value  = item.Qty_NG;
            ws1.Cells[row, 7].Value  = item.Operation;
            ws1.Cells[row, 8].Value  = $"{item.Employer_code} {item.Employer_name}".Trim();
            ws1.Cells[row, 9].Value  = item.Note;
            ws1.Cells[row, 11].Value = item.Time_line?.ToString("dd/MM/yyyy HH:mm:ss");

            ws1.Cells[row, 6].Style.Numberformat.Format = "#,##0";
            ws1.Cells[row, 6].Style.Font.Bold = true;
            if ((item.Qty_NG ?? 0) > 10)
                ws1.Cells[row, 6].Style.Font.Color.SetColor(cRed);

            // Operation pill style
            SetFill(ws1.Cells[row, 7], cBlueMid);
            ws1.Cells[row, 7].Style.Font.Color.SetColor(cWhite);
            ws1.Cells[row, 7].Style.Font.Bold = true;
            ws1.Cells[row, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Row alternate background (skip op cell which has its own bg)
            foreach (int c in new[] { 1, 2, 3, 4, 5, 6, 8, 9, 11 })
                SetFill(ws1.Cells[row, c], rowBg);

            // Embed image
            bool imgOk = false;
            if (!string.IsNullOrWhiteSpace(item.Image_error))
            {
                var rel  = item.Image_error.TrimStart('/').Replace('/', Path.DirectorySeparatorChar);
                var phys = Path.Combine(_webHostEnvironment.WebRootPath, rel);
                if (System.IO.File.Exists(phys))
                {
                    try
                    {
                        var pic = ws1.Drawings.AddPicture($"img_{row}", new FileInfo(phys));
                        pic.SetPosition(row - 1, 2, 9, 2); // row/col 0-based
                        pic.SetSize(46, 46);
                        ws1.Row(row).Height = 45;
                        imgOk = true;
                    }
                    catch { }
                }
            }
            if (!imgOk)
            {
                ws1.Cells[row, 10].Value = "—";
                ws1.Row(row).Height = 18;
            }

            ws1.Cells[row, 1].Style.HorizontalAlignment  = ExcelHorizontalAlignment.Center;
            ws1.Cells[row, 4].Style.HorizontalAlignment  = ExcelHorizontalAlignment.Center;
            ws1.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ThinBorder(ws1.Cells[row, 1, row, 11]);

            row++;
            stt++;
        }

        ws1.Column(1).Width  = 6;
        ws1.Column(2).Width  = 16;
        ws1.Column(3).Width  = 18;
        ws1.Column(4).Width  = 13;
        ws1.Column(5).Width  = 26;
        ws1.Column(6).Width  = 10;
        ws1.Column(7).Width  = 22;
        ws1.Column(8).Width  = 22;
        ws1.Column(9).Width  = 22;
        ws1.Column(10).Width = 10;
        ws1.Column(11).Width = 20;
        ws1.PrinterSettings.PrintArea = ws1.Cells[1, 1, row - 1, 11];

        /* ═══════════════════════════════════════════════════
           TAB 2 – Defect by Operation
        ═══════════════════════════════════════════════════ */
        var ws2 = pkg.Workbook.Worksheets.Add("Defect by Operation");
        ws2.View.ShowGridLines = false;
        ws2.Cells.Style.Font.Name = "Calibri";
        ws2.Cells.Style.Font.Size = 11;

        int r2 = 1;
        ws2.Cells[r2, 1, r2, 4].Merge = true;
        ws2.Cells[r2, 1].Value = "Biểu đồ Defect theo Operation";
        StyleHeader(ws2.Cells[r2, 1, r2, 4], cBlueDark, cWhite, 14);
        ws2.Row(r2).Height = 30;
        r2 += 2;

        string[] opHdrs = { "Operation", "Qty NG", "Số lần phát sinh", "% Tổng" };
        for (int c = 0; c < opHdrs.Length; c++)
        {
            ws2.Cells[r2, c + 1].Value = opHdrs[c];
            StyleHeader(ws2.Cells[r2, c + 1], cBlueMid, cWhite);
            ThinBorder(ws2.Cells[r2, c + 1]);
        }
        ws2.Row(r2).Height = 22;
        ws2.View.FreezePanes(r2 + 1, 1);
        int opDataStart = r2 + 1;
        r2++;

        int opIdx = 0;
        foreach (var op in byOperation)
        {
            bool even = opIdx % 2 == 0;
            SetFill(ws2.Cells[r2, 1, r2, 4], even ? cWhite : cLightBlue);
            ws2.Cells[r2, 1].Value = op.Operation;
            ws2.Cells[r2, 2].Value = op.QtyNG;
            ws2.Cells[r2, 3].Value = op.Count;
            ws2.Cells[r2, 4].Value = kpiQtyNG > 0 ? Math.Round((double)op.QtyNG / kpiQtyNG * 100, 1) : 0.0;
            ws2.Cells[r2, 2].Style.Numberformat.Format = "#,##0";
            ws2.Cells[r2, 4].Style.Numberformat.Format = "0.0";
            ws2.Cells[r2, 2].Style.Font.Bold = true;
            ws2.Cells[r2, 2].Style.Font.Color.SetColor(cRed);
            ws2.Cells[r2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws2.Cells[r2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws2.Cells[r2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ThinBorder(ws2.Cells[r2, 1, r2, 4]);
            r2++;
            opIdx++;
        }
        int opDataEnd = r2 - 1;

        if (byOperation.Count > 0)
        {
            var barChart = ws2.Drawings.AddChart("OperationChart", eChartType.BarClustered) as ExcelBarChart;
            if (barChart != null)
            {
                barChart.SetPosition(r2 + 1, 0, 0, 0);
                barChart.SetSize(620, 350);
                var s = barChart.Series.Add(
                    ws2.Cells[opDataStart, 2, opDataEnd, 2],
                    ws2.Cells[opDataStart, 1, opDataEnd, 1]);
                s.Header = "Qty NG";
                barChart.Title.Text = "Defect theo Operation";
                barChart.Legend.Position = eLegendPosition.Bottom;
                barChart.YAxis.Title.Text = "Qty NG";
                barChart.XAxis.Title.Text = "Operation";
            }
        }

        ws2.Column(1).Width = 28;
        ws2.Column(2).Width = 12;
        ws2.Column(3).Width = 18;
        ws2.Column(4).Width = 12;
        ws2.PrinterSettings.PrintArea = ws2.Cells[1, 1, opDataEnd, 4];
        EmbedChartImage(ws2, dto?.ChartImages?.ByOperation, opDataEnd + 2, 0, 800, 400, openStreams);

        /* ═══════════════════════════════════════════════════
           TAB 3 – Pareto Defect
        ═══════════════════════════════════════════════════ */
        var ws3 = pkg.Workbook.Worksheets.Add("Pareto Defect");
        ws3.View.ShowGridLines = false;
        ws3.Cells.Style.Font.Name = "Calibri";
        ws3.Cells.Style.Font.Size = 11;

        int r3 = 1;
        ws3.Cells[r3, 1, r3, 5].Merge = true;
        ws3.Cells[r3, 1].Value = "Biểu đồ Pareto Defect theo tên lỗi";
        StyleHeader(ws3.Cells[r3, 1, r3, 5], cBlueDark, cWhite, 14);
        ws3.Row(r3).Height = 30;
        r3 += 2;

        string[] paretoHdrs = { "STT", "Defect Name", "Qty NG", "% Tổng", "% Tích Lũy" };
        for (int c = 0; c < paretoHdrs.Length; c++)
        {
            ws3.Cells[r3, c + 1].Value = paretoHdrs[c];
            StyleHeader(ws3.Cells[r3, c + 1], cBlueMid, cWhite);
            ThinBorder(ws3.Cells[r3, c + 1]);
        }
        ws3.Row(r3).Height = 22;
        ws3.View.FreezePanes(r3 + 1, 1);
        int paretoDataStart = r3 + 1;
        r3++;

        int pStt = 1;
        foreach (var p in paretoList)
        {
            bool even = pStt % 2 == 0;
            SetFill(ws3.Cells[r3, 1, r3, 5], even ? cLightBlue : cWhite);
            ws3.Cells[r3, 1].Value = pStt++;
            ws3.Cells[r3, 2].Value = $"{p.Code} - {p.Name}";
            ws3.Cells[r3, 3].Value = p.QtyNG;
            ws3.Cells[r3, 4].Value = p.PctOfTotal;
            ws3.Cells[r3, 5].Value = p.CumPct;
            ws3.Cells[r3, 3].Style.Numberformat.Format = "#,##0";
            ws3.Cells[r3, 4].Style.Numberformat.Format = "0.0";
            ws3.Cells[r3, 5].Style.Numberformat.Format = "0.0";
            ws3.Cells[r3, 3].Style.Font.Bold = true;
            ws3.Cells[r3, 3].Style.Font.Color.SetColor(cRed);
            if (p.CumPct <= 80) ws3.Cells[r3, 5].Style.Font.Color.SetColor(cGreen);
            ws3.Cells[r3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws3.Cells[r3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws3.Cells[r3, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws3.Cells[r3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ThinBorder(ws3.Cells[r3, 1, r3, 5]);
            r3++;
        }
        int paretoDataEnd = r3 - 1;

        // Pareto chart: bar (Qty NG) + line (% Tích Lũy) combo
        if (paretoList.Count > 0)
        {
            var paretoChart = ws3.Drawings.AddChart("ParetoChart", eChartType.BarClustered);
            paretoChart.SetPosition(r3 + 1, 0, 0, 0);
            paretoChart.SetSize(720, 360);
            var barSeries = paretoChart.Series.Add(
                ws3.Cells[paretoDataStart, 3, paretoDataEnd, 3],
                ws3.Cells[paretoDataStart, 2, paretoDataEnd, 2]);
            barSeries.Header = "Qty NG";

            try
            {
                var lineType   = paretoChart.PlotArea.ChartTypes.Add(eChartType.Line);
                var lineSeriesP = lineType.Series.Add(
                    ws3.Cells[paretoDataStart, 5, paretoDataEnd, 5],
                    ws3.Cells[paretoDataStart, 2, paretoDataEnd, 2]);
                lineSeriesP.Header = "% Tích Lũy";
                lineType.UseSecondaryAxis = true;
            }
            catch { /* secondary axis optional – base chart still renders */ }

            paretoChart.Title.Text = "Pareto Defect theo tên lỗi";
            paretoChart.Legend.Position = eLegendPosition.Bottom;
        }

        ws3.Column(1).Width = 6;
        ws3.Column(2).Width = 34;
        ws3.Column(3).Width = 12;
        ws3.Column(4).Width = 10;
        ws3.Column(5).Width = 12;
        ws3.PrinterSettings.PrintArea = ws3.Cells[1, 1, paretoDataEnd, 5];
        EmbedChartImage(ws3, dto?.ChartImages?.Pareto, paretoDataEnd + 2, 0, 900, 400, openStreams);

        /* ═══════════════════════════════════════════════════
           TAB 4 – Top 10 Operations
        ═══════════════════════════════════════════════════ */
        var ws4 = pkg.Workbook.Worksheets.Add("Top 10 Operations");
        ws4.View.ShowGridLines = false;
        ws4.Cells.Style.Font.Name = "Calibri";
        ws4.Cells.Style.Font.Size = 11;

        int r4 = 1;
        ws4.Cells[r4, 1, r4, 4].Merge = true;
        ws4.Cells[r4, 1].Value = "Biểu đồ Top 10 Operation nhiều lỗi nhất";
        StyleHeader(ws4.Cells[r4, 1, r4, 4], cBlueDark, cWhite, 14);
        ws4.Row(r4).Height = 30;
        r4 += 2;

        string[] top10Hdrs = { "Rank", "Operation", "Qty NG", "% Tổng" };
        for (int c = 0; c < top10Hdrs.Length; c++)
        {
            ws4.Cells[r4, c + 1].Value = top10Hdrs[c];
            StyleHeader(ws4.Cells[r4, c + 1], cBlueMid, cWhite);
            ThinBorder(ws4.Cells[r4, c + 1]);
        }
        ws4.Row(r4).Height = 22;
        int top10DataStart = r4 + 1;
        r4++;

        int rankNum = 1;
        foreach (var op in top10Ops)
        {
            bool even = rankNum % 2 == 0;
            SetFill(ws4.Cells[r4, 1, r4, 4], even ? cLightBlue : cWhite);
            ws4.Cells[r4, 1].Value = rankNum++;
            ws4.Cells[r4, 2].Value = op.Operation;
            ws4.Cells[r4, 3].Value = op.QtyNG;
            ws4.Cells[r4, 4].Value = kpiQtyNG > 0 ? Math.Round((double)op.QtyNG / kpiQtyNG * 100, 1) : 0.0;
            ws4.Cells[r4, 3].Style.Numberformat.Format = "#,##0";
            ws4.Cells[r4, 4].Style.Numberformat.Format = "0.0";
            ws4.Cells[r4, 3].Style.Font.Bold = true;
            ws4.Cells[r4, 3].Style.Font.Color.SetColor(cRed);
            ws4.Cells[r4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws4.Cells[r4, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws4.Cells[r4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ThinBorder(ws4.Cells[r4, 1, r4, 4]);
            r4++;
        }
        int top10DataEnd = r4 - 1;

        // Doughnut chart
        if (top10Ops.Count > 0)
        {
            var donutChart = ws4.Drawings.AddChart("Top10Chart", eChartType.Doughnut) as ExcelDoughnutChart;
            if (donutChart != null)
            {
                donutChart.SetPosition(r4 + 1, 0, 0, 0);
                donutChart.SetSize(520, 400);
                var s = donutChart.Series.Add(
                    ws4.Cells[top10DataStart, 3, top10DataEnd, 3],
                    ws4.Cells[top10DataStart, 2, top10DataEnd, 2]);
                s.Header = "Qty NG";
                donutChart.Title.Text = "Top 10 Operations nhiều lỗi nhất";
                donutChart.Legend.Position = eLegendPosition.Right;
            }
        }

        ws4.Column(1).Width = 8;
        ws4.Column(2).Width = 28;
        ws4.Column(3).Width = 12;
        ws4.Column(4).Width = 10;
        ws4.PrinterSettings.PrintArea = ws4.Cells[1, 1, top10DataEnd, 4];
        EmbedChartImage(ws4, dto?.ChartImages?.Top10, top10DataEnd + 2, 0, 600, 450, openStreams);

        /* ═══════════════════════════════════════════════════
           TAB 5 – Defect Trend
        ═══════════════════════════════════════════════════ */
        var ws5 = pkg.Workbook.Worksheets.Add("Defect Trend");
        ws5.View.ShowGridLines = false;
        ws5.Cells.Style.Font.Name = "Calibri";
        ws5.Cells.Style.Font.Size = 11;

        int r5 = 1;
        ws5.Cells[r5, 1, r5, 4].Merge = true;
        ws5.Cells[r5, 1].Value = "Defect Trend theo ngày";
        StyleHeader(ws5.Cells[r5, 1, r5, 4], cBlueDark, cWhite, 14);
        ws5.Row(r5).Height = 30;
        r5 += 2;

        string[] trendHdrs = { "Ngày", "Qty NG", "Số lần phát sinh lỗi", "Moving Avg 7 ngày" };
        for (int c = 0; c < trendHdrs.Length; c++)
        {
            ws5.Cells[r5, c + 1].Value = trendHdrs[c];
            StyleHeader(ws5.Cells[r5, c + 1], cBlueMid, cWhite);
            ThinBorder(ws5.Cells[r5, c + 1]);
        }
        ws5.Row(r5).Height = 22;
        ws5.View.FreezePanes(r5 + 1, 1);
        int trendDataStart = r5 + 1;
        r5++;

        for (int i = 0; i < dailyTrend.Count; i++)
        {
            var d = dailyTrend[i];
            bool even = i % 2 == 0;
            SetFill(ws5.Cells[r5, 1, r5, 4], even ? cWhite : cLightBlue);
            ws5.Cells[r5, 1].Value = d.Date.ToString("dd/MM/yyyy");
            ws5.Cells[r5, 2].Value = d.QtyNG;
            ws5.Cells[r5, 3].Value = d.Count;
            if (movingAvg[i].HasValue) ws5.Cells[r5, 4].Value = movingAvg[i]!.Value;
            else                       ws5.Cells[r5, 4].Value = "—";
            ws5.Cells[r5, 2].Style.Numberformat.Format = "#,##0";
            ws5.Cells[r5, 2].Style.Font.Bold = true;
            ws5.Cells[r5, 2].Style.Font.Color.SetColor(cRed);
            ws5.Cells[r5, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws5.Cells[r5, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws5.Cells[r5, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws5.Cells[r5, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ThinBorder(ws5.Cells[r5, 1, r5, 4]);
            r5++;
        }
        int trendDataEnd = r5 - 1;

        // Multi-series line chart
        if (dailyTrend.Count > 0)
        {
            var lineChart = ws5.Drawings.AddChart("TrendChart", eChartType.Line) as ExcelLineChart;
            if (lineChart != null)
            {
                lineChart.SetPosition(r5 + 1, 0, 0, 0);
                lineChart.SetSize(820, 320);

                var s1 = lineChart.Series.Add(
                    ws5.Cells[trendDataStart, 2, trendDataEnd, 2],
                    ws5.Cells[trendDataStart, 1, trendDataEnd, 1]);
                s1.Header = "Qty NG";

                var s2 = lineChart.Series.Add(
                    ws5.Cells[trendDataStart, 3, trendDataEnd, 3],
                    ws5.Cells[trendDataStart, 1, trendDataEnd, 1]);
                s2.Header = "Số lần phát sinh lỗi";

                // Moving average only from row 7 onward (first 6 are null)
                int maStart = trendDataStart + 6;
                if (maStart <= trendDataEnd)
                {
                    var s3 = lineChart.Series.Add(
                        ws5.Cells[maStart, 4, trendDataEnd, 4],
                        ws5.Cells[maStart, 1, trendDataEnd, 1]);
                    s3.Header = "Moving Avg 7 ngày";
                }

                lineChart.Title.Text = "Defect Trend theo ngày";
                lineChart.Legend.Position = eLegendPosition.Top;
                lineChart.YAxis.Title.Text  = "Qty NG";
                lineChart.XAxis.Title.Text  = "Ngày";
            }
        }

        ws5.Column(1).Width = 14;
        ws5.Column(2).Width = 12;
        ws5.Column(3).Width = 22;
        ws5.Column(4).Width = 18;
        ws5.PrinterSettings.PrintArea = ws5.Cells[1, 1, trendDataEnd, 4];
        EmbedChartImage(ws5, dto?.ChartImages?.Trend, trendDataEnd + 2, 0, 1000, 350, openStreams);

        // ── Return file ──
        var bytes = pkg.GetAsByteArray();
        foreach (var s in openStreams) s.Dispose();
        return File(bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"DefectReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        catch (Exception ex)
        {
            Response.StatusCode = 500;
            return Content($"Lỗi xuất Excel: {ex.Message}", "text/plain; charset=utf-8");
        }
    }

    private static void EmbedChartImage(ExcelWorksheet ws, string? base64,
                                        int startRow, int startCol,
                                        int widthPx, int heightPx,
                                        List<MemoryStream> openStreams)
    {
        if (string.IsNullOrEmpty(base64)) return;
        try
        {
            var comma = base64.IndexOf(',');
            var data  = comma >= 0 ? base64[(comma + 1)..] : base64;
            var bytes = Convert.FromBase64String(data);
            var ms    = new MemoryStream(bytes); // no 'using' — kept alive until GetAsByteArray()
            ms.Position = 0;
            openStreams.Add(ms);
            var picName = $"chartImg_{ws.Name}";
            var pic = ws.Drawings.AddPicture(picName, ms);
            pic.SetPosition(startRow, 0, startCol, 0);
            pic.SetSize(widthPx, heightPx);
        }
        catch { /* skip silently if image data is invalid */ }
    }

    } // DefectController
} // namespace


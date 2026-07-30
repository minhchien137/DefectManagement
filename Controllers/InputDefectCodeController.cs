using System.Security.Claims;
using System.Text.RegularExpressions;
using DefectManagement.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace DefectManagement.Controllers
{
    public class InputDefectCodeController : Controller
    {
        private const string AdminUsername = "admin";
        private const string AdminPassword = "123";

        private readonly ApplicationDbContext _context;

        public InputDefectCodeController(ApplicationDbContext context)
        {
            _context = context;
        }

        [AllowAnonymous]
        [HttpGet]
        public IActionResult Login(string returnUrl = null)
        {
            if (User.Identity != null && User.Identity.IsAuthenticated)
                return RedirectToAction(nameof(Create));

            ViewBag.ReturnUrl = returnUrl;
            return View();
        }

        [AllowAnonymous]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(string username, string password, string returnUrl = null)
        {
            if (username == AdminUsername && password == AdminPassword)
            {
                var claims = new List<Claim> { new Claim(ClaimTypes.Name, username) };
                var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);

                await HttpContext.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    new ClaimsPrincipal(identity),
                    new AuthenticationProperties { IsPersistent = false });

                if (!string.IsNullOrEmpty(returnUrl) && Url.IsLocalUrl(returnUrl))
                    return Redirect(returnUrl);

                return RedirectToAction(nameof(Create));
            }

            ViewBag.ErrorMessage = "Sai tài khoản hoặc mật khẩu!";
            ViewBag.ReturnUrl = returnUrl;
            return View();
        }

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return RedirectToAction(nameof(Login));
        }

        [Authorize]
        [HttpGet]
        public async Task<IActionResult> Create()
        {
            ViewBag.NextCodePreview = "R" + (await GetMaxCodeNumberAsync() + 1);
            return View();
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] List<QualityReasonRowInput> rows)
        {
            if (rows == null || rows.Count == 0)
                return Json(new { success = false, message = "Không có dữ liệu để lưu!" });

            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];

                if (string.IsNullOrWhiteSpace(row.Name) ||
                    string.IsNullOrWhiteSpace(row.Operation) ||
                    row.Priority == null)
                {
                    return Json(new { success = false, message = $"Dòng {i + 1}: vui lòng nhập đầy đủ thông tin bắt buộc!" });
                }

                if (Regex.IsMatch(row.Operation, @"\s"))
                {
                    return Json(new { success = false, message = $"Dòng {i + 1}: Operation không được chứa khoảng trắng!" });
                }
            }

            int nextNumber = await GetMaxCodeNumberAsync() + 1;

            var entities = new List<SVN_quality_reason>();
            foreach (var row in rows)
            {
                entities.Add(new SVN_quality_reason
                {
                    name = row.Name.Trim(),
                    code = "R" + nextNumber,
                    priority = row.Priority!.Value,
                    operation = row.Operation.Trim(),
                    create_uid = null,
                    write_uid = null,
                    create_date = null,
                    write_date = null
                });
                nextNumber++;
            }

            _context.SVN_quality_reason.AddRange(entities);
            await _context.SaveChangesAsync();

            return Json(new { success = true, message = $"Đã lưu thành công {entities.Count} bản ghi!" });
        }

        // Danh sách tên lỗi (name) không trùng lặp, dùng cho dropdown tìm kiếm/thêm mới
        [Authorize]
        [HttpGet]
        public async Task<IActionResult> GetDistinctNames()
        {
            var names = await _context.SVN_quality_reason
                .Where(x => x.name != null && x.name != "")
                .Select(x => x.name)
                .Distinct()
                .OrderBy(x => x)
                .ToListAsync();

            return Json(names);
        }

        [Authorize]
        [HttpGet]
        public async Task<IActionResult> GetList()
        {
            var data = await _context.SVN_quality_reason
                .OrderByDescending(x => x.id)
                .Select(x => new { x.id, x.code, x.name, x.priority, x.operation })
                .ToListAsync();

            return Json(data);
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> Update([FromBody] QualityReasonUpdateInput input)
        {
            if (input == null || input.Id <= 0)
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!" });

            if (string.IsNullOrWhiteSpace(input.Name) ||
                string.IsNullOrWhiteSpace(input.Operation) ||
                input.Priority == null)
            {
                return Json(new { success = false, message = "Vui lòng điền đầy đủ thông tin bắt buộc!" });
            }

            if (Regex.IsMatch(input.Operation, @"\s"))
                return Json(new { success = false, message = "Operation không được chứa khoảng trắng!" });

            var entity = await _context.SVN_quality_reason.FirstOrDefaultAsync(x => x.id == input.Id);
            if (entity == null)
                return Json(new { success = false, message = "Không tìm thấy bản ghi!" });

            entity.name = input.Name.Trim();
            entity.priority = input.Priority.Value;
            entity.operation = input.Operation.Trim();

            await _context.SaveChangesAsync();

            return Json(new { success = true, message = "Cập nhật thành công!" });
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> Delete([FromBody] QualityReasonDeleteInput input)
        {
            if (input == null || input.Id <= 0)
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!" });

            var entity = await _context.SVN_quality_reason.FirstOrDefaultAsync(x => x.id == input.Id);
            if (entity == null)
                return Json(new { success = false, message = "Không tìm thấy bản ghi!" });

            _context.SVN_quality_reason.Remove(entity);
            await _context.SaveChangesAsync();

            return Json(new { success = true, message = "Xóa thành công!" });
        }

        // Lấy số lớn nhất hiện có trong code dạng R<số>, ví dụ "R383" -> 383
        private async Task<int> GetMaxCodeNumberAsync()
        {
            var codes = await _context.SVN_quality_reason
                .Select(x => x.code)
                .Where(c => c != null)
                .ToListAsync();

            int max = 0;
            foreach (var c in codes)
            {
                if (c == null) continue;

                var m = Regex.Match(c, @"^R(\d+)$", RegexOptions.IgnoreCase);
                if (m.Success && int.TryParse(m.Groups[1].Value, out var n) && n > max)
                    max = n;
            }

            return max;
        }
    }

    public class QualityReasonRowInput
    {
        public string Name { get; set; }
        public int? Priority { get; set; }
        public string Operation { get; set; }
    }

    public class QualityReasonUpdateInput
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int? Priority { get; set; }
        public string Operation { get; set; }
    }

    public class QualityReasonDeleteInput
    {
        public int Id { get; set; }
    }
}

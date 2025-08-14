using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


[ApiController]
[Route("api/[controller]")]
public class OdooController : ControllerBase
{
    private readonly HttpClient _httpClient;
    private const string OdooApiUrl = "https://sigmaworldwide.io/web/dataset/call_kw/mrp.production/web_search_read";

    // private const string OdooApiNameEmployer = "https://sigmaworldwide.io/web/dataset/call_kw/hr.employee/web_search_read";

    public OdooController(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }


    // // API GET tên nhân viên dựa vào SVN Code

    // [HttpGet("SearchName")]
    // public async Task<IActionResult> SearchName([FromQuery] string nameEmployer)
    // {
    //     string nameJson = $@"
    //         {{
    //             ""id"": 13,
    //             ""jsonrpc"": ""2.0"",
    //             ""method"": ""call"",
    //             ""params"": {{
    //                 ""model"": ""hr.employee"",
    //                 ""method"": ""web_search_read"",
    //                 ""args"": [],
    //                 ""kwargs"": {{
    //                     ""limit"": 80,
    //                     ""offset"": 0,
    //                     ""order"": """",
    //                     ""context"": {{
    //                         ""lang"": ""vi_VN"",
    //                         ""tz"": ""Asia/Ho_Chi_Minh"",
    //                         ""uid"": 2,
    //                         ""allowed_company_ids"": [1],
    //                         ""bin_size"": true,
    //                         ""chat_icon"": true
    //                     }},
    //                     ""count_limit"": 10001,
    //                     ""domain"": [
    //                         ""|"",
    //                         [""name"", ""ilike"", ""{nameEmployer}""],
    //                     ],
    //                     ""fields"": [
    //                         ""__last_update"", ""activity_exception_decoration"", ""activity_exception_icon"",
    //                         ""activity_state"", ""activity_summary"", ""activity_type_icon"", ""activity_type_id"",
    //                         ""id"", ""hr_presence_state"", ""seniority_years"", ""user_id"", ""user_partner_id"",
    //                         ""hr_icon_display"", ""show_hr_icon_display"", ""image_128"", ""current_leave_id"",
    //                         ""current_leave_state"", ""leave_date_from"", ""leave_date_to"", ""is_absent"",
    //                         ""image_1024"", ""name"", ""job_title"", ""category_ids"", ""work_email"",
    //                         ""work_phone"", ""seniority_message"", ""activity_ids""
    //                     ]
    //                 }}
    //             }}
    //         }}";

    //     var jsonContentName = new StringContent(nameJson, Encoding.UTF8, "application/json");

    //     try
    //     {
    //         // Tạo request riêng để thêm cookie cho request này thôi
    //         using var request = new HttpRequestMessage(HttpMethod.Post, OdooApiNameEmployer)
    //         {
    //             Content = jsonContentName
    //         };

    //         // Thêm cookie cho request này
    //         var cookie = "frontend_lang=vi_VN; session_id=cc48f91eadf3641a29d02e5b5ebd7e623ceeac11; cids=1";
    //         request.Headers.Add("Cookie", cookie);

    //         var response = await _httpClient.SendAsync(request);
    //         response.EnsureSuccessStatusCode();

    //         var responseBody = await response.Content.ReadAsStringAsync();

    //         // Trích xuất product code từ response
    //         var name = ExtractNameemployer(responseBody);

    //         if (!string.IsNullOrEmpty(name))
    //         {
    //             return Ok(new { name = name });
    //         }
    //         else
    //         {
    //             return NotFound(new { message = "Product code not found" });
    //         }
    //     }
    //     catch (HttpRequestException ex)
    //     {
    //         return StatusCode(500, new { message = "Error calling Odoo API", details = ex.Message });
    //     }

    // }

    // private string ExtractNameemployer(string jsonResponse)
    // {
    //     try
    //     {
    //         var jsonObj = JObject.Parse(jsonResponse);
    //         var records = jsonObj["result"]?["records"] as JArray;

    //         if (records != null && records.Count > 0)
    //         {
    //             var firstRecord = records[0];
    //             var nameValue = firstRecord["name"]?.ToString();

    //             if (!string.IsNullOrEmpty(nameValue))
    //             {
    //                 return nameValue;
    //             }
    //         }

    //         return null;
    //     }
    //     catch (Exception ex)
    //     {
    //         Console.WriteLine($"Error extracting name employer: {ex.Message}");
    //         return null;
    //     }
    // }

    [HttpGet("search")]
    public async Task<IActionResult> Search([FromQuery] string productionCode)
    {
        // Tạo body JSON bằng chuỗi nội suy ($@"") như bạn muốn
        string finalJson = $@"
            {{
                ""id"": 555555555,
                ""jsonrpc"": ""2.0"",
                ""method"": ""call"",
                ""params"": {{
                    ""model"": ""mrp.production"",
                    ""method"": ""web_search_read"",
                    ""args"": [],
                    ""kwargs"": {{
                        ""limit"": 80,
                        ""offset"": 0,
                        ""order"": """",
                        ""context"": {{
                            ""lang"": ""vi_VN"",
                            ""tz"": ""Asia/Ho_Chi_Minh"",
                            ""uid"": 2,
                            ""allowed_company_ids"": [1],
                            ""bin_size"": true,
                            ""default_company_id"": 1
                        }},
                        ""count_limit"": 10001,
                        ""domain"": [
                            ""&"",
                            [""picking_type_id.active"", ""="", true],
                            ""&"",
                            [""state"", ""in"", [""draft"", ""confirmed"", ""progress"", ""to_close""]],
                            ""|"",
                            [""name"", ""ilike"", ""{productionCode}""],
                            [""origin"", ""ilike"", ""xxxxxxxxx""]
                        ],
                        ""fields"": [
                            ""activity_exception_decoration"", ""activity_exception_icon"", ""activity_state"",
                            ""activity_summary"", ""activity_type_icon"", ""activity_type_id"",
                            ""company_id"", ""product_uom_category_id"", ""priority"", ""message_needaction"",
                            ""name"", ""date_planned_start"", ""date_deadline"", ""product_id"",
                            ""lot_producing_id"", ""bom_id"", ""activity_ids"", ""origin"", ""user_id"",
                            ""components_availability_state"", ""components_availability"",
                            ""reservation_state"", ""product_qty"", ""product_uom_id"",
                            ""production_duration_expected"", ""production_real_duration"",
                            ""progress"", ""state"", ""delay_alert_date"", ""json_popover""
                        ]
                    }}
                }}
            }}";

        var jsonContent = new StringContent(finalJson, Encoding.UTF8, "application/json");

        try
        {
            // Tạo request riêng để thêm cookie cho request này thôi
            using var request = new HttpRequestMessage(HttpMethod.Post, OdooApiUrl)
            {
                Content = jsonContent
            };

            // Thêm cookie cho request này
            var cookie = "frontend_lang=en_US; cids=1; session_id=aa3054475768115f9a742ff9131aab5c9c548b3c";
            request.Headers.Add("Cookie", cookie);

            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();

            var responseBody = await response.Content.ReadAsStringAsync();

            // Trích xuất product code từ response
            var productCode = ExtractProductCode(responseBody);

            if (!string.IsNullOrEmpty(productCode))
            {
                return Ok(new { productCode = productCode });
            }
            else
            {
                return NotFound(new { message = "Product code not found" });
            }
        }
        catch (HttpRequestException ex)
        {
            return StatusCode(500, new { message = "Error calling Odoo API", details = ex.Message });
        }
    }

    private string ExtractProductCode(string jsonResponse)
    {
        try
        {
            var jsonObj = JObject.Parse(jsonResponse);
            var records = jsonObj["result"]?["records"] as JArray;

            if (records != null && records.Count > 0)
            {
                var firstRecord = records[0];
                var productId = firstRecord["product_id"] as JArray;

                if (productId != null && productId.Count >= 2)
                {
                    var productDescription = productId[1]?.ToString();
                    if (!string.IsNullOrEmpty(productDescription))
                    {
                        var match = Regex.Match(productDescription, @"^\[([^\]]+)\]");
                        if (match.Success)
                        {
                            return match.Groups[1].Value;
                        }
                    }
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting product code: {ex.Message}");
            return null;
        }
    }
}


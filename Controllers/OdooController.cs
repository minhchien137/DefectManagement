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

    private const string OdooApiNameEmployer = "https://sigmaworldwide.io/web/dataset/call_kw/hr.employee/web_search_read";

    public OdooController(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }


    // API GET tên nhân viên dựa vào SVN Code

    [HttpGet("SearchName")]
    public async Task<IActionResult> SearchName([FromQuery] string nameEmployer)
    {
        if (string.IsNullOrWhiteSpace(nameEmployer))
        {
            return BadRequest(new { message = "nameEmployer parameter is required" });
        }

        // Escape đúng cách cho JSON
        var escapedName = JsonConvert.ToString(nameEmployer).Trim('"');

        string nameJson = $@"{{
        ""id"": 13,
        ""jsonrpc"": ""2.0"",
        ""method"": ""call"",
        ""params"": {{
            ""model"": ""hr.employee"",
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
                    ""chat_icon"": true
                }},
                ""count_limit"": 10001,
                ""domain"": [
                    [""category_ids"", ""ilike"", ""{escapedName}""]
                ],
                ""fields"": [
                    ""__last_update"", ""activity_exception_decoration"", ""activity_exception_icon"",
                    ""activity_state"", ""activity_summary"", ""activity_type_icon"", ""activity_type_id"",
                    ""id"", ""hr_presence_state"", ""seniority_years"", ""user_id"", ""user_partner_id"",
                    ""hr_icon_display"", ""show_hr_icon_display"", ""image_128"", ""current_leave_id"",
                    ""current_leave_state"", ""leave_date_from"", ""leave_date_to"", ""is_absent"",
                    ""image_1024"", ""name"", ""job_title"", ""category_ids"", ""work_email"",
                    ""work_phone"", ""seniority_message"", ""activity_ids""
                ]
            }}
        }}
    }}";

        try
        {
            JObject.Parse(nameJson);
        }
        catch (JsonException ex)
        {
            Console.WriteLine($"Invalid JSON: {ex.Message}");
            return BadRequest(new { message = "Invalid JSON format", details = ex.Message });
        }

        var jsonContentName = new StringContent(nameJson, Encoding.UTF8, "application/json");

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, OdooApiNameEmployer)
            {
                Content = jsonContentName
            };


            var cookie = "frontend_lang=vi_VN; session_id=cc48f91eadf3641a29d02e5b5ebd7e623ceeac11; cids=1";
            request.Headers.Add("Cookie", cookie);
            request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
            request.Headers.Add("Accept", "application/json, text/javascript, */*; q=0.01");
            request.Headers.Add("X-Requested-With", "XMLHttpRequest");

            Console.WriteLine($"Sending request to: {OdooApiNameEmployer}");
            Console.WriteLine($"Request headers: {string.Join(", ", request.Headers.Select(h => $"{h.Key}={h.Value.FirstOrDefault()}"))}");

            var response = await _httpClient.SendAsync(request);

            var responseBody = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"Response Status: {response.StatusCode}");
            Console.WriteLine($"Response Body: {responseBody}");

            if (!response.IsSuccessStatusCode)
            {
                // Trả về thông tin lỗi chi tiết
                return StatusCode((int)response.StatusCode, new
                {
                    message = $"Odoo API returned {response.StatusCode}",
                    statusCode = (int)response.StatusCode,
                    responseBody = responseBody,
                    requestUrl = OdooApiNameEmployer
                });
            }

            // Kiểm tra response có chứa error từ Odoo không
            try
            {
                var jsonResponse = JObject.Parse(responseBody);
                if (jsonResponse["error"] != null)
                {
                    Console.WriteLine($"Odoo API Error: {jsonResponse["error"]}");
                    return BadRequest(new
                    {
                        message = "Odoo API returned error",
                        error = jsonResponse["error"]
                    });
                }
            }
            catch (JsonException)
            {
                Console.WriteLine("Response is not valid JSON");
                return StatusCode(500, new { message = "Invalid JSON response from Odoo" });
            }

            var name = ExtractNameemployer(responseBody);

            if (!string.IsNullOrEmpty(name))
            {
                return Ok(new { nameEmployer = name });
            }
            else
            {
                return NotFound(new { message = "Name Employer not found", responseBody = responseBody });
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"HTTP Request Error: {ex.Message}");
            return StatusCode(500, new { message = "HTTP Request Error", details = ex.Message });
        }
        catch (TaskCanceledException ex)
        {
            Console.WriteLine($"Request Timeout: {ex.Message}");
            return StatusCode(408, new { message = "Request Timeout", details = ex.Message });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"General Error: {ex.Message}");
            return StatusCode(500, new { message = "General Error", details = ex.Message });
        }
    }

    private string ExtractNameemployer(string jsonResponse)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(jsonResponse))
            {
                Console.WriteLine("Empty response from Odoo");
                return null;
            }

            var jsonObj = JObject.Parse(jsonResponse);

            // Kiểm tra structure của response
            Console.WriteLine($"Response keys: {string.Join(", ", jsonObj.Properties().Select(p => p.Name))}");

            if (jsonObj["error"] != null)
            {
                Console.WriteLine($"Odoo returned error: {jsonObj["error"]}");
                return null;
            }

            var result = jsonObj["result"];
            if (result == null)
            {
                Console.WriteLine("No 'result' field in response");
                return null;
            }

            var records = result["records"] as JArray;
            Console.WriteLine($"Found {records?.Count ?? 0} records");

            if (records != null && records.Count > 0)
            {
                var firstRecord = records[0];
                Console.WriteLine($"First record keys: {string.Join(", ", ((JObject)firstRecord).Properties().Select(p => p.Name))}");

                var nameValue = firstRecord["name"]?.ToString();
                Console.WriteLine($"Extracted name: {nameValue}");
                if (!string.IsNullOrEmpty(nameValue) && nameValue.Contains("-"))
                {
                    nameValue = nameValue.Split('-').Last().Trim();
                }

                return nameValue;
            }

            Console.WriteLine("No records found in response");
            return null;
        }
        catch (JsonException ex)
        {
            Console.WriteLine($"JSON parsing error: {ex.Message}");
            Console.WriteLine($"Response was: {jsonResponse}");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting name employer: {ex.Message}");
            return null;
        }
    }

    [HttpGet("search")]
    public async Task<IActionResult> Search([FromQuery] string productionCode)
    {
        // Tạo body JSON bằng chuỗi nội suy ($@"") như bạn muốn
        string finalJson = $@"
            {{
                ""id"": 31,
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
                            ""|"",
                            [""name"", ""ilike"", ""{productionCode}""],
                            [""origin"", ""ilike"", ""xxxxxx""]
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
            var cookie = "frontend_lang=vi_VN; session_id=cc48f91eadf3641a29d02e5b5ebd7e623ceeac11; cids=1";
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


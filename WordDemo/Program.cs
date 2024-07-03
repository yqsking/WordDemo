using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            GetOcrTableCellReplaceRule();
        }


        /// <summary>
        /// 获取制表位word表格替换规则
        /// </summary>
        /// <returns></returns>
        private static List<WordTable> GetOcrTableCellReplaceRule()
        {
            var wordTables = new List<WordTable>();
            string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files/2023017197_update.json");
            //string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files/Roll-例（原稿）_20240626143608_2.json");
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine("Json文件不存在");
            }
            string pdfJson = File.ReadAllText(jsonPath);
            //string pdfJson = GetPdfJson().GetAwaiter().GetResult();
            //return wordTables;
            string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files/2023017197_update.docx");
            //string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files/Roll-例（原稿）.docx");
            if (!File.Exists(wordPath))
            {
                Console.WriteLine("Word文件不存在");
            }
            Application wordApp = new Application();
            Document doc = wordApp.Documents.Open(wordPath, ReadOnly: false, Visible: false);
            doc.Activate();
            try
            {
                wordTables = WordHelper.GetWordTableList(pdfJson, doc);

            }
            catch (Exception ex)
            {
                $"解析制表位表格失败,{ex.Message}".Console(ConsoleColor.Red);
            }
            finally
            {
                // doc.Save();
                doc.Close();
                wordApp.Quit();
            }
            return wordTables;

        }

        private static async Task<string> GetPdfJson()
        {
            "开始获取pdf json数据".Console(ConsoleColor.Yellow);
            var watch = new Stopwatch();
            watch.Start();
            string pdfJson = null;
            string configJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            string configJsonStr = File.ReadAllText(configJsonPath);
            var config = JObject.Parse(configJsonStr);
            string address = config["ocrConfig"]["address"].ToString();
            string port = config["ocrConfig"]["port"].ToString();
            string api_key = config["ocrConfig"]["api_key"].ToString();
            string secret_key = config["ocrConfig"]["secret_key"].ToString();
            var baseUrl = $"http://{address}:{port}";
            string getTokenUrl = $"{baseUrl}/{ApiConstant.GetToken}";
            HttpResponseMessage httpResponse = null;
            string token = string.Empty;
            HttpClient client = new HttpClient();

            #region 授权
            "正在获取OCR服务授权》》》".Console(ConsoleColor.Yellow);
            var getTokenRequest = new
            {
                api_key,
                secret_key
            };
            var getTokenRequestParamter = new StringContent(JsonConvert.SerializeObject(getTokenRequest), Encoding.UTF8, "application/json");
            try
            {
                httpResponse = await client.PostAsync(getTokenUrl, getTokenRequestParamter);
            }
            catch (Exception ex)
            {
                throw new Exception($"请求{ApiConstant.GetToken}失败,{ex.Message}");
            }
            if (httpResponse.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("获取token失败");
            }
            string getTokenResultJson = await httpResponse.Content.ReadAsStringAsync();
            JObject getTokenResult = JObject.Parse(getTokenResultJson);
            token = getTokenResult["access_token"].ToString();
            #endregion

            #region 创建OCR任务
            "正在创建OCR任务》》》".Console(ConsoleColor.Yellow);
            client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            string createTaskUrl = $"{baseUrl}/{ApiConstant.CreateTask}";
            //中文
            string pdfUrl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files/Roll-例（原稿）.pdf");
            var pdfStream = File.Open(pdfUrl, FileMode.Open);
            var createTaskRequest = new
            {
                async_task = true,
                files = pdfStream,
                is_add_table_up_down_lines = true,
                mode = 2,
                priority = 1,
                is_use_physical_lines = true,
                physical_lines_interval = 1,//y轴坐标误差小于0.1的算同行
                physical_lines_precise = 2,//0:按y轴下标取整计算物理行 1：按y轴下标取一位小数计算物理行 2：按y轴下标取2位小数计算物理行
            };
            var createTaskParamter = new MultipartFormDataContent
            {
                { new StringContent(createTaskRequest.async_task.ToString()), nameof(createTaskRequest.async_task) },
                { new StreamContent(pdfStream), nameof(createTaskRequest.files), Path.GetFileName(pdfUrl) },
                { new StringContent(createTaskRequest.is_add_table_up_down_lines.ToString()), nameof(createTaskRequest.is_add_table_up_down_lines) },
                { new StringContent(createTaskRequest.mode.ToString()), nameof(createTaskRequest.mode) },
                { new StringContent(createTaskRequest.priority.ToString()), nameof(createTaskRequest.priority) },
                { new StringContent(createTaskRequest.is_use_physical_lines.ToString()), nameof(createTaskRequest.is_use_physical_lines) },
                { new StringContent(createTaskRequest.physical_lines_interval.ToString()), nameof(createTaskRequest.physical_lines_interval) },
                { new StringContent(createTaskRequest.physical_lines_precise.ToString()), nameof(createTaskRequest.physical_lines_precise) }
            };
            try
            {
                httpResponse = await client.PostAsync(createTaskUrl, createTaskParamter);
            }
            catch (Exception ex)
            {
                throw new Exception($"请求{ApiConstant.CreateTask}失败,{ex.Message}");
            }
            if (httpResponse.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("创建OCR识别任务失败");
            }
            string createTaskResultJson = await httpResponse.Content.ReadAsStringAsync();
            JObject createTaskResult = JObject.Parse(createTaskResultJson);
            string taskId = createTaskResult["task_id"].ToString();
            int selectNumber = 0;
            while (pdfJson == null)
            {
                selectNumber++;
                client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                try
                {
                    $"第{selectNumber}次获取OCR任务状态》》》".Console(ConsoleColor.Yellow);
                    string getTaskStatusUrl = $"{baseUrl}/{ApiConstant.GetTaskStatus}?task_id=" + taskId;
                    httpResponse = await client.GetAsync(getTaskStatusUrl);
                }
                catch (Exception ex)
                {
                    throw new Exception("获取OCR任务状态异常，" + ex.Message);
                }
                var getTaskStatusResultJson = await httpResponse.Content.ReadAsStringAsync();
                JObject getTaskStatusResult = JObject.Parse(getTaskStatusResultJson);
                var taskStatus = getTaskStatusResult["task_status"].ToString();
                if (taskStatus != "over")
                {
                    Thread.Sleep(1000);
                    continue;
                }
                string getTaskResultUrl = $"{baseUrl}/{ApiConstant.GetTaskResult}?task_id={taskId}";
                client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                try
                {
                    $"获取OCR任务结果》》》".Console(ConsoleColor.Yellow);
                    httpResponse = await client.GetAsync(getTaskResultUrl);
                }
                catch (Exception ex)
                {
                    throw new Exception("获取OCR任务结果异常，" + ex.Message);
                }
                string getTaskResultJson = await httpResponse.Content.ReadAsStringAsync();
                JObject getTaskResult = JObject.Parse(getTaskResultJson);
                pdfJson = getTaskResult["task_result"].ToString();
                string jsonFileName = Path.GetFileName(pdfUrl).Split('.').FirstOrDefault() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                string jsonUrl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Files/{jsonFileName}_{createTaskRequest.physical_lines_precise}.json");
                File.WriteAllText(jsonUrl, pdfJson);
                "json文件获取完毕".Console(ConsoleColor.Yellow);

            }
            #endregion

            watch.Stop();
            $"获取pdf json数据结束，耗时{watch.ElapsedMilliseconds / 1000}秒".Console(ConsoleColor.Yellow);
            return pdfJson;
        }
    }
}

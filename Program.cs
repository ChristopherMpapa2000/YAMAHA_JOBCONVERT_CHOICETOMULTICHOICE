using JobConvert_ChoiceToMultichoice.Item;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using WOLF_START_MigrateDAR;

namespace JobConvert_ChoiceToMultichoice
{
    class Program
    {
        public static void Log(String iText)
        {
            string pathlog = ItemConfig.LogFile;
            String logFolderPath = System.IO.Path.Combine(pathlog, DateTime.Now.ToString("yyyyMMdd"));

            if (!System.IO.Directory.Exists(logFolderPath))
            {
                System.IO.Directory.CreateDirectory(logFolderPath);
            }
            String logFilePath = System.IO.Path.Combine(logFolderPath, DateTime.Now.ToString("yyyyMMdd") + ".txt");

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(logFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] listText = iText.Split('|').ToArray();

                    foreach (String s in listText)
                    {
                        sbLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {s}");
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing log file: {ex.Message}");
            }
        }
        public static void LogError(String iText)
        {

            string pathlog = ItemConfig.LogFile;
            String logFolderPath = System.IO.Path.Combine(pathlog, DateTime.Now.ToString("yyyyMMdd"));

            if (!System.IO.Directory.Exists(logFolderPath))
            {
                System.IO.Directory.CreateDirectory(logFolderPath);
            }
            String logFilePath = System.IO.Path.Combine(logFolderPath, DateTime.Now.ToString("yyyyMMdd") + "LogError.txt");

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(logFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] listText = iText.Split('|').ToArray();

                    foreach (String s in listText)
                    {
                        sbLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {s}");
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing log file: {ex.Message}");
            }
        }
        static void Main(string[] args)
        {
            Log("====== Start Process ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            Log(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));
            DataClasses1DataContext db = new DataClasses1DataContext(ItemConfig.dbConnectionString);
            if (db.Connection.State == ConnectionState.Open)
            {
                db.Connection.Close();
                db.Connection.Open();
            }
            db.Connection.Open();
            db.CommandTimeout = 0;
            var excelData = GetExcel.ReadExcelToDataTables(ItemConfig.PathFileExcel);
            if (excelData != null)
            {
                string sheetName = ItemConfig.SheetName_Migrate;
                if (excelData.ContainsKey(sheetName))
                {
                    DataTable sheetData = excelData[sheetName];

                    for (int index = 1; index < sheetData.Rows.Count; index++)
                    {
                        ConvertChoiceToMultichoice(db, sheetData.Rows[index]);
                    }
                }
            }
            Log("====== End Process Process ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
        }
        public static void ConvertChoiceToMultichoice(DataClasses1DataContext db, DataRow RowIndex)
        {
            string docno = RowIndex.ItemArray[0].ToString();
            string PageforReview = RowIndex.ItemArray[2].ToString();
            string PageforAcceptResult = RowIndex.ItemArray[4].ToString();
            string Figvalue = "Accept Countermeasures Idea,Follow up action,Close";
            var FigList = Figvalue.Split(',').ToList();
            List<TRNMemo> listmemo = new List<TRNMemo>();
            if (!string.IsNullOrEmpty(ItemConfig.Memoid))
            {
                listmemo = db.TRNMemos.Where(m => m.MemoId == Convert.ToInt32(ItemConfig.Memoid)).ToList();
            }
            else
            {
                listmemo = db.TRNMemos.Where(m => m.DocumentNo == docno).ToList();
            }
            Log("InsertMultichoice: " + listmemo.FirstOrDefault().DocumentNo + "|Memoid: " + listmemo.FirstOrDefault().MemoId);
            Console.WriteLine("InsertMultichoice: " + listmemo.FirstOrDefault().DocumentNo);
            if (listmemo.Count > 0)
            {
                foreach (var item in listmemo)
                {
                    JObject jsonAdvanceForm = JsonUtils.createJsonObject(item.MAdvancveForm);
                    JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                    foreach (JObject jItems in itemsArray)
                    {
                        JArray jLayoutArray = (JArray)jItems["layout"];
                        if (jLayoutArray.Count >= 1)
                        {

                            JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                            JObject jData = (JObject)jLayoutArray[0]["data"];
                            string type = (String)jTemplateL["type"];
                            if (type == "cb")
                            {
                                if (!string.IsNullOrEmpty(PageforReview))
                                {
                                    if ((String)jTemplateL["label"] == "Page for Review ")
                                    {
                                        if (jData.ContainsKey("value") && jData["value"] != null)
                                        {
                                            string oldvalue = jData["value"].ToString();
                                            string Newvalue = oldvalue + "," + PageforReview;
                                            var newList = Newvalue.Split(',').Distinct().ToList();
                                            var sortedList = newList.OrderBy(value => FigList.IndexOf(value)).Where(value => FigList.Contains(value)).ToList();
                                            string sortedNewValue = string.Join(",", sortedList);

                                            JArray itemArray = new JArray();
                                            JObject attribute = (JObject)jTemplateL["attribute"];
                                            JArray items = (JArray)attribute["items"];

                                            int indexitem = items.Count();
                                            string[] Boxitem = new string[indexitem];
                                            if (jData != null)
                                            {
                                                string itemvalue = oldvalue;
                                                string[] dataListArray = sortedNewValue.Split(',');

                                                for (int F = 0; F < indexitem; F++)
                                                {
                                                    string itemc = items[F]["item"].ToString();
                                                    if (dataListArray.Contains(itemc.Trim()))
                                                    {
                                                        Boxitem[F] = "Y";
                                                    }
                                                    else
                                                    {
                                                        Boxitem[F] = "N";
                                                    }
                                                }
                                                jData["value"] = sortedNewValue;
                                                jData["item"] = new JArray(Boxitem);
                                                string logbox = string.Join(", ", Boxitem);
                                                Log("Data Page for Review: " + sortedNewValue + "|Boxitem: " + logbox);
                                                Console.WriteLine("Data Page for Review: " + sortedNewValue);
                                            }
                                        }
                                    }
                                }
                                if (!string.IsNullOrEmpty(PageforAcceptResult))
                                {
                                    if ((String)jTemplateL["label"] == "Page for Accept Result    ")
                                    {
                                        if (jData.ContainsKey("value") && jData["value"] != null)
                                        {
                                            string oldvalue = jData["value"].ToString();
                                            string Newvalue = oldvalue + "," + PageforAcceptResult;
                                            var newList = Newvalue.Split(',').Distinct().ToList();
                                            var sortedList = newList.OrderBy(value => FigList.IndexOf(value)).Where(value => FigList.Contains(value)).ToList();
                                            string sortedNewValue = string.Join(",", sortedList);

                                            JArray itemArray = new JArray();
                                            JObject attribute = (JObject)jTemplateL["attribute"];
                                            JArray items = (JArray)attribute["items"];

                                            int indexitem = items.Count();
                                            string[] Boxitem = new string[indexitem];
                                            if (jData != null)
                                            {
                                                string itemvalue = oldvalue;
                                                string[] dataListArray = sortedNewValue.Split(',');

                                                for (int F = 0; F < indexitem; F++)
                                                {
                                                    string itemc = items[F]["item"].ToString();
                                                    if (dataListArray.Contains(itemc.Trim()))
                                                    {
                                                        Boxitem[F] = "Y";
                                                    }
                                                    else
                                                    {
                                                        Boxitem[F] = "N";
                                                    }
                                                }
                                                jData["value"] = sortedNewValue;
                                                jData["item"] = new JArray(Boxitem);
                                                string logbox = string.Join(", ", Boxitem);
                                                Log("Page for Accept Result: " + sortedNewValue + "|Boxitem: " + logbox);
                                                Console.WriteLine("Page for Accept Result: " + sortedNewValue);
                                            }
                                        }
                                    }
                                    if ((String)jTemplateL["label"] == "Page for Accept Result ")
                                    {
                                        if (jData.ContainsKey("value") && jData["value"] != null)
                                        {
                                            string oldvalue = jData["value"].ToString();
                                            string Newvalue = oldvalue + "," + PageforAcceptResult;
                                            var newList = Newvalue.Split(',').Distinct().ToList();
                                            var sortedList = newList.OrderBy(value => FigList.IndexOf(value)).Where(value => FigList.Contains(value)).ToList();
                                            string sortedNewValue = string.Join(",", sortedList);

                                            JArray itemArray = new JArray();
                                            JObject attribute = (JObject)jTemplateL["attribute"];
                                            JArray items = (JArray)attribute["items"];

                                            int indexitem = items.Count();
                                            string[] Boxitem = new string[indexitem];
                                            if (jData != null)
                                            {
                                                string itemvalue = oldvalue;
                                                string[] dataListArray = sortedNewValue.Split(',');

                                                for (int F = 0; F < indexitem; F++)
                                                {
                                                    string itemc = items[F]["item"].ToString();
                                                    if (dataListArray.Contains(itemc.Trim()))
                                                    {
                                                        Boxitem[F] = "Y";
                                                    }
                                                    else
                                                    {
                                                        Boxitem[F] = "N";
                                                    }
                                                }
                                                jData["value"] = sortedNewValue;
                                                jData["item"] = new JArray(Boxitem);
                                                string logbox = string.Join(", ", Boxitem);
                                                Log("Page for Accept Result: " + sortedNewValue + "|Boxitem: " + logbox);
                                                Console.WriteLine("Page for Accept Result: " + sortedNewValue);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    string Newmav = JsonConvert.SerializeObject(jsonAdvanceForm);
                    item.MAdvancveForm = Newmav;
                    db.SubmitChanges();
                    Log("===============================================================");
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobConvert_ChoiceToMultichoice.Item
{
    class ItemConfig
    {
        public static string dbConnectionString
        {
            get
            {
                var ServarName = ConfigurationManager.AppSettings["ServarName"];
                var Database = ConfigurationManager.AppSettings["Database"];
                var Username_database = ConfigurationManager.AppSettings["Username_database"];
                var Password_database = ConfigurationManager.AppSettings["Password_database"];
                var dbConnectionString = $"data source={ServarName};initial catalog={Database};persist security info=True;user id={Username_database};password={Password_database};Connection Timeout=200";

                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return string.Empty;
            }
        }
        public static string LogFile
        {
            get
            {
                var LogFile = System.Configuration.ConfigurationSettings.AppSettings["LogFile"];
                if (!string.IsNullOrEmpty(LogFile))
                {
                    return (LogFile);
                }
                return string.Empty;
            }
        }
        public static string Memoid
        {
            get
            {
                var Memoid = System.Configuration.ConfigurationSettings.AppSettings["Memoid"];
                if (!string.IsNullOrEmpty(Memoid))
                {
                    return (Memoid);
                }
                return string.Empty;
            }
        }
        public static string PathFileExcel
        {
            get
            {
                var PathFileExcel = ConfigurationManager.AppSettings["PathFileExcel"];

                if (!string.IsNullOrEmpty(PathFileExcel))
                {
                    return PathFileExcel;
                }
                return string.Empty;
            }
        }
        public static string SheetName_Migrate
        {
            get
            {
                var SheetName_Migrate = ConfigurationManager.AppSettings["SheetName_Migrate"];

                if (!string.IsNullOrEmpty(SheetName_Migrate))
                {
                    return SheetName_Migrate;
                }
                return string.Empty;
            }
        }
    }
}

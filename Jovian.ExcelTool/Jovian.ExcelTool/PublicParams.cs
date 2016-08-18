using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jovian.ExcelTool
{
    /// <summary>
    /// LPY 2016-8-4 添加
    /// 参数
    /// </summary>
    public class PublicParams
    {
        public static string OracleConnStr = GetAppConfigValueByString("OracleConnStr");//Oracle数据库连接语句
        public static string xmlHistory = "history.xml";
        public static string TemplateFileName = GetAppConfigValueByString("TemplateFileName");

        public static Type mainType = typeof(MainWindow);

        public static bool IsLogWrite = GetAppConfigValueByString("IsLogWrite") == "1" ? true : false;//日志是否输出
        public static string logFilePath = string.Format("{0}\\log\\{1}\\{2}\\{3}.txt", Environment.CurrentDirectory, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("yyyyMM"), DateTime.Now.ToString("yyyyMMdd"));
        /// <summary>
        /// LPY 2015-9-9 添加
        /// 根据Key值，从App.config文件中获取配置项的value
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetAppConfigValueByString(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml;

namespace Jovian.ExcelTool
{
    /// <summary>
    /// 
    /// </summary>
    public class XmlHelper
    {
        public static string xmlHistory = PublicParams.xmlHistory;

        public static void CheckXMLFile(string fileName)
        {
            XmlDocument dom = new XmlDocument();
            try
            {
                dom.Load(PublicParams.xmlHistory);
            }
            catch (Exception ex)
            {
                XmlHelper.CreateXMLFile(PublicParams.xmlHistory);
                LogHelper.WriteLog("不存在历史记录文件，已创建。原始消息：" + ex.Message);
            }
        }

        public static bool CreateXMLFile(string fileName)
        {
            bool isSuccess = false;
            try
            {
                XmlDocument dom = new XmlDocument();
                XmlDeclaration xmlDec = dom.CreateXmlDeclaration("1.0", "gb2312", null);
                XmlElement xe = dom.CreateElement("Root");
                dom.AppendChild(xmlDec);
                dom.AppendChild(xe);
                dom.Save(fileName);
                isSuccess = true;
            }
            catch (Exception)
            {
            }
            return isSuccess;
        }
    }
}

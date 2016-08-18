using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using Jovian.ExcelTool.classes;
using System.Collections.ObjectModel;

namespace Jovian.ExcelTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnSelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog() { Filter = "Excel Files 1997-2003 (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*" };
            var result = openFileDialog.ShowDialog();

            string fileName = string.Empty;
            string filePath = string.Empty;
            string fileExtension = string.Empty;

            ArrayList alTemplet = new ArrayList();
            if (result == true)
            {
                fileName = openFileDialog.FileName;
                filePath = System.IO.Path.GetFullPath(fileName);
                fileExtension = System.IO.Path.GetExtension(fileName);
            }
            else
            {
                return;
            }
            //System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog() { Filter = "Excel Files (*.xlsx)|*.xlsx" };
            //if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    MessageBox.Show(openFileDialog.FileName);
            //}

            //if (fileExtension == ".xls")
            //    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            //else
            //    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

            //OleDbConnection conn = new OleDbConnection(connStr);
            //conn.Open();

            //DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null,null,null,"TABLE"});//获取数据源的表定义元数据

            //foreach (DataRow row in dt.Rows)
            //{
            //    string tableName = row["Table_Name"].ToString();
            //    LogHelper.WriteLog(tableName);

            //    DataTable dtc = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tableName, null });
            //    foreach (DataRow col in dtc.Rows)
            //    {
            //        string colName = col["Column_Name"].ToString();
            //        LogHelper.WriteLog(colName + "  ");

            //        alTemplet.Add(colName);
                    
            //    }
            //}

            //string[] arrayTemplet = (string[])alTemplet.ToArray(typeof(string));


            //for (int i = 0; i < arrayTemplet.Length; i++)
            //{
            //    LogHelper.WriteLog(arrayTemplet[i].ToString());
            //}
            
            //OleDbDataAdapter da = new OleDbDataAdapter();
            //da.SelectCommand = new OleDbCommand(string.Format("Select * FROM [{0}]", "Sheet1$"), conn);
            //DataSet ds = new DataSet();
            //da.Fill(ds, "Table0");
            string sheetName = getFirstSheetNameByFileName(fileName);
            if (sheetName == null)
            {
                MessageBox.Show("所选文件中不含工作表！");
                return;
            }

            DataSet ds = ExcelHelper.GetDataSetBySheetName(fileName, sheetName);
            ObservableCollection<Offender> ocOffenders = new ObservableCollection<Offender>();
            ArrayList alColumnsName = ExcelHelper.GetColumnsNameBySheetName(fileName, sheetName);
            alColumnsName.Sort();
            if (CheckFile(alColumnsName))
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ocOffenders.Add(new Offender(ds.Tables[0].Rows[i]));
                }

                lvMain.ItemsSource = ocOffenders;
            }
            else
            {
                //string logFilePath = string.Format("{0}\\log\\{1}\\{2}\\{3}.txt", Environment.CurrentDirectory,DateTime.Now.ToString("yyyy"),DateTime.Now.ToString("yyyyMM"),DateTime.Now.ToString("yyyyMMdd")) ;
                //System.IO.File.Open(logFilePath, System.IO.FileMode.Open);
                MessageBox.Show("所选文件不符合模板要求，请检查确认。\n具体信息请查看日志文件！\n" + PublicParams.logFilePath, "提示");
                return;
            }

            
        }

        private bool CheckFile(ArrayList al)
        {
            string templateFileName = Environment.CurrentDirectory + PublicParams.TemplateFileName;
            ArrayList alTemplate = ExcelHelper.GetColumnsNameBySheetName(templateFileName, "Sheet1$");
            bool result = true;
            if (alTemplate.Count != al.Count)
            {
                result = false;
                LogHelper.WriteLog("所选文件列数与模板文件列数不相同，请检查确认");
            }
            for (int i = 0; i < al.Count; i++)
            {
                if (al[i].ToString() != alTemplate[i].ToString()) 
                {
                    result = false;
                    LogHelper.WriteLog(string.Format("所选文件“{0}”列与模板文件不同，应为：“{1}”请检查确认",al[i].ToString(),alTemplate[i].ToString()));
                }
            }
            return result;            
        }

        /// <summary>
        /// 返回所选文件的第一个工作表名称，一般为Sheet1$
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string getFirstSheetNameByFileName(string fileName)
        {
            DataTable dtSchema = ExcelHelper.GetExcelSchemaTable(fileName);
            if (dtSchema.Rows.Count <= 0)
            {
                LogHelper.WriteLog("所选文件中不含工作表");
                return null;
            }
            return dtSchema.Rows[0]["Table_Name"].ToString();
        }

        private void CheckData(ref ObservableCollection<Offender> oc)
        {
            string labIDCondition = GetConditionOfLabID(oc);
            string cardIDCondition = GetConditionOfCardID(oc);

            bool isExists = false;

            string strQuery = "select ID_CARD_NO,LAB_ID from offender o where 1=1 ";
            if (labIDCondition != "")
                strQuery += " and LAB_ID in ("+labIDCondition+") ";
            if (cardIDCondition != "")
                strQuery += " and ID_CARD_NO in (" + cardIDCondition + ") ";

            DataTable dtResult = OracleHelper.ExecuteQuery(strQuery).Tables[0];
            if (dtResult.Rows.Count > 0)
            {
                isExists = true;
                for (int i = 0; i < dtResult.Rows.Count; i++)
                {
                    string cardID=dtResult.Rows[i]["ID_CARD_NO"].ToString();
                    LogHelper.WriteLog("身份证号：" + cardID + "的人员信息已存在，未能导入数据库，请检查确认！");
                    for (int j = 0; j < oc.Count; j++)
                    {
                        if (oc[j].CardID == cardID)
                        {
                            oc.RemoveAt(j);
                            break;
                        }
                    }
                }
            }

            if (isExists)
            {
                
                MessageBox.Show(string.Format("部分人员信息未能导入，具体查看日志记录！\n{0}", PublicParams.logFilePath), "提示");
            }

        }

        private string GetConditionOfLabID(ObservableCollection<Offender> oc)
        {
            StringBuilder result = new StringBuilder();
            if (oc.Count <= 0)
                return null;
            for (int i = 0; i < oc.Count-1; i++)
            {
                result.AppendFormat("'{0}',", oc[i].LabID.ToString());
            }
            result.Append ("'"+ oc[oc.Count - 1].LabID.ToString()+"'");
            return result.ToString();
        }

        private string GetConditionOfCardID(ObservableCollection<Offender> oc)
        {
            StringBuilder result = new StringBuilder();
            if (oc.Count <= 0)
                return null;
            for (int i = 0; i < oc.Count - 1; i++)
            {
                result.AppendFormat("'{0}',", oc[i].CardID.ToString());
            }
            result.Append("'" + oc[oc.Count - 1].CardID.ToString()+"'");
            return result.ToString();
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            App.Current.Shutdown();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            ObservableCollection<Offender> ocOffenders = lvMain.ItemsSource as ObservableCollection<Offender>;
            CheckData(ref ocOffenders);//导入数据库前对数据做检查
            ArrayList alInsertSql = new ArrayList();

            foreach (Offender offender in ocOffenders)
            {
                string sqlInsert = string.Format(@"insert into OFFENDER 
(
remark,
lab_id,
personnel_name,
gender,
nationality,
birth_date,
personnel_type,
native_place_addr,
residence_addr,
id_card_no,
case_name,
case_property,
RECEPTION_REGIONALISM,
RECEPTION_ORG_NAME,
RECEPTION_MAN,
RECEPTION_TEL,
TRANSFER_DATE,
ID,
CONSIGNMENT_ID,
INIT_SERVER_NO,
PERSONNEL_NO
) values ({0})", GetInsertValues(offender));

                alInsertSql.Add(sqlInsert);

                
                //int flag = OracleHelper.ExecuteSql(sqlInsert);
                //if (flag > 0)
                //{
                //    LogHelper.WriteLog( string.Format("{0} 导入数据 {1} 条",DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"),flag));
                //    MessageBox.Show("导入成功，更多导入信息请查看日志文件！\n" + PublicParams.logFilePath, "提示");
                //}
                //else
                //    MessageBox.Show("导入失败，具体原因请查看日志文件！\n" + PublicParams.logFilePath, "提示");
            }

            if (OracleHelper.ExecuteSqlTran(alInsertSql))
            {
                LogHelper.WriteLog(string.Format("{0} 导入数据 {1} 条", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), alInsertSql.Count));
                MessageBox.Show("导入成功，更多导入信息请查看日志文件！\n" + PublicParams.logFilePath, "提示");
            }
            else
                MessageBox.Show("导入失败，具体原因请查看日志文件！\n" + PublicParams.logFilePath, "提示");
        }

        private string GetInsertValues(Offender offender)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}'",
                offender.Remark,
                offender.LabID,
                offender.Name,
                offender.Sex,
                offender.Nation,
                offender.BirthDay==""?null:Convert.ToDateTime(offender.BirthDay).ToString(),
                "",//offender.Nation,
                offender.HomeAddr,
                offender.LiveAddr,
                offender.CardID,
                offender.CaseName,
                "",//offender.CaseNature,
                offender.CompanyAreaID,
                offender.SenderAddr,
                offender.Sender,
                offender.Tel,
                null,//DBNull.Value,//offender.SendTime,
                Guid.NewGuid().ToString().Replace("-",""),
                Guid.NewGuid().ToString().Replace("-", ""),
                "123456",//INIT_SERVER_NO
                Guid.NewGuid().ToString().Replace("-", "")
                );
            return sb.ToString();
        }

        
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using Oracle.DataAccess.Client;
using System.Collections;

namespace Jovian.ExcelTool
{
    /// <summary>
    /// LPY 2016-8-5 添加
    /// Oracle数据库操作帮助类
    /// </summary>
    public class OracleHelper
    {
        public static readonly string strOracleConn = PublicParams.OracleConnStr;

        #region 执行简单SQL语句
        /// <summary>
        /// LPY 2016-8-9
        /// 执行一条sql语句，判断是否有返回结果
        /// </summary>
        /// <param name="strSql"></param>
        /// <returns></returns>
        public static bool Exists(string strSql)
        {
            using (OracleConnection oraConn = new OracleConnection(strOracleConn))
            {
                using (OracleCommand oraCmd = new OracleCommand(strSql, oraConn))
                {
                    try
                    {
                        oraConn.Open();
                        object obj = oraCmd.ExecuteScalar();

                        if ((Object.Equals(obj, null)) || Object.Equals(obj, System.DBNull.Value))
                            return false;
                        else
                            return true;
                    }
                    catch (Exception ex)
                    {
                        oraConn.Close();
                        LogHelper.WriteLog(ex.Message);
                        return false;
                        //throw new Exception(ex.Message);
                    }
                    finally 
                    {
                        oraCmd.Dispose();
                        oraConn.Close();
                    }
                }
            }
        }

        /// <summary>
        /// LPY 2016-8-9
        /// 执行sql语句（增删改），返回影响的记录数
        /// </summary>
        /// <param name="strSql">sql语句</param>
        /// <param name="oraParams">受影响的记录数</param>
        /// <returns></returns>
        public static int ExecuteSql(string strSql,params OracleParameter[] oraParams)
        {
            using (OracleConnection oraConn = new OracleConnection(strOracleConn))
            {
                using (OracleCommand oraCmd = new OracleCommand(strSql, oraConn))
                {
                    try
                    {
                        PrepareCommand(oraCmd, oraConn, null, strSql, oraParams);
                        int rows = oraCmd.ExecuteNonQuery();
                        oraCmd.Parameters.Clear();
                        return rows;
                    }
                    catch (Exception ex)
                    {
                        oraConn.Close();
                        LogHelper.WriteLog(ex.Message);
                        throw new Exception(ex.Message);                        
                    }
                }
            }
        }

        /// <summary>
        /// LPY 2016-8-9
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="strSql"></param>
        /// <param name="oraParams"></param>
        /// <returns></returns>
        public static DataSet ExecuteQuery(string strSql, params OracleParameter[] oraParams)
        {
            using (OracleConnection oraConn = new OracleConnection(strOracleConn))
            {
                OracleCommand oraCmd = new OracleCommand();
                PrepareCommand(oraCmd, oraConn, null, strSql, oraParams);
                using (OracleDataAdapter da = new OracleDataAdapter(oraCmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds);
                        oraCmd.Parameters.Clear();
                    }
                    catch (Exception ex)
                    {
                        LogHelper.WriteLog(ex.Message);
                        throw;
                    }
                    finally
                    {
                        oraCmd.Dispose();
                        oraConn.Close();
                    }
                    return ds;
                }
            }
        }

        /// <summary>
        /// LPY 2016-8-9
        /// 根据列名和表名，获取最大编号值（已自动+1）
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static int GetMaxID(string fieldName, string tableName)
        {
            string strSql = "select max(" + fieldName + ")+1 from " + tableName;
            object obj = GetSingle(strSql);
            if (obj == null)
                return 1;
            else
                return int.Parse(obj.ToString());
        }


        public static object GetSingle(string strSql)
        {
            using (OracleConnection oraConn = new OracleConnection(strOracleConn))
            {
                using (OracleCommand oraCmd = new OracleCommand(strSql,oraConn))
                {
                    try
                    {
                        oraConn.Open();
                        object obj = oraCmd.ExecuteScalar();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, DBNull.Value)))
                            return null;
                        else
                            return obj;
                    }
                    catch (Exception ex)
                    {
                        oraConn.Close();
                        throw new Exception(ex.Message);
                    }
                    finally
                    {
                        oraCmd.Dispose();
                        oraConn.Close();
                    }
                }
            }
        }

        /// <summary>
        /// LPY 2016-8-10
        /// 执行多条sql语句
        /// </summary>
        /// <param name="strSqlList"></param>
        /// <returns></returns>
        public static bool ExecuteSqlTran(ArrayList strSqlList)
        {
            bool result = false;
            using (OracleConnection oraConn = new OracleConnection(strOracleConn))
            {
                oraConn.Open();
                OracleCommand oraCmd = new OracleCommand();
                oraCmd.Connection = oraConn;
                OracleTransaction oraTrans = oraConn.BeginTransaction();
                oraCmd.Transaction = oraTrans;
                try
                {
                    for (int i = 0; i < strSqlList.Count; i++)
                    {
                        string strSql = strSqlList[i].ToString();
                        if (strSql.Trim().Length >= 1)
                        {
                            oraCmd.CommandText = strSql;
                            oraCmd.ExecuteNonQuery();
                        }
                    }
                    oraTrans.Commit();
                    result = true;
                }
                catch (Exception ex)
                {
                    result = false;
                    oraTrans.Rollback();

                    throw new Exception(ex.Message);
                }
                finally
                {
                    oraCmd.Dispose();
                    oraConn.Close();
                }
                return result;
            }
        }

        /// <summary>
        /// LPY 2016-8-10
        /// 执行查询语句，返回OracleDataReader（注意：调用该方法后，一定要对返回的OracleDataReader进行Close）
        /// </summary>
        /// <param name="strSql"></param>
        /// <returns></returns>
        public static OracleDataReader GetOracleDataReader(string strSql,params OracleParameter[] oraParams)
        {
            OracleConnection oraConn = new OracleConnection(strOracleConn);
            OracleCommand oraCmd = new OracleCommand(strSql, oraConn);
            try
            {
                //oraConn.Open();
                PrepareCommand(oraCmd, oraConn, null, strSql, oraParams);
                OracleDataReader oraDR = oraCmd.ExecuteReader(CommandBehavior.CloseConnection);
                oraCmd.Parameters.Clear();
                return oraDR;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(ex.Message);
                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region 存储过程操作方法

        /// <summary>
        /// LPY 2016-8-10
        /// 执行存储过程，返回数据集
        /// </summary>
        /// <param name="procName"></param>
        /// <param name="oraParams"></param>
        /// <returns></returns>
        public static DataSet ExecuteProc(string procName, params OracleParameter[] oraParams)
        {
            using (OracleConnection oraConn = new OracleConnection(strOracleConn))
            {
                DataSet ds = new DataSet();
                try
                {
                    oraConn.Open();
                    OracleDataAdapter da = new OracleDataAdapter();
                    da.SelectCommand = BuildQueryCommand(oraConn, procName, oraParams);
                    da.Fill(ds);
                    return ds;
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(ex.Message);
                    throw new Exception(ex.Message) ;
                }
                finally
                {
                    oraConn.Close();
                }
            }
        }
        #endregion


        #region 辅助方法
        private static void PrepareCommand(OracleCommand oraCmd, OracleConnection oraConn, OracleTransaction oraTrans, string cmdText, OracleParameter[] oraParams)
        {
            if (oraConn.State != ConnectionState.Open) 
                oraConn.Open();

            oraCmd.Connection = oraConn;
            oraCmd.CommandText = cmdText;
            if (oraTrans != null)
                oraCmd.Transaction = oraTrans;

            oraCmd.CommandType = CommandType.Text;

            if (oraParams != null)
            {
                foreach (OracleParameter param in oraParams)
                {
                    oraCmd.Parameters.Add(param);
                }
            }
        }


        private static OracleCommand BuildQueryCommand(OracleConnection oraConn, string procName, params IDataParameter[] oraParams)
        {
            OracleCommand oraCmd = new OracleCommand(procName, oraConn);
            oraCmd.CommandType = CommandType.StoredProcedure;
            if (oraParams != null)
            {
                foreach (OracleParameter param in oraParams)
                {
                    oraCmd.Parameters.Add(param);
                }
            }
            return oraCmd;
        }
        #endregion
    }
}

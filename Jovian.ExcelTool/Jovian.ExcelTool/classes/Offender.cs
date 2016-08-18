using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jovian.ExcelTool.classes
{
    /// <summary>
    /// 作者：卢平义
    /// 时间：2016-08-05 16:58:48
    /// 说明：Offender
    /// </summary>
    public class Offender 
    {
        public Offender()
        {
        }
        
        private int id;

        /// <summary>
        /// 编号
        /// </summary>
        public int Id 
        {
            get { return id; }
            set { id = value; }
        }
        private string remark;

        /// <summary>
        /// 备注
        /// </summary>
        public string Remark 
        {
            get { return remark; }
            set { remark = value; }
        }
        private string labID;

        /// <summary>
        /// 实验室编号
        /// </summary>
        public string LabID 
        {
            get { return labID; }
            set { labID = value; }
        }
        private string name;

        /// <summary>
        /// 人员姓名
        /// </summary>
        public string Name 
        {
            get { return name; }
            set { name = value; }
        }
        private string sex;

        /// <summary>
        /// 性别
        /// </summary>
        public string Sex 
        {
            get { return sex; }
            set { sex = value; }
        }
        private string nation;

        /// <summary>
        /// 民族
        /// </summary>
        public string Nation 
        {
            get { return nation; }
            set { nation = value; }
        }
        private string birthDay;

        /// <summary>
        /// 出生日期
        /// </summary>
        public string BirthDay 
        {
            get { return birthDay; }
            set { birthDay = value; }
        }
        private string peopleType;

        /// <summary>
        /// 人员类型
        /// </summary>
        public string PeopleType 
        {
            get { return peopleType; }
            set { peopleType = value; }
        }
        private string homeAddr;

        /// <summary>
        /// 户籍地址
        /// </summary>
        public string HomeAddr 
        {
            get { return homeAddr; }
            set { homeAddr = value; }
        }
        private string liveAddr;

        /// <summary>
        /// 居住地址
        /// </summary>
        public string LiveAddr 
        {
            get { return liveAddr; }
            set { liveAddr = value; }
        }
        private string cardID;

        /// <summary>
        /// 身份证号
        /// </summary>
        public string CardID 
        {
            get { return cardID; }
            set { cardID = value; }
        }
        private string caseName;

        /// <summary>
        /// 涉案名称
        /// </summary>
        public string CaseName 
        {
            get { return caseName; }
            set { caseName = value; }
        }
        private string caseNature;

        /// <summary>
        /// 涉案性质
        /// </summary>
        public string CaseNature 
        {
            get { return caseNature; }
            set { caseNature = value; }
        }
        private string companyAreaID;

        /// <summary>
        /// 委托单位行政区划
        /// </summary>
        public string CompanyAreaID 
        {
            get { return companyAreaID; }
            set { companyAreaID = value; }
        }
        private string senderAddr;

        /// <summary>
        /// 送检人通讯地址
        /// </summary>
        public string SenderAddr 
        {
            get { return senderAddr; }
            set { senderAddr = value; }
        }
        private string sender;

        /// <summary>
        /// 送检人姓名
        /// </summary>
        public string Sender 
        {
            get { return sender; }
            set { sender = value; }
        }
        private string tel;

        /// <summary>
        /// 联系电话
        /// </summary>
        public string Tel 
        {
            get { return tel; }
            set { tel = value; }
        }
        private string sendTime;

        /// <summary>
        /// 送检时间
        /// </summary>
        public string SendTime 
        {
            get { return sendTime; }
            set { sendTime = value; }
        }

        public Offender(string _remark, string _labID, string _name, string _sex)//, string _nation, string _birthDay, string _peopleType, string _homeAddr, string _liveAddr, string _cardID, string _caseName)
        {
            Remark = _remark;
            LabID = _labID;
            Name = _name;
            Sex = _sex;
        }

        public Offender(DataRow dr)
        {
            Remark = dr["备注"].ToString();
            LabID = dr["实验室编号"].ToString();
            Name = dr["人员姓名"].ToString();
            Sex= dr["性别"].ToString();
            Nation= dr["民族"].ToString();
            BirthDay= dr["出生日期"].ToString();
            PeopleType= dr["人员类型"].ToString();
            HomeAddr= dr["户籍地址"].ToString();
            LiveAddr= dr["居住地址"].ToString();
            CardID= dr["身份证号"].ToString();
            CaseName= dr["涉案名称"].ToString();
            CaseNature= dr["涉案性质"].ToString();
            CompanyAreaID= dr["委托单位行政区划"].ToString();
            SenderAddr= dr["送检人通讯地址"].ToString();
            Sender= dr["送检人姓名"].ToString();
            Tel= dr["联系电话"].ToString();
            SendTime = dr["送检时间"].ToString();
        }
    }
}













using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jovian.ExcelTool.classes
{
    /// <summary>
    /// ���ߣ�¬ƽ��
    /// ʱ�䣺2016-08-05 16:58:48
    /// ˵����Offender
    /// </summary>
    public class Offender 
    {
        public Offender()
        {
        }
        
        private int id;

        /// <summary>
        /// ���
        /// </summary>
        public int Id 
        {
            get { return id; }
            set { id = value; }
        }
        private string remark;

        /// <summary>
        /// ��ע
        /// </summary>
        public string Remark 
        {
            get { return remark; }
            set { remark = value; }
        }
        private string labID;

        /// <summary>
        /// ʵ���ұ��
        /// </summary>
        public string LabID 
        {
            get { return labID; }
            set { labID = value; }
        }
        private string name;

        /// <summary>
        /// ��Ա����
        /// </summary>
        public string Name 
        {
            get { return name; }
            set { name = value; }
        }
        private string sex;

        /// <summary>
        /// �Ա�
        /// </summary>
        public string Sex 
        {
            get { return sex; }
            set { sex = value; }
        }
        private string nation;

        /// <summary>
        /// ����
        /// </summary>
        public string Nation 
        {
            get { return nation; }
            set { nation = value; }
        }
        private string birthDay;

        /// <summary>
        /// ��������
        /// </summary>
        public string BirthDay 
        {
            get { return birthDay; }
            set { birthDay = value; }
        }
        private string peopleType;

        /// <summary>
        /// ��Ա����
        /// </summary>
        public string PeopleType 
        {
            get { return peopleType; }
            set { peopleType = value; }
        }
        private string homeAddr;

        /// <summary>
        /// ������ַ
        /// </summary>
        public string HomeAddr 
        {
            get { return homeAddr; }
            set { homeAddr = value; }
        }
        private string liveAddr;

        /// <summary>
        /// ��ס��ַ
        /// </summary>
        public string LiveAddr 
        {
            get { return liveAddr; }
            set { liveAddr = value; }
        }
        private string cardID;

        /// <summary>
        /// ���֤��
        /// </summary>
        public string CardID 
        {
            get { return cardID; }
            set { cardID = value; }
        }
        private string caseName;

        /// <summary>
        /// �永����
        /// </summary>
        public string CaseName 
        {
            get { return caseName; }
            set { caseName = value; }
        }
        private string caseNature;

        /// <summary>
        /// �永����
        /// </summary>
        public string CaseNature 
        {
            get { return caseNature; }
            set { caseNature = value; }
        }
        private string companyAreaID;

        /// <summary>
        /// ί�е�λ��������
        /// </summary>
        public string CompanyAreaID 
        {
            get { return companyAreaID; }
            set { companyAreaID = value; }
        }
        private string senderAddr;

        /// <summary>
        /// �ͼ���ͨѶ��ַ
        /// </summary>
        public string SenderAddr 
        {
            get { return senderAddr; }
            set { senderAddr = value; }
        }
        private string sender;

        /// <summary>
        /// �ͼ�������
        /// </summary>
        public string Sender 
        {
            get { return sender; }
            set { sender = value; }
        }
        private string tel;

        /// <summary>
        /// ��ϵ�绰
        /// </summary>
        public string Tel 
        {
            get { return tel; }
            set { tel = value; }
        }
        private string sendTime;

        /// <summary>
        /// �ͼ�ʱ��
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
            Remark = dr["��ע"].ToString();
            LabID = dr["ʵ���ұ��"].ToString();
            Name = dr["��Ա����"].ToString();
            Sex= dr["�Ա�"].ToString();
            Nation= dr["����"].ToString();
            BirthDay= dr["��������"].ToString();
            PeopleType= dr["��Ա����"].ToString();
            HomeAddr= dr["������ַ"].ToString();
            LiveAddr= dr["��ס��ַ"].ToString();
            CardID= dr["���֤��"].ToString();
            CaseName= dr["�永����"].ToString();
            CaseNature= dr["�永����"].ToString();
            CompanyAreaID= dr["ί�е�λ��������"].ToString();
            SenderAddr= dr["�ͼ���ͨѶ��ַ"].ToString();
            Sender= dr["�ͼ�������"].ToString();
            Tel= dr["��ϵ�绰"].ToString();
            SendTime = dr["�ͼ�ʱ��"].ToString();
        }
    }
}













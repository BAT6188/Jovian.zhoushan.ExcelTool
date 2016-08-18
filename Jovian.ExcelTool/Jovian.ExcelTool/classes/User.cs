using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jovian.ExcelTool.classes
{
    public class User
    {
        public User()
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
        private string name;

        /// <summary>
        /// 用户名
        /// </summary>
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        private string pwd;

        /// <summary>
        /// 密码
        /// </summary>
        public string Pwd
        {
            get { return pwd; }
            set { pwd = value; }
        }
        private int remember;

        /// <summary>
        /// 记住密码
        /// </summary>
        public int Remember
        {
            get { return remember; }
            set { remember = value; }
        }
        private int autoLogin;

        /// <summary>
        /// 自动登录
        /// </summary>
        public int AutoLogin
        {
            get { return autoLogin; }
            set { autoLogin = value; }
        }

        public User(string name,string pwd,int remember,int autoLogin)
        {
            this.Name = name;
            this.Pwd = pwd;
            this.Remember = remember;
            this.AutoLogin = autoLogin;
        }
    }
}

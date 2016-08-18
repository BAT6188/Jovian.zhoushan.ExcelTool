using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using System.Xml;
using Jovian.ExcelTool.classes;

namespace Jovian.ExcelTool
{
    /// <summary>
    /// Login.xaml 的交互逻辑
    /// </summary>
    public partial class Login : Window
    {
        public static string connStr = "Data Source=192.168.1.99/ORCL;User ID=lpy;PassWord=lpy";
        public static string connStr2 = "";

        public Login()
        {
            InitializeComponent();
            LoadHistory();
            //LogHelper.WriteLog("程序启动成功！");
        }

        private void LoadHistory()
        {
            XmlDocument dom = new XmlDocument();
            XmlHelper.CheckXMLFile(PublicParams.xmlHistory);
            dom.Load(PublicParams.xmlHistory);
            XmlNode nodeRoot = dom.SelectSingleNode("Root");
            if (nodeRoot.ChildNodes.Count <= 0)
                return;

            for (int i = 0; i < nodeRoot.ChildNodes.Count; i++)
            {
                ComboBoxItem cbi = new ComboBoxItem() { Content = nodeRoot.ChildNodes[i]["Name"].InnerText, Style = this.Resources["styleForComboxBoxOfUserName"] as Style,FontSize=24 };
                cbUserName.Items.Add(cbi);
                if (((XmlElement)nodeRoot.ChildNodes[i]).GetAttribute("Last") == "true")
                    cbUserName.SelectedItem = cbi;
            }
        }

        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            string name = cbUserName.Text.Trim();
            string pwd = pbPwd.Password.Trim();

            if (name == "" || pwd == "")
            {
                MessageBox.Show("用户名、密码均不能为空！","提示",MessageBoxButton.OK,MessageBoxImage.Information);
                return;
            }

            int remember = 0;
            int autoLogin = 0;
            if ((bool)cbRemberPwd.IsChecked)
                remember = 1;
            if ((bool)cbAutoLogin.IsChecked)
                autoLogin = 1;

            User user = new User(name,pwd,remember,autoLogin);
            //
            if (!CheckUserExistsAndUpdate(user))
                SaveUser(user);


            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Hide();

        }

        private void SaveUser(User user)
        {
            XmlDocument dom = new XmlDocument();
            XmlHelper.CheckXMLFile(PublicParams.xmlHistory);
            dom.Load(PublicParams.xmlHistory);
            XmlNode nodeRoot = dom.SelectSingleNode("Root");
            ClearLastUserName(ref nodeRoot);
            XmlElement xe = dom.CreateElement("User"); xe.SetAttribute("Last", "true");
            XmlElement name = dom.CreateElement("Name"); name.InnerText = user.Name.Trim(); xe.AppendChild(name);
            XmlElement pwd = dom.CreateElement("Pwd"); pwd.InnerText = CryptHelper.Encrypt(user.Pwd.Trim()); xe.AppendChild(pwd);
            XmlElement remember = dom.CreateElement("Remember"); remember.InnerText = user.Remember.ToString(); xe.AppendChild(remember);
            XmlElement autoLogin = dom.CreateElement("AutoLogin"); autoLogin.InnerText = user.AutoLogin.ToString(); xe.AppendChild(autoLogin);

            nodeRoot.AppendChild(xe);
            dom.Save(PublicParams.xmlHistory);
        }

        private bool CheckUserExistsAndUpdate(User user)
        {
            XmlDocument dom = new XmlDocument();
            XmlHelper.CheckXMLFile(PublicParams.xmlHistory);
            dom.Load(PublicParams.xmlHistory);
            XmlNode nodeRoot = dom.SelectSingleNode("Root");
            ClearLastUserName(ref nodeRoot);
            bool result = false;
            for (int i = 0; i < nodeRoot.ChildNodes.Count; i++)
            {
                if (nodeRoot.ChildNodes[i]["Name"].InnerText == user.Name)
                {
                    result = true;
                    ((XmlElement)nodeRoot.ChildNodes[i]).SetAttribute("Last", "true");
                    nodeRoot.ChildNodes[i]["Pwd"].InnerText = CryptHelper.Encrypt(user.Pwd);
                    nodeRoot.ChildNodes[i]["Remember"].InnerText = cbRemberPwd.IsChecked == true ? "1" : "0";
                    nodeRoot.ChildNodes[i]["AutoLogin"].InnerText = cbAutoLogin.IsChecked == true ? "1" : "0";
                    break;
                }
                else
                    continue;
            }
            dom.Save(PublicParams.xmlHistory);

            return result;            
        }
        
        private void btnDelOneUserName_Click(object sender, RoutedEventArgs e)
        {
            Button btnSender = sender as Button;
            MessageBox.Show(btnSender.Tag.ToString());
        }

        private void cbUserName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            User user = GetUserByName((cb.SelectedItem as ComboBoxItem).Content.ToString());
            if (user.Name == string.Empty)
                return;

            pbPwd.Password = user.Pwd;
            cbRemberPwd.IsChecked = user.Remember == 1 ? true : false;
            cbAutoLogin.IsChecked = user.AutoLogin == 1 ? true : false;
            //MessageBox.Show((cb.SelectedItem as ComboBoxItem).Content.ToString());
        }

        private User GetUserByName(string name)
        {
            XmlDocument dom = new XmlDocument();
            XmlHelper.CheckXMLFile(PublicParams.xmlHistory);
            dom.Load(PublicParams.xmlHistory);
            XmlNode nodeRoot = dom.SelectSingleNode("Root");

            User user = new User();
            for (int i = 0; i < nodeRoot.ChildNodes.Count; i++)
            {
                if (nodeRoot.ChildNodes[i]["Name"].InnerText == name)
                {
                    user.Name = nodeRoot.ChildNodes[i]["Name"].InnerText;
                    user.Remember = Convert.ToInt32(nodeRoot.ChildNodes[i]["Remember"].InnerText);
                    if (user.Remember == 1)
                        user.Pwd = CryptHelper.Decrypt(nodeRoot.ChildNodes[i]["Pwd"].InnerText);
                    user.AutoLogin = Convert.ToInt32(nodeRoot.ChildNodes[i]["AutoLogin"].InnerText);
                    break;
                }
                else
                    continue;
            }
            return user;
        }

        private void ClearLastUserName(ref XmlNode node)
        {
            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                ((XmlElement)node.ChildNodes[i]).SetAttribute("Last", "false");
            }
        }
    }
}

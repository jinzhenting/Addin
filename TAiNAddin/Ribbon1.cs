using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Drawing;

namespace TAiN
{
    public partial class TAiNAddin
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////服务器配置
        //ServerIP
        public string serverip;
        public string ServerIP
        {
            get { return serverip; }
            set { serverip = value; }
        }
        //UserID
        public string userid;
        public string UserID
        {
            get { return userid; }
            set { userid = value; }
        }
        //Password
        public string password;
        public string Password
        {
            get { return password; }
            set { password = value; }
        }
        //DataName
        public string dataname;
        public string DataName
        {
            get { return dataname; }
            set { dataname = value; }
        }
        //Server
        public string server;
        public string Server
        {
            get { return server; }
            set { server = value; }
        }

        //读取配置文件
        public void ServerSettings()
        {
            try
            {
                //注意XML内容区分大小写
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load("server.xml");
                //获取子节点
                XmlNodeList xmlNodeList = xmlDocument.SelectSingleNode("//server").ChildNodes;
                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    //遍历字段
                    XmlElement xmlElement = (XmlElement)xmlNode;
                    if (xmlElement.Name == "serverip") { serverip = xmlElement.GetAttribute("key").ToString(); }
                    if (xmlElement.Name == "dataname") { dataname = xmlElement.GetAttribute("key").ToString(); }
                    if (xmlElement.Name == "userid") { userid = xmlElement.GetAttribute("key").ToString(); }
                    if (xmlElement.Name == "password") { password = xmlElement.GetAttribute("key").ToString(); }
                }
                server = "Server=" + serverip + "; Initial Catalog=" + dataname + "; User ID=" + userid + "; Password=" + password;
            }
            catch (Exception exception) { MessageBox.Show(exception.ToString()); }
        }
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////服务器配置

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //服务器配置
            ServerSettings();
            button1.Image= Image.FromFile(@"8.png");
        }
        
        /// <summary>
        /// 返回dataTable
        /// </summary>
        /// <param name="select">SQL语句</param>
        /// <returns>DataTable</returns>
        public DataTable Select(string select)
        {
            try
            {
                DataTable dataTable = new DataTable();
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
                SqlConnection sqlConnection = new SqlConnection(Server);
                SqlCommand sqlCommand = new SqlCommand(select, sqlConnection);
                sqlConnection.Open();
                sqlDataAdapter.SelectCommand = sqlCommand;
                dataTable.Clear();
                sqlDataAdapter.Fill(dataTable);
                sqlConnection.Close();
                return dataTable;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                return null;
            }
        }


        //修改
        public void upData(string Updata)
        {
            try
            {
                SqlConnection sqlConnection = new SqlConnection(Server);
                sqlConnection.Open();
                //UPDATE = "update 表 set 列='值', 列='值', 列='值' WHERE  列='值'";
                SqlCommand sqlCommand = new SqlCommand(Updata, sqlConnection);
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                sqlConnection.Close();
            }
            catch (Exception exception) { MessageBox.Show(exception.ToString()); }
        }

        //插入
        public void inSert(string Insert)
        {
            try
            {
                SqlConnection sqlConnection = new SqlConnection(Server);
                sqlConnection.Open();
                //INSERT = "insert into 表 (列,列,列) values('值','值','值')";
                SqlCommand sqlCommand = new SqlCommand(Insert, sqlConnection);
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                sqlConnection.Close();
            }
            catch (Exception exception) { MessageBox.Show(exception.ToString()); }
        }

        //删除
        public void Delete(string Delete)
        {
            try
            {
                SqlConnection sqlConnection = new SqlConnection(Server);
                sqlConnection.Open();
                //DELETE = "delete FROM LINKS WHERE 列='值' AND  列='值'";
                SqlCommand sqlCommand = new SqlCommand(Delete, sqlConnection);
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            }
            catch (Exception exception) { MessageBox.Show(exception.ToString()); }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            DataTable datatable = Select("SELECT UserID AS 编号, UserName AS 用户名, Password AS 密码, Property AS 权限, Department AS 部门, Comment AS 描述, CreateTime AS 创建时间, Creator AS 创建人, UpdateTime AS 修改时间, Editor AS 修改人 FROM [User] ORDER BY UserName ASC");
            for (int i = 0; i < datatable.Rows.Count; i++)
            {
                string str = datatable.Rows[i]["用户名"].ToString();
                RibbonDropDownItem ribbonDropDownItem = str;
                comboBox1.Items.Add(ribbonDropDownItem);
                Application.DoEvents();
            }
            }
    }
}

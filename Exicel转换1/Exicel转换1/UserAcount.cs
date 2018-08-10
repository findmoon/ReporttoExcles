using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SQLite;
using System.Security.Cryptography;

namespace Exicel转换1
{
    //用户管理的类
    public class UserAcount
    {
        //用户权限
        //0, 最高权限，
        //1, 权限管理用户，设定或授予用户具有的权限。
        //2，专有管理员 可以编辑设定报价单的价格基准，在需要调整时进行调整
        //3，管理员，可以查看编辑生成报价，可以查询sqlite中的log 价格的记录
        //4, 普通用户（默认），进行评估报告的转换，不包含报价单CPP
        //5, 查看已转换的评估报告信息
        public string UserName { get; }//用户名唯一，与userid类似
        public int Level { get; }
        public string NickName { get; set; }
        //public string UserGroup { get; set; }//用户组的方式管理

        //状态，成功 1、user密码错误 2、过期 0
        public int accountState { get; set; }

        private static UserAcount userAcount = null;

        //单例模式，永远只有一个账户登录
        private UserAcount(string user, int level, int accountState,string nickname = null)
        {
            this.accountState = accountState;
            UserName = user;
            Level = level;
            NickName = nickname;
        }

        private UserAcount(int accountState)
        {
            this.accountState = accountState;
        }

        //权限对应关系
        //目前只实现 查看和修改 2 3 4的用户
        Dictionary<int, string> priviDict = new Dictionary<int, string>()
        {
            {0,"超级管理员" },
            {1,"权限管理员"},
            {2,"专有管理员" },
            {3,"管理用" },
            {4,"普通用户" },
            {5,"viewer" }
        };

        ///添加用户
        //创建\注册用户,独立于实例,默认level4，注册时默认，观众添加可指定
        /// <summary>
        /// 创建注册用户
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="pwd"></param>
        /// <param name="level"></param>
        /// <param name="nickname">默认为空</param>
        /// <returns>返回int，0：失败，1：成功，2：用户名重复</returns>
        public static int CreatAcount(string userName, string pwd, int level = 4, string nickname = null)
        {
            //if (Level>2)
            //{

            //}
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //查询是否存在用户
                        comm.CommandText = "SELECT * from UserTable WHERE User=@User";
                        comm.Parameters.AddWithValue("@User", userName);
                        conn.Open();
                        SQLiteDataReader dataReader = comm.ExecuteReader();
                        if (dataReader.HasRows)//用户名重复
                        {
                            return 2;
                        }
                        //关闭DataReader的链接
                        dataReader.Close();
                        comm.Parameters.Clear();//清空参数

                        //获取MD5加密后的密码
                        pwd = GetMd5_32bit(pwd);
                        if (nickname != null)
                        {
                            comm.CommandText = "INSERT INTO UserTable VALUES(@User,@Pwd,@Level,@NickName)";
                            comm.Parameters.AddWithValue("@NickName", nickname);
                        }
                        else
                        {
                            comm.CommandText = "INSERT INTO UserTable(User,Pwd,Level) VALUES(@User,@Pwd,@Level)";
                        }
                        comm.Parameters.AddWithValue("@User", userName);
                        comm.Parameters.AddWithValue("@Pwd", pwd);
                        comm.Parameters.AddWithValue("@Level", level);

                        if (conn.State == System.Data.ConnectionState.Closed)
                        {
                            conn.Open();
                        }
                        //插入用户
                        int affectRow = comm.ExecuteNonQuery();
                        if (affectRow != 1)
                        {
                            //创建用户失败
                            return 0;
                        }
                        else
                        {
                            return 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                return 0;
            }
        }

        //MD5加密密码
        public static string GetMd5_32bit(string pwd)
        {
            StringBuilder stringMD5Builder = new StringBuilder();
            //实例化MD5对象
            MD5 md5 = MD5.Create();

            //对字符串的二进制数组进行加密，获得的也是一个二进制数组
            byte[] strBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(pwd));

            //通过循环将字节类型的数组转换为字符串
            for (int i = 0; i < strBytes.Length; i++)
            {
                //将得到的字符串进行16进制类型格式化，(X)得到的是大写字母，使用小写字母(x)得到小写，X2表示获得16进制的2位数表示
                stringMD5Builder.Append(strBytes[i].ToString("X2"));

            }

            return stringMD5Builder.ToString();
        }

        //登录账户
        /// <summary>
        /// 登陆
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="pwd"></param>
        /// <returns></returns>
        public static UserAcount LoginAcount(string userName, string pwd)
        {
            if (pwd == "0" || string.IsNullOrEmpty(pwd))
            {
                return null;
            }
            if (userAcount != null)
            {
                return userAcount;
            }
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取MD5加密后的密码
                        pwd = GetMd5_32bit(pwd);
                        comm.CommandText = "SELECT Level,NickName,ExpireTime from UserTable WHERE User=@User AND Pwd=@Pwd";
                        comm.Parameters.AddWithValue("@User", userName);
                        comm.Parameters.AddWithValue("@Pwd", pwd);

                        conn.Open();
                        SQLiteDataReader dataReader = comm.ExecuteReader();
                        if (dataReader.HasRows)
                        {
                            dataReader.Read();
                            
                            if (DateTime.Compare(dataReader.GetDateTime(2),DateTime.Now)<0)
                            {
                                //重置密码  为0, 
                                //获取 加密密码
                                comm.Parameters.Clear();
                                dataReader.Close();
                                string newPwd = GetMd5_32bit("0");
                                string expireTime = "1970-01-01";

                                comm.CommandText = "UPDATE UserTable SET Pwd=@NewPwd,ExpireTime=@ExpireTime WHERE User=@User";
                                comm.Parameters.AddWithValue("@NewPwd", newPwd);
                                comm.Parameters.AddWithValue("@User", userName);
                                comm.Parameters.AddWithValue("@ExpireTime", expireTime);
                                if (conn.State== System.Data.ConnectionState.Closed)
                                {
                                    conn.Open();
                                }
                                //int affectInt = comm.ExecuteNonQuery();

                                comm.ExecuteNonQuery();
                                return new UserAcount(0);
                            }
                            return new UserAcount(userName, dataReader.GetInt16(0),1, dataReader["NickName"].Equals(DBNull.Value) ? null : dataReader["NickName"].ToString());
                        }
                        else
                        {
                            return new UserAcount(2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        //删除用户
        public bool DeleteAcount(string userName)
        {
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取所有用户
                        //只获取234的用户。应该根据权限 获取对应的用户
                        comm.CommandText = "DELETE FROM UserTable WHERE User=@User";
                        comm.Parameters.AddWithValue("@User", userName);
                        conn.Open();
                        int affectInt = comm.ExecuteNonQuery();

                        if (affectInt != 1)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        ////有权限的用户，添加其他用户
        //public bool AddAcount(string userName,string pwd)
        //{

        //}

        //有权限的用户，修改账户权限
        public bool ModefyUserPriv(string userName, int newLevel)
        {
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取所有用户
                        //只获取234的用户。应该根据权限 获取对应的用户
                        comm.CommandText = "UPDATE UserTable SET Level=@Level WHERE User=@User";
                        comm.Parameters.AddWithValue("@Level", newLevel);
                        comm.Parameters.AddWithValue("@User", userName);
                        conn.Open();
                        int affectInt = comm.ExecuteNonQuery();

                        if (affectInt != 1)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        //修改用户密码
        public bool ModefyUserPwd(string userName, string oldPwd, string newPwd)
        {
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取 加密密码
                        oldPwd = GetMd5_32bit(oldPwd);
                        newPwd = GetMd5_32bit(newPwd);
                        comm.CommandText = "UPDATE UserTable SET Pwd=@NewPwd WHERE User=@User AND Pwd=@OldPwd";
                        comm.Parameters.AddWithValue("@NewPwd", newPwd);
                        comm.Parameters.AddWithValue("@User", userName);
                        comm.Parameters.AddWithValue("@OldPwd", oldPwd);
                        conn.Open();
                        int affectInt = comm.ExecuteNonQuery();

                        if (affectInt != 1)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        //重载 直接修改
        public bool ModefyUserPwd(string userName, string newPwd)
        {
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取 加密密码
                        newPwd = GetMd5_32bit(newPwd);
                        comm.CommandText = "UPDATE UserTable SET Pwd=@NewPwd WHERE User=@User";
                        comm.Parameters.AddWithValue("@NewPwd", newPwd);
                        comm.Parameters.AddWithValue("@User", userName);
                        conn.Open();
                        //int affectInt = comm.ExecuteNonQuery();

                        if (comm.ExecuteNonQuery() != 1)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        public bool ModefyUserName(string userName, string newUserName)
        {
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取所有用户
                        //只获取234的用户。应该根据权限 获取对应的用户
                        comm.CommandText = "UPDATE UserTable SET User=@NewUser WHERE User=@OldUser";
                        comm.Parameters.AddWithValue("@NewUser", newUserName);
                        comm.Parameters.AddWithValue("@OldUser", userName);
                        conn.Open();
                        int affectInt = comm.ExecuteNonQuery();

                        if (affectInt != 1)
                        {
                            return false;
                        }

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        //查看已有用户信息，不同权限查看结果不同
        public void ViewAcount()
        {

        }

        //加载所有用户,应该根据权限不同 加载用户不同
        public Dictionary<string, string> LoadAllUser()
        {
            string connStr = string.Format("Data Source={0};Version=3;", @".\data\convertDB");
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    using (SQLiteCommand comm = new SQLiteCommand(conn))
                    {
                        //获取所有用户
                        //只获取234的用户。应该根据权限 获取对应的用户
                        comm.CommandText = "SELECT User,Level FROM UserTable WHERE Level=2 OR Level=3 OR Level=4";
                        conn.Open();
                        SQLiteDataReader dataReader = comm.ExecuteReader();

                        //返回字典，用户名和权限组成
                        Dictionary<string, string> userAndPrivDict = new Dictionary<string, string>();

                        while (dataReader.Read())//读取信息
                        {
                            userAndPrivDict.Add(dataReader["User"].ToString(), priviDict[dataReader.GetInt32(1)]);
                        }

                        return userAndPrivDict;
                    }
                }
            }
            catch (Exception ex)
            {

                return null;
            }
        }
    }
}

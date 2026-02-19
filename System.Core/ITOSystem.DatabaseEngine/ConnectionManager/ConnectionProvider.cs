using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Configuration;
using System.Security.Cryptography;

namespace ITOSystem.DatabaseEngine.ConnectionManager
{
    internal class ConnectionProvider
    {

        public string connectionStirng()
        {
            //AppSettingsReader appReader = new AppSettingsReader();
            string connStr = getConnectionStirng();
            return connStr;
        }
        public string connectionTranStirng()
        {
            //AppSettingsReader appReader = new AppSettingsReader();
            string connStr = getTranConnectionStirng();
            return connStr;
        }
        public string connectionOracleStirng()
        {
            //AppSettingsReader appReader = new AppSettingsReader();
            string connStr = getOracleConnectionStirng();
            return connStr;
        }
        private static string getConnectionStirng()
        {
            ConnectionProvider conProvider = new ConnectionProvider();
            string[] parseData = conProvider.ParseXMLFile().Split(',');
            string connStr = string.Empty;
            if (parseData[4].ToUpper() == "TRUE")
            {
                //connStr = "Data Source=" + dataSecurity.DecryptFile(parseData[0].ToString(), keyString) + ";Initial Catalog=" + dataSecurity.DecryptFile(parseData[1].ToString(), keyString) + ";Integrated Security=True";
                connStr = "Data Source=" + parseData[0].ToString() + ";Initial Catalog=" + parseData[1].ToString() + ";Integrated Security=True";

            }
            else
            {
                //connStr = "Data Source=" + dataSecurity.DecryptFile(parseData[0].ToString(), keyString) + ";Initial Catalog=" + dataSecurity.DecryptFile(parseData[1].ToString(), keyString) + ";user id=" + dataSecurity.DecryptFile(parseData[2].ToString(), keyString) + ";password=" + dataSecurity.DecryptFile(parseData[3].ToString(), keyString) + "";
                connStr = "Data Source=" + parseData[0].ToString() + ";Initial Catalog=" + parseData[1].ToString() + ";user id=" + parseData[2].ToString() + ";password=" + parseData[3].ToString() + "";
            }


            return connStr;
        }
        private static string getOracleConnectionStirng()
        {
            //EncryptionDecryption dataSecurity = new EncryptionDecryption();
            ConnectionProvider conProvider = new ConnectionProvider();
            string[] parseData = conProvider.ParseTranXMLFile().Split(',');
            string connStr = string.Empty;

            //connStr = "Data Source=" + dataSecurity.DecryptFile(parseData[0].ToString(), keyString) + ";Initial Catalog=" + dataSecurity.DecryptFile(parseData[1].ToString(), keyString) + ";user id=" + dataSecurity.DecryptFile(parseData[2].ToString(), keyString) + ";password=" + dataSecurity.DecryptFile(parseData[3].ToString(), keyString) + "";
            //connStr = "Data Source=" + parseData[0].ToString() + ";Initial Catalog=" + parseData[1].ToString() + ";user id=" + parseData[2].ToString() + ";password=" + parseData[3].ToString() + "";
            //connStr = "Data Source=192.168.10.18:1521/ORCL;User ID=KioTransUser;Password=kiotransadmin;Unicode=True";
            //connStr = "Data Source=SHARA;User ID=KioTransUser;Password=kiotransadmin;Unicode=True";
            //connStr = "Data Source=192.168.10.246:1521/KIONETWARE;User ID=KioTransUser;Password=kiotransadmin;Unicode=True";
            connStr = "Data Source=" + parseData[0].ToString() + ";user id=" + parseData[2].ToString() + ";password=" + parseData[3].ToString() + ";Unicode=True";
            //connStr = "Data Source=172.16.10.22:1521/ORCL;User ID=KioTransUser;Password=kiotransadmin;Unicode=True";
            //connStr = "Data Source=10.0.17.20:1521//RND;User ID=KIOTRANSUSER;Password=kiotransadmin;Unicode=True";
            return connStr;
        }
        private static string getTranConnectionStirng()
        {
            ConnectionProvider conProvider = new ConnectionProvider();
            string[] parseData = conProvider.ParseTranXMLFile().Split(',');
            string connStr = string.Empty;
            if (parseData[4].ToUpper() == "TRUE")
            {
                connStr = "Data Source=" + parseData[0].ToString() + ";Initial Catalog=" + parseData[1].ToString() + ";Integrated Security=True";

            }
            else
            {
                connStr = "Data Source=" + parseData[0].ToString() + ";Initial Catalog=" + parseData[1].ToString() + ";user id=" + parseData[2].ToString() + ";password=" + parseData[3].ToString() + "";
            }


            return connStr;
        }

        #region ***** Single DES *****
        static byte[] Key = ASCIIEncoding.ASCII.GetBytes("itclmech");
        static byte[] IV = ASCIIEncoding.ASCII.GetBytes("atronics");

        public string DES_Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(Key, IV), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }
        #endregion

        private string ParseXMLFile()
        {
            try
            {
                string dataSource = string.Empty;
                string dataBase = string.Empty;
                string userName = string.Empty;
                string passWord = string.Empty;
                string authenticationType = string.Empty;
                string isbn = String.Empty;

                XmlDocument doc = new XmlDocument();
                doc.Load(System.Windows.Forms.Application.StartupPath + "\\Application Configuration\\KioNetwareMonitor.xml");

                XmlNodeList bookList = doc.GetElementsByTagName("CONNECTION");

                foreach (XmlNode node in bookList)
                {
                    XmlElement bookElement = (XmlElement)node;

                    dataSource = DES_Decrypt(bookElement.GetElementsByTagName("DATASOURCE")[0].InnerText);
                    dataBase = DES_Decrypt(bookElement.GetElementsByTagName("DATABASE")[0].InnerText);
                    userName = DES_Decrypt(bookElement.GetElementsByTagName("USERNAME")[0].InnerText);
                    passWord = DES_Decrypt(bookElement.GetElementsByTagName("PASSWORD")[0].InnerText);
                    authenticationType = DES_Decrypt(bookElement.GetElementsByTagName("IsWindowsAuthentication")[0].InnerText);
                    if (bookElement.HasAttributes)
                    {
                        isbn = bookElement.Attributes["TYPE"].InnerText;
                        if (isbn == "SQL_SERVER")
                        {
                            break;
                        }

                    }
                }

                return dataSource.ToString() + "," + dataBase.ToString() + "," + userName.ToString() + "," + passWord.ToString() + "," + authenticationType.ToString();
            }
            catch (Exception errorException)
            {
                throw errorException;
            }
        }

       

        private string ParseTranXMLFile()
        {
            try
            {
                string dataSource = string.Empty;
                string dataBase = string.Empty;
                string userName = string.Empty;
                string passWord = string.Empty;
                string authenticationType = string.Empty;
                string isbn = String.Empty;

             //   this.configPath = AppDomain.CurrentDomain.BaseDirectory;
               // this.dsDB = new DataSet("dsDB");
               // this.dsDB.ReadXmlSchema(configPath + XMLFileName + ".xsd");
               // this.dsDB.ReadXml(configPath + XMLFileName + ".xml");

                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

                XmlDocument doc = new XmlDocument();
                doc.Load(System.Windows.Forms.Application.StartupPath  + "\\DBConfig.xml");

                XmlNodeList bookList = doc.GetElementsByTagName("CONNECTION");

                foreach (XmlNode node in bookList)
                {
                    XmlElement bookElement = (XmlElement)node;

                    dataSource = DES_Decrypt(bookElement.GetElementsByTagName("DATASOURCE")[0].InnerText);
                    dataBase = DES_Decrypt(bookElement.GetElementsByTagName("DATABASE")[0].InnerText);
                    userName = DES_Decrypt(bookElement.GetElementsByTagName("USERNAME")[0].InnerText);
                    passWord = DES_Decrypt(bookElement.GetElementsByTagName("PASSWORD")[0].InnerText);
                    authenticationType = DES_Decrypt(bookElement.GetElementsByTagName("IsWindowsAuthentication")[0].InnerText);
                    if (bookElement.HasAttributes)
                    {
                        isbn = bookElement.Attributes["TYPE"].InnerText;
                        if (isbn == "SQL_SERVER")
                        {
                            break;
                        }

                    }
                }

                return dataSource.ToString() + "," + dataBase.ToString() + "," + userName.ToString() + "," + passWord.ToString() + "," + authenticationType.ToString();
            }
            catch (Exception errorException)
            {
                throw errorException;
            }
        }
    }
    internal class EncryptionDecryption
    {
        public static string DecryptAndWriteToFile(string txtLine, string key)
        {

            string strPwd;
            string[] strText;
            //string line=textLine.ToString();
            string msg = string.Empty;

            EncryptionDecryption encryptDecrypt = new EncryptionDecryption();

            key = key.ToUpper();

            strText = txtLine.Split('©');
            strPwd = key.Substring(8, 8).ToString();
            msg = encryptDecrypt.DecryptFile(strText[0].ToString(), strPwd);

            strPwd = key.Substring(0, 8).ToString();
            msg = encryptDecrypt.DecryptFile(msg, strPwd);

            return msg;
        }

        public byte[] ReadFile(string filePath)
        {
            byte[] buffer;

            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            try
            {
                int length = (int)fileStream.Length;  // get file length

                buffer = new byte[length];            // create buffer
                int count;                            // actual number of bytes read
                int sum = 0;                          // total number of bytes read

                // read until Read method returns 0 (end of the stream has been reached)
                while ((count = fileStream.Read(buffer, sum, length - sum)) > 0)
                    sum += count;  // sum is a buffer offset for next reading
            }
            finally
            {
                fileStream.Close();
            }
            return buffer;
        }

        public string DecryptFile(string strText, string key)
        {
            int i;
            int c;
            string strBuff;
            string strPwd;

            strBuff = "";
            strPwd = key;
            key = key.ToUpper();
            Conversion conversion = new Conversion();

            if (strPwd.Length > 0)
            {
                for (i = 0; i < strText.Length; i++)
                {

                    char chr = Convert.ToChar(strText.Substring(i, 1));
                    c = Convert.ToInt32(chr);
                    if (c < 48 || c > 255)
                    {
                        c = conversion.CharToCode(chr);
                    }

                    int kk = (i) % strPwd.Length;
                    kk = kk + 1;
                    if (kk == 8)
                    {
                        kk = 0;
                    }
                    string ss = strPwd.Substring(kk, 1);  //, 1
                    char ch = Convert.ToChar(ss);
                    c = c - Convert.ToInt32(ch);
                    if (c < 0)
                    {
                        c = 256 + c;
                    }
                    if (c > 127)
                    {
                        strBuff = strBuff + conversion.CodeToChar(c);
                    }
                    else
                    {
                        strBuff = strBuff + Convert.ToChar(c);
                    }
                }
            }
            else
            {
                strBuff = strText;
            }

            return strBuff;
            //MessageBox.Show(strPwd.Length.ToString());
        }

        public static string Encryption(string password, string key)
        {
            int i;
            int c;
            string strBuff = "";
            string strPwd = string.Empty;

            //string key = "ITCLimited";
            key = key.ToUpper();

            Conversion conversion = new Conversion();

            #region 1st Part

            strPwd = key.Substring(0, 5).ToString();

            if (strPwd.Length > 0)
            {
                for (i = 0; i < password.Length; i++)
                {

                    char chr = Convert.ToChar(password.Substring(i, 1));
                    c = Convert.ToInt32(chr);
                    if (c < 48 || c > 255)
                    {
                        c = conversion.CharToCode(chr);
                    }
                    //c = c + Convert.ToInt16(Convert.ToChar(strPwd.Substring((i % strPwd.Length) + 1, 1)));
                    int kk = (i) % strPwd.Length;
                    kk = kk + 1;
                    if (kk == 5)
                    {
                        kk = 0;
                    }
                    string ss = strPwd.Substring(kk, 1);  //, 1
                    char ch = Convert.ToChar(ss);
                    c = c + Convert.ToInt32(ch);
                    if (c < 0)
                    {
                        c = 256 + c;
                    }
                    if (c > 127)
                    {
                        strBuff = strBuff + conversion.CodeToChar(c);
                    }
                    else
                    {
                        strBuff = strBuff + Convert.ToChar(c);
                    }
                }
            }
            else
            {
                strBuff = password;
            }

            #endregion

            password = strBuff;
            strBuff = string.Empty;

            #region 2nd Part

            strPwd = key.Substring(5, 5).ToString();

            if (strPwd.Length > 0)
            {
                for (i = 0; i < password.Length; i++)
                {

                    char chr = Convert.ToChar(password.Substring(i, 1));

                    c = Convert.ToInt32(chr);
                    if (c < 48 || c > 255)
                    {
                        c = conversion.CharToCode(chr);
                    }
                    //c = c + Convert.ToInt16(Convert.ToChar(strPwd.Substring((i % strPwd.Length) + 1, 1)));
                    int kk = (i) % strPwd.Length;
                    kk = kk + 1;
                    if (kk == 5)
                    {
                        kk = 0;
                    }
                    string ss = strPwd.Substring(kk, 1);  //, 1
                    char ch = Convert.ToChar(ss);
                    c = c + Convert.ToInt32(ch);
                    if (c < 0)
                    {
                        c = 256 + c;
                    }
                    if (c > 127)
                    {
                        strBuff = strBuff + conversion.CodeToChar(c);
                    }
                    else
                    {
                        strBuff = strBuff + Convert.ToChar(c);
                    }
                }
            }
            else
            {
                strBuff = password;
            }

            #endregion

            return strBuff;
            //MessageBox.Show(strPwd.Length.ToString());
        }
    }
    internal class Conversion
    {
        public static string BinaryToOctate(string str)
        {
            if (str == "000")
            {
                return "0";
            }
            else if (str == "001")
            {
                return "1";
            }
            else if (str == "010")
            {
                return "2";
            }
            else if (str == "011")
            {
                return "3";
            }
            else if (str == "100")
            {
                return "4";
            }
            else if (str == "101")
            {
                return "5";
            }
            else if (str == "110")
            {
                return "6";
            }
            else if (str == "111")
            {
                return "7";
            }
            else
            {
                return "8";
            }
        }

        public static string OctateToBinary(string str)
        {
            if (str == "1")
            {
                return "000";
            }
            else if (str == "1")
            {
                return "001";
            }
            else if (str == "2")
            {
                return "010";
            }
            else if (str == "3")
            {
                return "011";
            }
            else if (str == "4")
            {
                return "100";
            }
            else if (str == "5")
            {
                return "101";
            }
            else if (str == "6")
            {
                return "110";
            }
            else if (str == "7")
            {
                return "111";
            }
            else
            {
                return "000";
            }
        }

        public char CodeToChar(int charCode)
        {
            int code = charCode;
            char character = '#';
            switch (code)
            {
                case 128:
                    character = '€';
                    break;
                case 129:
                    character = '';
                    //return character;
                    break;
                case 130:
                    character = '‚';
                    //return character;
                    break;
                case 131:
                    character = 'ƒ';
                    //return character;
                    break;
                case 132:
                    character = '„';
                    //return character;
                    break;
                case 133:
                    character = '…';
                    //return character;
                    break;
                case 134:
                    character = '†';
                    //return character;
                    break;
                case 135:
                    character = '‡';
                    //return character;
                    break;
                case 136:
                    character = 'ˆ';
                    //return character;
                    break;
                case 137:
                    character = '‰';
                    //return character;
                    break;
                case 138:
                    character = 'Š';
                    //return character;
                    break;
                case 139:
                    character = '‹';
                    //return character;
                    break;
                case 140:
                    character = 'Œ';
                    //return character;
                    break;
                case 141:
                    character = '';
                    //return character;
                    break;
                case 142:
                    character = 'Ž';
                    //return character;
                    break;
                case 143:
                    character = '';
                    //return character;
                    break;
                case 144:
                    character = '';
                    //return character;
                    break;
                case 145:
                    character = '‘';
                    //return character;
                    break;
                case 146:
                    character = '’';
                    //return character;
                    break;
                case 147:
                    character = '“';
                    //return character;
                    break;
                case 148:
                    character = '”';
                    //return character;
                    break;
                case 149:
                    character = '•';
                    //return character;
                    break;
                case 150:
                    character = '–';
                    //return character;
                    break;
                case 151:
                    character = '—';
                    //return character;
                    break;
                case 152:
                    character = '˜';
                    //return character;
                    break;
                case 153:
                    character = '™';
                    //return character;
                    break;
                case 154:
                    character = 'š';
                    //return character;
                    break;
                case 155:
                    character = '›';
                    //return character;
                    break;
                case 156:
                    character = 'œ';
                    //return character;
                    break;
                case 157:
                    character = '';
                    //return character;
                    break;
                case 158:
                    character = 'ž';
                    //return character;
                    break;
                case 159:
                    character = 'Ÿ';
                    //return character;
                    break;
                case 160:
                    character = ' ';
                    //return character;
                    break;
                case 161:
                    character = '¡';
                    //return character;
                    break;
                case 162:
                    character = '¢';
                    //return character;
                    break;
                case 163:
                    character = '£';
                    //return character;
                    break;
                case 164:
                    character = '¤';
                    //return character;
                    break;
                case 165:
                    character = '¥';
                    //return character;
                    break;
                case 166:
                    character = '¦';
                    //return character;
                    break;
                case 167:
                    character = '§';
                    //return character;
                    break;
                case 168:
                    character = '¨';
                    //return character;
                    break;
                case 169:
                    character = '©';
                    //return character;
                    break;
                case 170:
                    character = 'ª';
                    //return character;
                    break;
                case 171:
                    character = '«';
                    //return character;
                    break;
                case 172:
                    character = '¬';
                    //return character;
                    break;
                case 173:
                    character = '­';
                    //return character;
                    break;
                case 174:
                    character = '®';
                    //return character;
                    break;
                case 175:
                    character = '¯';
                    //return character;
                    break;
                case 176:
                    character = '°';
                    //return character;
                    break;
                case 177:
                    character = '±';
                    //return character;
                    break;
                case 178:
                    character = '²';
                    //return character;
                    break;
                case 179:
                    character = '³';
                    //return character;
                    break;
                case 180:
                    character = '´';
                    //return character;
                    break;
                case 181:
                    character = 'µ';
                    //return character;
                    break;
                case 182:
                    character = '¶';
                    //return character;
                    break;
                case 183:
                    character = '·';
                    //return character;
                    break;
                case 184:
                    character = '¸';
                    //return character;
                    break;
                case 185:
                    character = '¹';
                    //return character;
                    break;
                case 186:
                    character = 'º';
                    //return character;
                    break;
                case 187:
                    character = '»';
                    //return character;
                    break;
                case 188:
                    character = '¼';
                    //return character;
                    break;
                case 189:
                    character = '½';
                    //return character;
                    break;
                case 190:
                    character = '¾';
                    //return character;
                    break;
                case 191:
                    character = '¿';
                    //return character;
                    break;
                case 192:
                    character = 'À';
                    //return character;
                    break;
                case 193:
                    character = 'Á';
                    //return character;
                    break;
                case 194:
                    character = 'Â';
                    //return character;
                    break;
                case 195:
                    character = 'Ã';
                    //return character;
                    break;
                case 196:
                    character = 'Ä';
                    //return character;
                    break;
                case 197:
                    character = 'Å';
                    //return character;
                    break;
                case 198:
                    character = 'Æ';
                    //return character;
                    break;
                case 199:
                    character = 'Ç';
                    //return character;
                    break;
                case 200:
                    character = 'È';
                    //return character;
                    break;
                case 201:
                    character = 'É';
                    //return character;
                    break;
                case 202:
                    character = 'Ê';
                    //return character;
                    break;
                case 203:
                    character = 'Ë';
                    //return character;
                    break;
                case 204:
                    character = 'Ì';
                    //return character;
                    break;
                case 205:
                    character = 'Í';
                    //return character;
                    break;
                case 206:
                    character = 'Î';
                    //return character;
                    break;
                case 207:
                    character = 'Ï';
                    //return character;
                    break;
                case 208:
                    character = 'Ð';
                    //return character;
                    break;
                case 209:
                    character = 'Ñ';
                    //return character;
                    break;
                case 210:
                    character = 'Ò';
                    //return character;
                    break;
                case 211:
                    character = 'Ó';
                    //return character;
                    break;
                case 212:
                    character = 'Ô';
                    //return character;
                    break;
                case 213:
                    character = 'Õ';
                    //return character;
                    break;
                case 214:
                    character = 'Ö';
                    //return character;
                    break;
                case 215:
                    character = '×';
                    //return character;
                    break;
                case 216:
                    character = 'Ø';
                    //return character;
                    break;
                case 217:
                    character = 'Ù';
                    //return character;
                    break;
                case 218:
                    character = 'Ú';
                    //return character;
                    break;
                case 219:
                    character = 'Û';
                    //return character;
                    break;
                case 220:
                    character = 'Ü';
                    //return character;
                    break;
                case 221:
                    character = 'Ý';
                    //return character;
                    break;
                case 222:
                    character = 'Þ';
                    //return character;
                    break;
                case 223:
                    character = 'ß';
                    //return character;
                    break;
                case 224:
                    character = 'à';
                    //return character;
                    break;
                case 225:
                    character = 'á';
                    //return character;
                    break;
                case 226:
                    character = 'â';
                    //return character;
                    break;
                case 227:
                    character = 'ã';
                    //return character;
                    break;
                case 228:
                    character = 'ä';
                    //return character;
                    break;
                case 229:
                    character = 'å';
                    //return character;
                    break;
                case 230:
                    character = 'æ';
                    //return character;
                    break;
                case 231:
                    character = 'ç';
                    //return character;
                    break;
                case 232:
                    character = 'è';
                    //return character;
                    break;
                case 233:
                    character = 'é';
                    //return character;
                    break;
                case 234:
                    character = 'ê';
                    //return character;
                    break;
                case 235:
                    character = 'ë';
                    //return character;
                    break;
                case 236:
                    character = 'ì';
                    //return character;
                    break;
                case 237:
                    character = 'í';
                    //return character;
                    break;
                case 238:
                    character = 'î';
                    //return character;
                    break;
                case 239:
                    character = 'ï';
                    //return character;
                    break;
                case 240:
                    character = 'ð';
                    //return character;
                    break;
                case 241:
                    character = 'ñ';
                    //return character;
                    break;
                case 242:
                    character = 'ò';
                    //return character;
                    break;
                case 243:
                    character = 'ó';
                    //return character;
                    break;
                case 244:
                    character = 'ô';
                    //return character;
                    break;
                case 245:
                    character = 'õ';
                    //return character;
                    break;
                case 246:
                    character = 'ö';
                    //return character;
                    break;
                case 247:
                    character = '÷';
                    //return character;
                    break;
                case 248:
                    character = 'ø';
                    //return character;
                    break;
                case 249:
                    character = 'ù';
                    //return character;
                    break;
                case 250:
                    character = 'ú';
                    //return character;
                    break;
                case 251:
                    character = 'û';
                    //return character;
                    break;
                case 252:
                    character = 'ü';
                    //return character;
                    break;
                case 253:
                    character = 'ý';
                    //return character;
                    break;
                case 254:
                    character = 'þ';
                    //return character;
                    break;
                case 255:
                    character = 'ÿ';
                    //return character;
                    break;
            }

            return character;
        }

        public int CharToCode(char chr)
        {
            char character = chr;
            int code = 48;
            switch (character)
            {
                case '€':
                    code = 128;
                    break;
                case '':
                    code = 129;
                    break;
                case '‚':
                    code = 130;
                    //return character;
                    break;
                case 'ƒ':
                    code = 131;
                    //return character;
                    break;
                case '„':
                    code = 132;
                    //return character;
                    break;
                case '…':
                    code = 133;
                    //return character;
                    break;
                case '†':
                    code = 134;
                    //return character;
                    break;
                case '‡':
                    code = 135;
                    //return character;
                    break;
                case 'ˆ':
                    code = 136;
                    //return character;
                    break;
                case '‰':
                    code = 137;
                    //return character;
                    break;
                case 'Š':
                    code = 138;
                    //return character;
                    break;
                case '‹':
                    code = 139;
                    //return character;
                    break;
                case 'Œ':
                    code = 140;
                    //return character;
                    break;
                case '':
                    code = 141;
                    //return character;
                    break;
                case 'Ž':
                    code = 142;
                    //return character;
                    break;
                case '':
                    code = 143;
                    //return character;
                    break;
                case '':
                    code = 144;
                    //return character;
                    break;
                case '‘':
                    code = 145;
                    //return character;
                    break;
                case '’':
                    code = 146;
                    //return character;
                    break;
                case '“':
                    code = 147;
                    //return character;
                    break;
                case '”':
                    code = 148;
                    //return character;
                    break;
                case '•':
                    code = 149;
                    //return character;
                    break;
                case '–':
                    code = 150;
                    //return character;
                    break;
                case '—':
                    code = 151;
                    //return character;
                    break;
                case '˜':
                    code = 152;
                    //return character;
                    break;
                case '™':
                    code = 153;
                    //return character;
                    break;
                case 'š':
                    code = 154;
                    //return character;
                    break;
                case '›':
                    code = 155;
                    //return character;
                    break;
                case 'œ':
                    code = 156;
                    //return character;
                    break;
                case '':
                    code = 157;
                    //return character;
                    break;
                case 'ž':
                    code = 158;
                    //return character;
                    break;
                case 'Ÿ':
                    code = 159;
                    //return character;
                    break;
                case ' ':
                    code = 160;
                    //return character;
                    break;
                case '¡':
                    code = 161;
                    //return character;
                    break;
                case '¢':
                    code = 162;
                    //return character;
                    break;
                case '£':
                    code = 163;
                    //return character;
                    break;
                case '¤':
                    code = 164;
                    //return character;
                    break;
                case '¥':
                    code = 165;
                    //return character;
                    break;
                case '¦':
                    code = 166;
                    //return character;
                    break;
                case '§':
                    code = 167;
                    //return character;
                    break;
                case '¨':
                    code = 168;
                    //return character;
                    break;
                case '©':
                    code = 169;
                    //return character;
                    break;
                case 'ª':
                    code = 170;
                    //return character;
                    break;
                case '«':
                    code = 171;
                    //return character;
                    break;
                case '¬':
                    code = 172;
                    //return character;
                    break;
                case '­':
                    code = 173;
                    //return character;
                    break;
                case '®':
                    code = 174;
                    //return character;
                    break;
                case '¯':
                    code = 175;
                    //return character;
                    break;
                case '°':
                    code = 176;
                    //return character;
                    break;
                case '±':
                    code = 177;
                    //return character;
                    break;
                case '²':
                    code = 178;
                    //return character;
                    break;
                case '³':
                    code = 179;
                    //return character;
                    break;
                case '´':
                    code = 180;
                    //return character;
                    break;
                case 'µ':
                    code = 181;
                    //return character;
                    break;
                case '¶':
                    code = 182;
                    //return character;
                    break;
                case '·':
                    code = 183;
                    //return character;
                    break;
                case '¸':
                    code = 184;
                    //return character;
                    break;
                case '¹':
                    code = 185;
                    //return character;
                    break;
                case 'º':
                    code = 186;
                    //return character;
                    break;
                case '»':
                    code = 187;
                    //return character;
                    break;
                case '¼':
                    code = 188;
                    //return character;
                    break;
                case '½':
                    code = 189;
                    //return character;
                    break;
                case '¾':
                    code = 190;
                    //return character;
                    break;
                case '¿':
                    code = 191;
                    //return character;
                    break;
                case 'À':
                    code = 192;
                    //return character;
                    break;
                case 'Á':
                    code = 193;
                    //return character;
                    break;
                case 'Â':
                    code = 194;
                    //return character;
                    break;
                case 'Ã':
                    code = 195;
                    //return character;
                    break;
                case 'Ä':
                    code = 196;
                    //return character;
                    break;
                case 'Å':
                    code = 197;
                    //return character;
                    break;
                case 'Æ':
                    code = 198;
                    //return character;
                    break;
                case 'Ç':
                    code = 199;
                    //return character;
                    break;
                case 'È':
                    code = 200;
                    //return character;
                    break;
                case 'É':
                    code = 201;
                    //return character;
                    break;
                case 'Ê':
                    code = 202;
                    //return character;
                    break;
                case 'Ë':
                    code = 203;
                    //return character;
                    break;
                case 'Ì':
                    code = 204;
                    //return character;
                    break;
                case 'Í':
                    code = 205;
                    //return character;
                    break;
                case 'Î':
                    code = 206;
                    //return character;
                    break;
                case 'Ï':
                    code = 207;
                    //return character;
                    break;
                case 'Ð':
                    code = 208;
                    //return character;
                    break;
                case 'Ñ':
                    code = 209;
                    //return character;
                    break;
                case 'Ò':
                    code = 210;
                    //return character;
                    break;
                case 'Ó':
                    code = 211;
                    //return character;
                    break;
                case 'Ô':
                    code = 212;
                    //return character;
                    break;
                case 'Õ':
                    code = 213;
                    //return character;
                    break;
                case 'Ö':
                    code = 214;
                    //return character;
                    break;
                case '×':
                    code = 215;
                    //return character;
                    break;
                case 'Ø':
                    code = 216;
                    //return character;
                    break;
                case 'Ù':
                    code = 217;
                    //return character;
                    break;
                case 'Ú':
                    code = 218;
                    //return character;
                    break;
                case 'Û':
                    code = 219;
                    //return character;
                    break;
                case 'Ü':
                    code = 220;
                    //return character;
                    break;
                case 'Ý':
                    code = 221;
                    //return character;
                    break;
                case 'Þ':
                    code = 222;
                    //return character;
                    break;
                case 'ß':
                    code = 223;
                    //return character;
                    break;
                case 'à':
                    code = 224;
                    //return character;
                    break;
                case 'á':
                    code = 225;
                    //return character;
                    break;
                case 'â':
                    code = 226;
                    //return character;
                    break;
                case 'ã':
                    code = 227;
                    //return character;
                    break;
                case 'ä':
                    code = 228;
                    //return character;
                    break;
                case 'å':
                    code = 229;
                    //return character;
                    break;
                case 'æ':
                    code = 230;
                    //return character;
                    break;
                case 'ç':
                    code = 231;
                    //return character;
                    break;
                case 'è':
                    code = 232;
                    //return character;
                    break;
                case 'é':
                    code = 233;
                    //return character;
                    break;
                case 'ê':
                    code = 234;
                    //return character;
                    break;
                case 'ë':
                    code = 235;
                    //return character;
                    break;
                case 'ì':
                    code = 236;
                    //return character;
                    break;
                case 'í':
                    code = 237;
                    //return character;
                    break;
                case 'î':
                    code = 238;
                    //return character;
                    break;
                case 'ï':
                    code = 239;
                    //return character;
                    break;
                case 'ð':
                    code = 240;
                    //return character;
                    break;
                case 'ñ':
                    code = 241;
                    //return character;
                    break;
                case 'ò':
                    code = 242;
                    //return character;
                    break;
                case 'ó':
                    code = 243;
                    //return character;
                    break;
                case 'ô':
                    code = 244;
                    //return character;
                    break;
                case 'õ':
                    code = 245;
                    //return character;
                    break;
                case 'ö':
                    code = 246;
                    //return character;
                    break;
                case '÷':
                    code = 247;
                    //return character;
                    break;
                case 'ø':
                    code = 248;
                    //return character;
                    break;
                case 'ù':
                    code = 249;
                    //return character;
                    break;
                case 'ú':
                    code = 250;
                    //return character;
                    break;
                case 'û':
                    code = 251;
                    //return character;
                    break;
                case 'ü':
                    code = 252;
                    //return character;
                    break;
                case 'ý':
                    code = 253;
                    //return character;
                    break;
                case 'þ':
                    code = 254;
                    //return character;
                    break;
                case 'ÿ':
                    code = 255;
                    //return character;
                    break;
            }

            return code;
        }


    }
}

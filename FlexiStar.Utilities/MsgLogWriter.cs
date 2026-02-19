using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Microsoft.VisualBasic;
using System.Threading;
//
namespace FlexiStar.Utilities
{
    public class MsgLogWriter
    {
        public void logTrace(string strFilePath,string strFileName, string Msg)
        {
           FileInfo objFile = new FileInfo(strFileName);

            try
            {
                FileStream fs = null;
                StreamWriter sr = null;

                if (!Directory.Exists(strFilePath))
                    Directory.CreateDirectory(strFilePath);

                if (!Directory.Exists(strFilePath + "\\BackUp"))
                    Directory.CreateDirectory(strFilePath + "\\BackUp");

                if (!File.Exists(strFilePath + "\\" + strFileName))
                    fs = File.Create(strFilePath + "\\" + strFileName);
                else
                {
                    FileInfo _file = new FileInfo(strFilePath + "\\" + strFileName);
                    long _filesize = _file.Length;
                    if (System.DateTime.Now.Hour == 0)
                    {
                        if (!File.Exists(strFilePath + "\\BackUp\\" + System.DateTime.Now.ToString("dd.MM.yyyy") + "_" + strFileName))
                        {
                            try
                            {
                                File.Move(strFilePath + "\\" + strFileName, strFilePath + "\\BackUp\\" + System.DateTime.Now.ToString("dd.MM.yyyy") + "_" + strFileName);
                                //Directory.Move(strFilePath + System.DateTime.Now.Date.ToShortDateString() + "_" + strFileName, strFilePath + "\\BackUp\\" + System.DateTime.Now.Date.ToShortDateString() + "_" + strFileName);
                            }
                            catch (Exception io) 
                            {
 
                            }
                            finally
                            {
                                if (!File.Exists(strFilePath + "\\" + strFileName))
                                    fs = File.Create(strFilePath + "\\" + strFileName);
                            }
                        }
                    }
                    else if (_filesize > 10000000)
                    {
                        if (!File.Exists(strFilePath + "\\BackUp\\" + System.DateTime.Now.ToString("dd.MM.yyyy") + "_" + strFileName))
                        {
                            try
                            {
                                File.Move(strFilePath + "\\" + strFileName, strFilePath + "\\BackUp\\" + System.DateTime.Now.ToString("dd.MM.yyyy") + "_" + strFileName);
                            }
                            catch (Exception io)
                            {

                            }
                            finally
                            {
                                if (!File.Exists(strFilePath + "\\" + strFileName))
                                    fs = File.Create(strFilePath + "\\" + strFileName);
                            }
                        }
                    }
                    else
                    {
                        sr = File.AppendText(strFilePath + "\\" + strFileName);
                        sr.WriteLine(Msg);
                        sr.Close();
                    }
                }
                if (fs != null)
                {
                    if (File.Exists(strFilePath + "\\" + strFileName))
                    {
                        sr = new StreamWriter(fs, System.Text.Encoding.ASCII);
                        sr.WriteLine(Msg);
                        sr.Close();
                        fs.Close();
                    }
                } 
            }
            catch (Exception ex)
            {
                //ExceptionLogWriter(ex.Message + "\nAt FlexiStar.Utilities.MsgLogWriter.logTrace()", objConUtl.FilePath, objConUtl.ExceptionFileName);
                //Thread.ResetAbort();
            }
        }


        public void LogWriter(string message, messageType enmsgType)
        {
            ConfigurationUtility objConUtl = ConfigurationUtility.CreateInstance();
            try
            {
                messageType _enmsgType = enmsgType;

                switch (_enmsgType)
                {
                    case messageType.Port:
                        {
                            logTrace(objConUtl.FilePath, objConUtl.PortFileName, System.DateTime.Now.ToString() + " : " + message);
                            break;
                        }
                    case messageType.Trace:
                        {
                            logTrace(objConUtl.FilePath, objConUtl.TraceFileName, System.DateTime.Now.ToString() + " : " + message);
                            break;
                        }
                    case messageType.Execption:
                        {
                            logTrace(objConUtl.FilePath, objConUtl.ExceptionFileName, System.DateTime.Now.ToString() + " : " + message);
                            break;
                        }
                }
            }
            catch (Exception ex)
            {

            }

        }

        
    }

    public enum messageType : int
    {
        Port = 1,
        Trace = 2,
        Execption = 3
    }
}

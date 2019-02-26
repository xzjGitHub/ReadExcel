using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class CodeGenerator
    {
        StringBuilder m_str = new StringBuilder();
        string _Indend;

        public void Print(params object[] values)
        {
            foreach(object obj in values)
            {
                m_str.Append(obj.ToString());
            }
        }

        public void BeginLine()
        {
            m_str.Append(_Indend);
        }

        public void EndLine()
        {
            m_str.Append("\n");
        }

        /// <summary>
        /// 把多个参数打到一行里
        /// </summary>
        /// <param name="values"></param>
        public void PrintLine(params object[] values)
        {
            BeginLine();
            Print(values);
            EndLine();
        }

        /// <summary>
        /// 进入一个层次, 例如括号
        /// </summary>
        public void In()
        {
            _Indend += "\t";
        }

        /// <summary>
        /// 退出一个层次
        /// </summary>
        public void Out()
        {
            if(_Indend.Length > 0)
            {
                _Indend = _Indend.Substring(1);
            }
        }

        public void WriteFile(string path)
        {
            if(File.Exists(path))
                File.Delete(path);

            FileStream fs = new FileStream(path,FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            sw.Write(m_str.ToString());
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }

        public void WriteFile(string path,string info)
        {
            if(File.Exists(path))
                File.Delete(path);

            FileStream fs = new FileStream(path,FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            sw.Write(info);
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }

        public static void WriteLog(string path,string des)
        {
            if(File.Exists(path))
                File.Delete(path);

            FileStream fs = new FileStream(path,FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            sw.Write(des.ToString());
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }
    }
}

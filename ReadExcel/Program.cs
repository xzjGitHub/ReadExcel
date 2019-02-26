using NPOI.HSSF.UserModel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using LitJson;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;

namespace ReadExcel
{
 

    class Program
    {
        private static string excelPath = @"C:\Users\Yoki\Desktop\aa";
        private static string configPath = @"C:\Users\Yoki\Desktop\bb\";
        private static string codePath = "";
        private static string logPath = @"C:\Users\Yoki\Desktop\AllExcel";

        private static StringBuilder Log;

        static void Main(string[] args)
        {
            Log = new StringBuilder();
            Console.WriteLine("开始导表");
            Log.Append("开始导表\n");
            Console.WriteLine();
            if(args.Length > 0)
            {
                excelPath = args[0].Replace("/","\\");
                codePath = args[1].Replace("/","\\");
                configPath = args[2].Replace("/","\\");
                logPath = args[3].Replace("/","\\");
            }
            else
            {
                configPath = @"C:\Users\Yoki\Desktop\TestReadExcel\Assets\Resources\Config\";
                codePath = @"C:\Users\Yoki\Desktop\TestReadExcel\Assets\Scripts\";
            }

            string[] files = Directory.GetFiles(excelPath,"*.xlsx");
            for(int i = 0; i < files.Length; i++)
            {
                Log.Append(files[i] + " 开始导表\n");
                Console.WriteLine(files[i] + " 开始导表");
                XSSFWorkbook hssfworkbook = LoadExcel(files[i]);
                XSSFSheet sheet = (XSSFSheet)hssfworkbook.GetSheetAt(0);
                GenertorDao(sheet);
                GenertorLogic(sheet);
                GenertorJson(sheet);
                Console.WriteLine(files[i] + " 导表完成");
                Log.Append(files[i] + " 导表完成\n\n");
                Console.WriteLine();
            }
            Console.WriteLine();
            Console.WriteLine("所有配置表导表完成");
            Log.Append("所有配置表导表完成");
            // CodeGenerator.WriteLog(logPath + "\\log.txt",Log.ToString());
            Console.Read();
        }


        /// <summary>
        /// 加载excel
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        static XSSFWorkbook LoadExcel(string path)
        {
            XSSFWorkbook XSSFworkbook = new XSSFWorkbook();
            try
            {
                using(FileStream file = new FileStream(path,FileMode.Open,FileAccess.Read))
                {
                    XSSFworkbook = new XSSFWorkbook(file);
                }
            }
            catch(Exception e)
            {
                throw e;
            }
            return XSSFworkbook;
        }
 
        /// <summary>
        /// 生成对应的cs文件
        /// </summary>
        /// <param name="sheet"></param>
        static void GenertorDao(XSSFSheet sheet)
        {
            CodeGenerator code = new CodeGenerator();

            code.PrintLine("//***************************************************************");
            code.PrintLine("//类名：",sheet.SheetName,"Ex");
            code.PrintLine("//作者：",System.Environment.MachineName);
            code.PrintLine("//日期：",DateTime.Now.ToString());
            code.PrintLine("//作用：",sheet.SheetName,"的数据类");
            code.PrintLine("//注意:","不要在此类里面写代码!!!");
            code.PrintLine("//***************************************************************");

            code.PrintLine();
            code.PrintLine("using System;");
            code.PrintLine("using System.Collections.Generic;");
            code.PrintLine();
            code.PrintLine("public class ",sheet.SheetName,"{");
            code.In();

            #region  生成变量
            XSSFRow typeRow = (XSSFRow)sheet.GetRow(0);
            XSSFRow desRow = (XSSFRow)sheet.GetRow(1);
            XSSFRow nameRow = (XSSFRow)sheet.GetRow(2);
            for(int i = 0; i < typeRow.LastCellNum; i++)
            {
                string type = typeRow.GetCell(i).ToString();
                string des = string.Empty;
                if (desRow.GetCell(i) != null)
                {
                    des = desRow.GetCell(i).ToString();
                }
 
                string name = nameRow.GetCell(i).ToString();

                code.PrintLine("/// <summary>");
                code.PrintLine("///" + des);
                code.PrintLine("/// </summary>");

                if(!type.Contains("List") && !type.Contains("Dictionary"))
                {
                    code.PrintLine("public ",type," ",name,";");
                }
                else
                {
                    code.PrintLine("public ",type," ",name,"=new ",type,"();");
                }
            }
            #endregion


            code.Out();
            code.PrintLine("}");
            code.WriteFile(codePath + sheet.SheetName + ".cs");

            Console.WriteLine(sheet.SheetName + ".cs 代码生成完成");
            Log.Append(sheet.SheetName + ".cs 代码生成完成\n");
        }

        /// <summary>
        /// 生成cs管理文件
        /// </summary>
        /// <param name="sheet"></param>
        static void GenertorLogic(XSSFSheet sheet)
        {
            CodeGenerator code = new CodeGenerator();
            DateTime dt = DateTime.Now;
            code.PrintLine("//***************************************************************");
            code.PrintLine("//类名：",sheet.SheetName,"Ex");
            code.PrintLine("//作者：",System.Environment.MachineName);
            code.PrintLine("//日期：",DateTime.Now.ToString());
            code.PrintLine("//作用：",sheet.SheetName,"的工具类");
            code.PrintLine("//注意:","不要在此类里面写代码!!!,如有需要，请添加一个此类的部分类");
            code.PrintLine("//***************************************************************");
            code.PrintLine();
            code.PrintLine("using System;");
            code.PrintLine("using System.Collections.Generic;");
            code.PrintLine("using UnityEngine;");
            code.PrintLine("using Newtonsoft.Json;");
            code.PrintLine();
            code.PrintLine("public partial class ",sheet.SheetName,"Ex");
            code.PrintLine("{");
            code.In();
            code.PrintLine();

            //字典
            code.PrintLine("/// <summary>");
            code.PrintLine("///"," 所有的配置表数据信息");
            code.PrintLine("/// </summary>");
            code.PrintLine("private static List<",sheet.SheetName,"> m_all",sheet.SheetName,"s"," = new List<",sheet.SheetName,">();");
            //属性
            code.PrintLine("public static List<",sheet.SheetName,"> ","All",sheet.SheetName,"s");
            code.PrintLine("{");
            code.In();
            code.PrintLine("get");
            code.PrintLine("{");
            code.In();
            code.PrintLine("if(m_all",sheet.SheetName,"s",".Count == 0)");
            code.In();
            code.PrintLine("Init();");
            code.Out();
            code.PrintLine("return"," m_all",sheet.SheetName,"s",";");
            code.Out();
            code.PrintLine("}");
            code.Out();
            code.PrintLine("}");

            //get方法
            code.PrintLine();
            code.PrintLine("/// <summary>");
            code.PrintLine("///"," 更据配置表id获取对象");
            code.PrintLine("/// </summary>");
            code.PrintLine("/// <param name=\"id\">配置表id</param>");
            code.PrintLine("/// <returns></returns>");
            code.PrintLine("public static ",sheet.SheetName," Get",sheet.SheetName,"(int id)");
            code.PrintLine("{");
            code.In();

            code.PrintLine("if(m_all"+sheet.SheetName+"s.Count==0)");
            code.PrintLine("{");
            code.In();
            code.PrintLine("Init();");
            code.Out();
            code.PrintLine("}");
            code.PrintLine(sheet.SheetName + " m_" + sheet.SheetName + "=null;");
            code.PrintLine("m_" + sheet.SheetName+"=","m_all"+sheet.SheetName+ "s.Find(a => a.id == id);");
            code.PrintLine("if(m_"+ sheet.SheetName + "==null)");
            code.PrintLine("{");
            code.In();
            code.PrintLine("Debug.LogError(m_" + sheet.SheetName + "==null);");
            code.Out();
            code.PrintLine("}");
            code.PrintLine("return " + "m_" + sheet.SheetName + ";");

            code.Out();
            code.PrintLine("}");
            code.PrintLine();


            //初始化方法 init
            code.PrintLine("/// <summary>");
            code.PrintLine("///"," 初始化");
            code.PrintLine("/// </summary>");
            code.PrintLine("private static void Init()");
            code.PrintLine("{");
            code.In();
            code.PrintLine("TextAsset textAsset=Resources.Load<TextAsset>(\"Config/",sheet.SheetName,"\");");
            code.PrintLine("string jsonInfo=textAsset.text;");
            //fan xue lie hua dai ma
            code.PrintLine("m_all",sheet.SheetName,"s",string .Format("= JsonConvert.DeserializeObject<{0}>(jsonInfo);","List<"+sheet.SheetName+">"));

            code.Out();
            code.PrintLine("}");

            code.PrintLine();
            code.Out();
            code.PrintLine("}");
            code.WriteFile(codePath + sheet.SheetName + "Ex.cs");
            Console.WriteLine(sheet.SheetName + "Ex.cs 代码生成完成");
            Log.Append(sheet.SheetName + "Ex.cs 代码生成完成\n");
        }


        static void GenertorJson(XSSFSheet sheet)
        {
            List<Dictionary<string,object>> table = new List<Dictionary<string,object>>();
            XSSFRow typeRow = (XSSFRow)sheet.GetRow(0);
            XSSFRow nameRow = (XSSFRow)sheet.GetRow(2);
            for(int i = 3; i <= sheet.LastRowNum; i++)
            {
                Dictionary<string,object> row = new Dictionary<string,object>();
                XSSFRow value = (XSSFRow)sheet.GetRow(i);
                for(int j = 0; j < value.LastCellNum; j++)
                {
                    string cellInfo = value.GetCell(j).ToString();
                    //List<Dictionary>
                    if(typeRow.GetCell(j).ToString().Contains("List<Dictionary"))
                    {
                        List<Dictionary<string,string>> list = new List<Dictionary<string,string>>();
                        string[] dictArr = cellInfo.Split('-');
                        for(int tempIndex1 = 0; tempIndex1 < dictArr.Length; tempIndex1++)
                        {
                            int IndexofA = dictArr[tempIndex1].IndexOf("[");
                            int IndexofB = dictArr[tempIndex1].IndexOf("]");
                            string temp = dictArr[tempIndex1].Substring(IndexofA + 1,IndexofB - IndexofA - 1);
                            string[] listArr= temp.Split(',');
                            Dictionary<string,string> dict = new Dictionary<string,string>();
                            dict[listArr[0]] = listArr[1];
                            list.Add(dict);
                        }
                        row[nameRow.GetCell(j).ToString()] = list;
                    }
                    //List<List
                    else if (typeRow.GetCell(j).ToString().Contains("List<List"))
                    {
                        List<List<string>> list = new List<List<string>>();
                        string[] listArr1 = cellInfo.Split('-');
                        for (int temp1 = 0; temp1 < listArr1.Length; temp1++)
                        {
                            int IndexofA = listArr1[temp1].IndexOf("[");
                            int IndexofB = listArr1[temp1].IndexOf("]");
                            string temp = listArr1[temp1].Substring(IndexofA + 1,IndexofB - IndexofA - 1);
                            string[] listArr2 = temp.Split(',');
                            List<string> tempList = new List<string>();
                            for (int temp2 = 0; temp2 < listArr2.Length; temp2++)
                            {
                                tempList.Add(listArr2[temp2]);
                            }
                            list.Add(tempList);
                        }
                        row[nameRow.GetCell(j).ToString()] = list;
                    }
                    //List
                    else if (typeRow.GetCell(j).ToString().Contains("List"))
                    {
                        List<string> list = new List<string>();
                        string[] arr = cellInfo.Split(',');
                        for (int k = 0; k < arr.Length; k++)
                        {
                            list.Add(arr[k]);
                        }
                        row[nameRow.GetCell(j).ToString()] = list;
                    }
                    //qitaleixing
                    else if (typeRow.GetCell(j).ToString().Contains("Dictionary"))
                    {
                        Dictionary<string, string> dict = new Dictionary<string, string>();
                        string[] arr = cellInfo.Split(',');
                        dict[arr[0]] = arr[1];
                        row[nameRow.GetCell(j).ToString()] = dict;
                    }
                    else
                    {
                        row[nameRow.GetCell(j).ToString()] = value.GetCell(j).ToString();
                    }
                }
                table.Add(row);
            }
            string json = string.Empty;

            //使用Json.Net进行序列化
            json = JsonConvert.SerializeObject(table);
            CodeGenerator.WriteLog(configPath + sheet.SheetName + ".json",uncode(json));
        }

        public static string uncode(string str)
        {
            return new Regex(@"\\u([0-9A-F]{4})",RegexOptions.IgnoreCase | RegexOptions.Compiled).Replace(
              str,x => string.Empty + Convert.ToChar(Convert.ToUInt16(x.Result("$1"),16)));
        }
    }

}

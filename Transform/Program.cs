using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Transform
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] DatArr = { "lc101", "lc102", "lc103", "lc104", "lc105", "lc106", "lc107", "lc108", "lc109",
                                "lc201", "lc202", "lc203", "lc204", "lc205", "lc206", "lc207", "lc208",
                                "lr101", "lr102", "lr103", "lr104", "lr105", "lr106", "lr107", "lr108", "lr109", "lr110", "lr111", "lr112",
                                "lr201", "lr202", "lr203", "lr204", "lr205", "lr206", "lr207", "lr208", "lr209", "lr210", "lr211",
                                "lrc101", "lrc102", "lrc103", "lrc104", "lrc105", "lrc106", "lrc107", "lrc108",
                                "lrc201", "lrc202", "lrc203", "lrc204", "lrc205", "lrc206", "lrc207", "lrc208", 
                              };
            for(int m = 0; m < DatArr.Count(); m++)
            {
                string DatFile = DatArr[m];
                string filepath = "G:\\PdpData\\pdp_100\\" + DatFile + ".txt";
                string strFileName = "G:\\PdpData\\pdp_100\\" + DatFile + ".xlsx";
                Console.WriteLine("读取文档：" + DatFile);
                StreamReader sr = new StreamReader(filepath, Encoding.Default);
                String line;
                List<int[]> CusData = new List<int[]>();
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.ToString();
                    string[] LineArr = line.Split('\t');
                    if (LineArr.Count() < 5) continue;
                    else
                    {
                        int[] arr = new int[LineArr.Count()];
                        for (int i = 0; i < LineArr.Count(); i++)
                        {
                            arr[i] = Convert.ToInt32(LineArr[i]);
                        }
                        CusData.Add(arr);
                    }
                }
                //创建Excel并保存
                Console.WriteLine("----------------创建EXCEL---------------");
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    //clsLog.m_CreateErrorLog("无法创建Excel对象，可能计算机未安装Excel", "", "");
                    return;
                }
                //創建Excel對象
                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                //Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                if (worksheet == null)
                {
                    worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                }
                else
                {
                    worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, worksheet, 1, Type.Missing);
                }
                Microsoft.Office.Interop.Excel.Range range = null;
                //数据复制到excel中
                int rowIndex = 0;
                worksheet.Name = "CusData";
                for (int i = 0; i < CusData.Count; i++)
                {
                    rowIndex++;
                    int[] arr = CusData[i];
                    //tripID                       
                    for (int j = 1; j <= arr.Length; j++)
                    {
                        xlApp.Cells[rowIndex, j] = arr[j - 1];
                    }
                }
                //下面是将Excel存储在服务器上指定的路径与存储的名称
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(strFileName);
                    
                }
                catch
                {
                    return;
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

                    // 很多文章上都说必须调用此方法， 但是我试过没有调用oExcel.Quit() 的情况， 进程也能安全退出，
                    //还是保留着吧。ITPUB个人空间%_.N2X%BjUFl
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    // 垃圾回收是必须的。 测试如果不执行垃圾回收， 无法关闭Excel 进程。
                    xlApp = null;
                    GC.Collect();
                    Console.WriteLine("==========完成转换！==========");
                }
            }
            
            
        }
    }
}

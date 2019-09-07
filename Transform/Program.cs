using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//用来将txt转换为excel方便后期处理
namespace Transform
{
    class Program
    {
        static void Main(string[] args)
        {
            TransTJFDat();
        }
        //针对100数量的数据进行处理
        static void TransDat100()
        {
            string[] DatArr = { "lc101", "lc102", "lc103", "lc104", "lc105", "lc106", "lc107", "lc108", "lc109",
                                "lc201", "lc202", "lc203", "lc204", "lc205", "lc206", "lc207", "lc208",
                                "lr101", "lr102", "lr103", "lr104", "lr105", "lr106", "lr107", "lr108", "lr109", "lr110", "lr111", "lr112",
                                "lr201", "lr202", "lr203", "lr204", "lr205", "lr206", "lr207", "lr208", "lr209", "lr210", "lr211",
                                "lrc101", "lrc102", "lrc103", "lrc104", "lrc105", "lrc106", "lrc107", "lrc108",
                                "lrc201", "lrc202", "lrc203", "lrc204", "lrc205", "lrc206", "lrc207", "lrc208",
                              };
            string FileAdress = "G:\\PdpData\\pdp_100\\";
            for (int m = 0; m < DatArr.Count(); m++)
            {
                string DatFile = DatArr[m];
                string filepath = FileAdress + DatFile + ".txt";
                string strFileName = FileAdress + DatFile + ".xlsx";
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
        
        //针对200数量的数据进行处理
        static void TransDat200()
        {
            string[] DatArr = { "LC1_2_1","LC1_2_2","LC1_2_3","LC1_2_4","LC1_2_5",
                                "LC1_2_6","LC1_2_7","LC1_2_8","LC1_2_9","LC1_2_10",
                                "LC2_2_1","LC2_2_2","LC2_2_3","LC2_2_4","LC2_2_5",
                                "LC2_2_6","LC2_2_7","LC2_2_8","LC2_2_9","LC2_2_10",
                                "LR1_2_1","LR1_2_2","LR1_2_3","LR1_2_4","LR1_2_5",
                                "LR1_2_6","LR1_2_7","LR1_2_8","LR1_2_9","LR1_2_10",
                                "LR2_2_1","LR2_2_2","LR2_2_3","LR2_2_4","LR2_2_5",
                                "LR2_2_6","LR2_2_7","LR2_2_8","LR2_2_9","LR2_2_10",
                                "LRC1_2_1","LRC1_2_2","LRC1_2_3","LRC1_2_4","LRC1_2_5",
                                "LRC1_2_6","LRC1_2_7","LRC1_2_8","LRC1_2_9","LRC1_2_10",
                                "LRC2_2_1","LRC2_2_2","LRC2_2_3","LRC2_2_4","LRC2_2_5",
                                "LRC2_2_6","LRC2_2_7","LRC2_2_8","LRC2_2_9","LRC2_2_10",
                              };
            string FileAdress = "D:\\VRP-Route-Data\\PDPTW\\pdp_200\\";
            for (int m = 0; m < DatArr.Count(); m++)
            {
                string DatFile = DatArr[m];
                string filepath = FileAdress + DatFile + ".txt";
                string strFileName = FileAdress + DatFile + ".xlsx";
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

        //针对200数量的数据进行处理
        static void TransDatSolomon()
        {
            string[] DatArr = {
                                "r101", "r102","r103", "r104", "r105","r106","r107","r108","r109","r110","r111","r112",
                                "r201", "r202","r203","r204", "r205", "r206", "r207","r208","r209","r210","r211",
                                "c101", "c102","c103", "c104", "c105", "c106","c107", "c108","c109",
                                "c201","c202","c203", "c204","c205","c206","c207","c208",
                                "rc101", "rc102", "rc103", "rc104", "rc105", "rc106", "rc107",
                                "rc201", "rc202", "rc203", "rc204", "rc205", "rc206", "rc207", "rc208"
                              };
            string FileAdress = "D:\\VRP-Route-Data\\solomen数据\\In\\";
            for (int m = 0; m < DatArr.Count(); m++)
            {
                string DatFile = DatArr[m];
                string filepath = FileAdress + DatFile + ".txt";
                string strFileName = FileAdress + DatFile + ".xlsx";
                Console.WriteLine("读取文档：" + DatFile);
                StreamReader sr = new StreamReader(filepath, Encoding.Default);
                String line;
                List<int[]> CusData = new List<int[]>();
                bool CheckNum = false;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.ToString();
                    string[] LineArr = line.Split(' ');
                    LineArr = LineArr.Where(x => x != "").ToArray();
                    if (LineArr.Count() < 5) continue;
                    else
                    {
                        if(!CheckNum)
                        {
                            CheckNum = true;
                            continue;
                        }
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

        //针对实例数据进行处理
        static void TransTJFDat()
        {
            string[] DatArr = {
                "SR80U401", "SC90U401", "SRC70U401", "SRC65S401", "SR85S401", "SC75S401",
                "SR228U401", "SC200U401", "SRC186U401", "SR150S401", "SC260S401", "SRC290S401",
                "SR430U401", "SC450U401", "SRC686U401", "SR330S401", "SC620S401", "SRC460S401",
                              };
            string FileAdress = "D:\\VRP-Route-Data\\机场接送实例数据\\随机实例文档\\随机实例文档\\";
            //打开0点坐标
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\VRP-Route-Data\\机场接送实例数据\\随机实例文档\\vrp实例随机产生速度和机场坐标——MYY.xlsx;Extended Properties='Excel 12.0;HDR=no;IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            OleDbDataAdapter myCommand1 = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", strConn);
            DataTable tripData = new DataTable();
            try
            {
                myCommand1.Fill(tripData);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception(ex.Message);
            }
            for (int m = 0; m < DatArr.Count(); m++)
            {
                string DatFile = DatArr[m];
                string filepath = FileAdress + DatFile + ".txt";
                string strFileName = FileAdress + DatFile + ".xlsx";
                Console.WriteLine("读取文档：" + DatFile);
                StreamReader sr = new StreamReader(filepath, Encoding.Default);
                String line;
                List<double[]> CusData = new List<double[]>();
                //0点
                double[] arr = new double[6];
                arr[1] = Convert.ToDouble(tripData.Rows[m + 1][3].ToString());
                arr[2] = Convert.ToDouble(tripData.Rows[m + 1][4].ToString());
                arr[4] = 24;
                arr[5] = Convert.ToDouble(tripData.Rows[m + 1][2].ToString());
                CusData.Add(arr);
                //customer data
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.ToString();
                    string[] LineArr = line.Split(' ');
                    arr = new double[6];
                    for (int i = 0; i < 3; i++)
                    {
                        arr[i] = Convert.ToInt32(LineArr[i]);
                    }
                    //early time
                    DateTime dt = Convert.ToDateTime(LineArr[3]);
                    double cusTim = Math.Round(dt.Hour + dt.Minute / 60.0, 3);
                    arr[3] = cusTim;
                    //last time
                    dt = Convert.ToDateTime(LineArr[10]);
                    cusTim = Math.Round(dt.Hour + dt.Minute / 60.0, 3);
                    arr[4] = cusTim;
                    //request
                    arr[5] = Convert.ToInt32(LineArr[11]);
                    CusData.Add(arr);
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
                    arr = CusData[i];
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

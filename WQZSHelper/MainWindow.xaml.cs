using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace WQZSHelper
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /*
         * 每天定位记录查询下来
         * 日期 定位次数 0-30 31-60 61-120 121-240 240
         */

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var open = new Microsoft.Win32.OpenFileDialog();

            if (open.ShowDialog().GetValueOrDefault())
            {
                IWorkbook workbook = new HSSFWorkbook(open.OpenFile());
                
                foreach (ISheet sheet in workbook)
                {
                    var information = new List<LocationInfo>();
                    var lastRowNum = sheet.LastRowNum;
                    var n = 0;
                    int n1 = 0, n2 = 0;
                    IRow row;
                    for (int i = 1; i <= lastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        information.Add(new LocationInfo()
                                            {
                                                Name = row.GetCell(0).StringCellValue,
                                                Date = DateTime.Parse(row.GetCell(7).StringCellValue)
                                            });
                    }


                    var dates = information.GroupBy(d => d.Date.Date);

                    var beginDate = new DateTime(2014, 10, 17);
                    var endDate = new DateTime(2014, 11, 17);
                    var dateDiff = (endDate - beginDate).Days;

                    for (var date = beginDate; date <= endDate; date = date.AddDays(1), n++)
                    {
                        row = sheet.CreateRow(lastRowNum + 2 + n);
                        row.CreateCell(0).SetCellValue(date.Date.ToString("MM-dd"));
                        if (dates.Any(d => d.Key.ToString().Equals(date.Date.ToString())))
                        {
                            var infoByDate = dates.First(d => d.Key.ToString().Equals(date.Date.ToString()));

                            row.CreateCell(0).SetCellValue(infoByDate.Key.ToString("MM-dd"));
                            var count = infoByDate.Count();
                            if (count > 28)
                                n1++;
                            else if (count > 14)
                                n2++;
                            row.CreateCell(1).SetCellValue(count);



                            var infos = infoByDate.OrderBy(d => d.Date);
                            var num = new Number();

                            LocationInfo lastInfo = null;
                            foreach (LocationInfo info in infos)
                            {
                                if (lastInfo != null)
                                    num.AddNumber(lastInfo.Date, info.Date);
                                lastInfo = info;
                            }


                            row.CreateCell(3).SetCellValue(num.N1);
                            row.CreateCell(4).SetCellValue(num.N2);
                            row.CreateCell(5).SetCellValue(num.N3);
                            row.CreateCell(6).SetCellValue(num.N4);
                            row.CreateCell(7).SetCellValue(num.N5);

                        }
                        else
                        {
                            //row.CreateCell(1).SetCellValue(0);
                        }
                    }


                    row = sheet.CreateRow(lastRowNum + 2 + n + 2);
                    row.CreateCell(0).SetCellValue(sheet.SheetName);
                    row.CreateCell(1).SetCellValue(n1);
                    row.CreateCell(2).SetCellValue(n2);
                    row.CreateCell(3).SetCellValue(dateDiff - n1 - n2);


                    //row = null;
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Write(ms);
                    ms.Flush();
                    ms.Position = 0;
                    File.WriteAllBytes(open.FileName + 1, ms.ToArray());

                }                
            }
            MessageBox.Show("OK");
        }
    }





    public class LocationInfo
    {
        public string Name { get; set; }
        public DateTime Date { get; set; }
    }

    public class LocationInfos
    {
        public List<LocationInfo> Infos { get; set; }
    }

    /// <summary>
    /// 数字分类
    /// </summary>
    public struct Number
    {
        /// <summary>
        /// 0-30
        /// </summary>
        public int N1;
        /// <summary>
        /// 31-60
        /// </summary>
        public int N2;
        /// <summary>
        /// 61-120
        /// </summary>
        public int N3;
        /// <summary>
        /// 121-240
        /// </summary>
        public int N4;
        /// <summary>
        /// 241
        /// </summary>
        public int N5;



        public void AddNumber(DateTime d1, DateTime d2)
        {
            var diff = Math.Abs((d2 - d1).TotalMinutes);
            if (diff < 31)
                N1++;
            else if (diff < 61)
                N2++;
            else if (diff < 121)
                N3++;
            else if (diff < 241)
                N4++;
            else
                N5++;
        }
    }



}

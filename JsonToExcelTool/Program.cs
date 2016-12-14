using JsonToExcelTool.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcelTool
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Local json file? y/n");
            var str = Console.ReadLine();
            if (str != "y" && str != "n")
            {
                Console.WriteLine("Please input y or n");
                str = Console.ReadLine();
            }

            string jsonStr = "";
            if (str == "y")
            {
                Console.WriteLine("Please put your local file in the TxtFile folder.Then press any key to continue");
                Console.ReadKey();
               
                string[] filenames = Directory.GetFiles(Environment.CurrentDirectory+ "\\TxtFile", "*.txt", SearchOption.AllDirectories);
                if (filenames.Length == 0)
                {
                    Console.WriteLine("Do not get the file,Please put your local file in the TxtFile folder then reopen the program");
                    Console.ReadKey();
                    return;
                    
                }
                for (int i = 0; i < filenames.Length; i++)
                {
                    string[] fileNameArr = filenames[i].Split('\\');
                    string fileName = fileNameArr[fileNameArr.Length - 1].Split('.')[0];
                    jsonStr = File.ReadAllText(filenames[i]);
                    CreateExcelFile(jsonStr,fileName);
                }
                 
               
            }
            else if (str == "n")
            {
                Console.WriteLine("Await download data from web...");
               GetRomoteJson();  
            }
            else
            {
                Console.WriteLine("Program has quit,Please reopen it");
                Console.ReadLine();
                return;
            } 
           
            Console.ReadKey();


        }


        public static void CreateExcelFile(string jsonStr, string fileName = "")
        {

            DtoModels model = Newtonsoft.Json.JsonConvert.DeserializeObject<DtoModels>(jsonStr);

            HSSFWorkbook wk = new HSSFWorkbook();

            ICellStyle style11 = wk.CreateCellStyle();
            style11.DataFormat = 194;

            //创建一个Sheet  
            ISheet sheet = wk.CreateSheet("sheet1");
            //在第一行创建行  
            IRow row = sheet.CreateRow(0);

            ICell cell0 = row.CreateCell(0);
            cell0.SetCellValue("offerType");

            ICell cell1 = row.CreateCell(1);
            cell1.SetCellValue("linkDest");

            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue("convertsOn");

            ICell cell3 = row.CreateCell(3);
            cell3.SetCellValue("appInfo/appID");

            ICell cell4 = row.CreateCell(4);
            cell4.SetCellValue("appInfo/previewLink");

            ICell cell5 = row.CreateCell(5);
            cell5.SetCellValue("appInfo/appName");

            ICell cell6 = row.CreateCell(6);
            cell6.SetCellValue("appInfo/appCategory");

            ICell cell7 = row.CreateCell(7);
            cell7.SetCellValue("targets/offerID");

            ICell cell8 = row.CreateCell(8);
            cell8.SetCellValue("targets/approvalStatus");

            ICell cell9 = row.CreateCell(9);
            cell9.SetCellValue("targets/offerStatus");

            ICell cell10 = row.CreateCell(10);
            cell10.SetCellValue("targets/trackingLink");

            ICell cell11 = row.CreateCell(11);
            cell11.SetCellValue("targets/countries");

            ICell cell12 = row.CreateCell(12);
            cell12.SetCellValue("targets/platforms");

            ICell cell13 = row.CreateCell(13);
            cell13.SetCellValue("targets/payoutUSD");

            ICell cell14 = row.CreateCell(14);
            cell14.SetCellValue("targets/endDate");

            ICell cell15 = row.CreateCell(15);
            cell15.SetCellValue("targets/dailyConversionCap");

            ICell cell16 = row.CreateCell(16);
            cell16.SetCellValue("targets/restrictions/allowIncent");

            ICell cell17 = row.CreateCell(17);
            cell17.SetCellValue("targets/restrictions/deviceIDRequired");

            ICell cell18 = row.CreateCell(18);
            cell18.SetCellValue("targets/restrictions/minOSVersion");


            ICell cell19 = row.CreateCell(19);
            cell19.SetCellValue("requestID");


            ICellStyle style = wk.CreateCellStyle();//创建样式对象
            IFont font = wk.CreateFont(); //创建一个字体样式对象
            font.FontName = "微软雅黑"; //和excel里面的字体对应  
            font.FontHeightInPoints = 12;//字体大小
            font.Boldweight = short.MinValue;//字体加粗
            style.SetFont(font); //将字体样式赋给样式对象

            for (int j = 0; j <= 19; j++)
            {
                row.GetCell(j).CellStyle = style;//把样式赋给单元格
            }

            int index1 = 0;
            foreach (var m in model.offers)
            {
                foreach (var target in m.targets)
                {

                    index1++;
                    IRow row1 = sheet.CreateRow(index1);
                    ICell cel0 = row1.CreateCell(0);
                    cel0.SetCellValue(m.offerType);

                    ICell cel1 = row1.CreateCell(1);
                    cel1.SetCellValue(m.linkDest);

                    ICell cel2 = row1.CreateCell(2);
                    cel2.SetCellValue(m.convertsOn);

                    ICell cel3 = row1.CreateCell(3);
                    cel3.SetCellValue(m.appInfo.appID);

                    ICell cel4 = row1.CreateCell(4);
                    cel4.SetCellValue(m.appInfo.previewLink);

                    ICell cel5 = row1.CreateCell(5);
                    cel5.SetCellValue(m.appInfo.appName);

                    ICell cel6 = row1.CreateCell(6);
                    cel6.SetCellValue(m.appInfo.appCategory);

                    ICell cel7 = row1.CreateCell(7);
                    cel7.SetCellValue(target.offerID);

                    ICell cel8 = row1.CreateCell(8);
                    cel8.SetCellValue(target.approvalStatus);

                    ICell cel9 = row1.CreateCell(9);
                    cel9.SetCellValue(target.offerStatus);

                    ICell cel10 = row1.CreateCell(10);
                    cel10.SetCellValue(target.trackingLink);

                    ICell cel11 = row1.CreateCell(11);
                    cel11.SetCellValue(string.Join(",", target.countries));

                    ICell cel12 = row1.CreateCell(12);
                    cel12.SetCellValue(string.Join(",", target.platforms));

                    ICell cel13 = row1.CreateCell(13);
                    cel13.SetCellValue(target.payout.amount.ToString());
                    cel13.CellStyle = style11;  //设置格式为数字

                    ICell cel14 = row1.CreateCell(14);
                    cel14.SetCellValue(target.endDate);

                    ICell cel15 = row1.CreateCell(15);
                    cel15.SetCellValue(target.dailyConversionCap);

                    ICell cel16 = row1.CreateCell(16);
                    cel16.SetCellValue(target.restrictions.allowIncent);

                    ICell cel17 = row1.CreateCell(17);
                    cel17.SetCellValue(target.restrictions.deviceIDRequired);

                    ICell cel18 = row1.CreateCell(18);
                    cel18.SetCellValue(target.restrictions.minOSVersion);

                    ICell cel19 = row1.CreateCell(19);
                    cel19.SetCellValue(model.requestID);

                }

            }

            for (int i = 0; i <= 19; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            fileName = fileName == "" ? DateTime.Now.ToString("yyyyMMddhhmmss") : fileName;
            //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建时不要打开该文件  
            using (FileStream fs = File.OpenWrite(Environment.CurrentDirectory + "\\Excel\\" + fileName + ".xls"))
            {
                wk.Write(fs);//向打开的这个xls文件中写入并保存。  
            }

            Console.WriteLine("Excel built success:File name is " + fileName + ".xls");

        }




        public static async void GetRomoteJson()
        {
            string Uri = "http://api.apptap.com/api/2/offers_feed?pubID=wawru5reyuya_mcaf&siteID=offers-feed";

            HttpClient httpClient = new HttpClient();

            // 创建一个异步GET请求，当请求返回时继续处理（Continue-With模式）  
            HttpResponseMessage response = await httpClient.GetAsync(Uri);
            response.EnsureSuccessStatusCode();
            string resultStr = await response.Content.ReadAsStringAsync();
            CreateExcelFile(resultStr);
        }

    }
}

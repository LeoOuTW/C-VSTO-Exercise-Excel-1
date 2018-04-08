using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using StockExcelVSTO.ViewModel;

namespace StockExcelVSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnUploadStockData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
                UploadStockViewModel uploadClass = new UploadStockViewModel();
                List<UploadStockDataViewModel> dataList = new List<UploadStockDataViewModel>();
                int versionData = 0;

                int i = 2;
                while ((currentSheet.Cells[i, 1] as Range).Value != null)
                {
                    for (int iMonth = 1; iMonth <= 12; iMonth++)
                    {
                        var Version = (currentSheet.Cells[i, 1] as Range).Value.ToString();
                        var StockYear = (currentSheet.Cells[i, 2] as Range).Value.ToString();
                        var StockMonth = (currentSheet.Cells[1, iMonth + 3] as Range).Value.ToString();
                        var StockNum = (currentSheet.Cells[i, iMonth + 3] as Range).Value.ToString();
                        if (StockNum == "")
                        {
                            StockNum = "0";
                        }
                        var MaterialName = (currentSheet.Cells[i, 3] as Range).Value.ToString().Trim();

                        dataList.Add(new UploadStockDataViewModel()
                        {
                            ConfirmMonth = ConvertMonth(StockMonth),
                            ConfirmYear = Convert.ToInt16(StockYear),
                            Name = MaterialName,
                            StockNum = Convert.ToInt32(StockNum),
                        });
                        versionData = Convert.ToInt16(Version);
                        //MessageBox.Show(Version + " / " + MaterialName + " / " + StockYear + " / " + StockMonth + " / " + StockNum);
                    }
                    i++;
                }

                if (dataList.Count > 0 && versionData != 0)
                {

                    //check year 是否一致, 是否為數值型態
                    if (Validation(dataList) == true)
                    {
                        uploadClass.Version = versionData;
                        uploadClass.StockDataList = dataList;


                        var myContent = JsonConvert.SerializeObject(uploadClass);
                        var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://xxxx.azurewebsites.net/api/Stock");
                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Method = "POST";

                        using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {

                            streamWriter.Write(myContent);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            if (httpResponse.StatusCode == HttpStatusCode.OK)
                            {
                                var result = streamReader.ReadToEnd();

                                var value = JsonConvert.DeserializeObject(result);

                                var msg = JsonConvert.DeserializeObject<ReturnMsgViewModel>(result);

                                MessageBox.Show("資料上傳成功，上傳後最新版本為第" + msg.value.version.ToString() + "版");
                            }
                            else
                            {
                                MessageBox.Show("資料上傳有誤，請洽詢系統管理員");
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("沒有資料可以上傳，如果問題，請洽詢系統管理員");
                }
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                MessageBox.Show("資料有誤，請重新檢視您的資料");
            }

        }

        private bool Validation(List<UploadStockDataViewModel> model)
        {
            //check year is 不一致
            var sameYear = model.GroupBy(x => x.ConfirmYear).Select(x => x.First()).ToList();
            if (sameYear.Count > 1)
            {
                MessageBox.Show("請填寫相同的年度");
                return false;
            }
            //check name is empty
            var emptyName = model.Where(x => x.Name == "");
            if (emptyName.Any())
            {
                MessageBox.Show("Item名稱有空值，請補上名稱或刪除");
                return false;
            }
            //check month equal 0
            var zeroMonth = model.Where(x => x.ConfirmMonth == 0);
            if (zeroMonth.Any())
            {
                MessageBox.Show("月份有誤，請確認");
                return false;
            }

            return true;
        }

        private int ConvertMonth (string monthName)
        {
            switch (monthName)
            {
                case "Jan":
                    return 1;
                case "Feb":
                    return 2;
                case "Mar":
                    return 3;
                case "Apr":
                    return 4;
                case "May":
                    return 5;
                case "Jun":
                    return 6;
                case "Jul":
                    return 7;
                case "Aug":
                    return 8;
                case "Sep":
                    return 9;
                case "Oct":
                    return 10;
                case "Nov":
                    return 11;
                case "Dec":
                    return 12;
                default:
                    return 0;
            }
        }



    }
}

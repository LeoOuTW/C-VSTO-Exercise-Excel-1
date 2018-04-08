using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockExcelVSTO.ViewModel
{
    class StockViewModel
    {
    }

    public class UploadStockViewModel
    {
        public int Version { get; set; }
        public List<UploadStockDataViewModel> StockDataList { get; set; }
    }

    public class UploadStockDataViewModel
    {
        public string Name { get; set; }

        public int StockNum { get; set; }

        public int ConfirmYear { get; set; }

        public int ConfirmMonth { get; set; }
    }

    public class ReturnMsgViewModel
    {
        public string contentType { get; set; }
        public string serializerSettings { get; set; }
        public string statusCode { get; set; }
        public ReturnVersion value { get; set; }
    }

    public class ReturnVersion
    {
        public int version { get; set; }
    }
}

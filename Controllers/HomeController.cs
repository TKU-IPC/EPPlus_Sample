using EPPlusExam.Infrastructure;
using EPPlusExam.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EPPlusExam.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            string subHeader = String.Format("{0} 學年記錄檔", 106);
            string[] fileTitle = { subHeader };
            List<PrintModel1> data = new List<PrintModel1>() { 
                  new PrintModel1(){ LOG_DATE = DateTime.Now, LOG_HOST = "163.13.241.1", LOG_MSG = "測試登入1", LOG_MSGID = "AC01", LOG_PROGID = "AC110", LOG_USERID = "116534"}
                , new PrintModel1(){ LOG_DATE = DateTime.Now, LOG_HOST = "163.13.241.2", LOG_MSG = "測試登入2", LOG_MSGID = "AC02", LOG_PROGID = "AC110", LOG_USERID = "116534"}
                , new PrintModel1(){ LOG_DATE = DateTime.Now, LOG_HOST = "163.13.241.3", LOG_MSG = "測試登入3", LOG_MSGID = "AC03", LOG_PROGID = "AC110", LOG_USERID = "116534"}
                , new PrintModel1(){ LOG_DATE = DateTime.Now, LOG_HOST = "163.13.241.4", LOG_MSG = "測試登入4", LOG_MSGID = "AC04", LOG_PROGID = "AC110", LOG_USERID = "116534"}
                , new PrintModel1(){ LOG_DATE = DateTime.Now, LOG_HOST = "163.13.241.5", LOG_MSG = "測試登入5", LOG_MSGID = "AC05", LOG_PROGID = "AC110", LOG_USERID = "116534"}            
            };
            var pdfFile = (new EPPlusExporter()).GenerateXlsx<IXlsHeaderSchema, PrintModel1>(fileTitle, null, data.OrderBy(c => c.LOG_DATE), null);
            return File(pdfFile, "application/vnd.ms-excel", Server.UrlPathEncode("EPPlustXlsFile1.xlsx"));
        }

        public ActionResult Contact()
        {
            var headers = new List<EPPlusXlsHeader>();
            headers.Add(new EPPlusXlsHeader() { name = "106年度" + System.Environment.NewLine + "決算數", align = 1, x1 = 1, y1 = 1, x2 = 1, y2 = 2, width = 20 });
            headers.Add(new EPPlusXlsHeader() { name = "科　　　　　　　目", align = 1, x1 = 2, y1 = 1, x2 = 3, y2 = 1 });
            headers.Add(new EPPlusXlsHeader() { name = "編號", align = 1, x1 = 2, y1 = 2, x2 = 2, y2 = 2, width = 10 });
            headers.Add(new EPPlusXlsHeader() { name = "名　　　稱", align = 1, x1 = 3, y1 = 2, x2 = 3, y2 = 2, width = 30 });

            headers.Add(new EPPlusXlsHeader() { name =  "106年度" + System.Environment.NewLine + "預算數", align = 1, x1 = 4, y1 = 1, x2 = 4, y2 = 2, width = 20 });
            headers.Add(new EPPlusXlsHeader() { name = "估計107年度" + System.Environment.NewLine + "決算數", align = 1, x1 = 5, y1 = 1, x2 = 5, y2 = 2, width = 20 });

            headers.Add(new EPPlusXlsHeader() { name = "106年度預算與估計107年度決算比較", align = 1, x1 = 6, y1 = 1, x2 = 7, y2 = 1, width = 40 });
            headers.Add(new EPPlusXlsHeader() { name = "差　　異", align = 1, x1 = 6, y1 = 2, x2 = 6, y2 = 2, width = 20 });
            headers.Add(new EPPlusXlsHeader() { name = "％", align = 1, x1 = 7, y1 = 2, x2 = 7, y2 = 2, width = 20 });
            headers.Add(new EPPlusXlsHeader() { name = "備　　　　　註", align = 1, x1 = 8, y1 = 1, x2 = 8, y2 = 2, width = 40 });

            List<PrintModel2> data = new List<PrintModel2>() { 
                  new PrintModel2(){ ACT_AMT_DIFF = 1000, ACT_BUG_AMT = 3000, ACT_EXP_INIT = 5000, ACT_ID = "1121", ACT_NAME = "會計科目1" , ACT_PREV_AMT = 3000, BUG_RATE = 13 , DOC_DOC = "測試資料1"}
                  , new PrintModel2(){ ACT_AMT_DIFF = 1200, ACT_BUG_AMT = 3300, ACT_EXP_INIT = 5000, ACT_ID = "1122", ACT_NAME = "會計科目2" , ACT_PREV_AMT = 4000, BUG_RATE = 23 , DOC_DOC = "測試資料2"}
                  , new PrintModel2(){ ACT_AMT_DIFF = 1300, ACT_BUG_AMT = 3500, ACT_EXP_INIT = 5000, ACT_ID = "1123", ACT_NAME = "會計科目3" , ACT_PREV_AMT = 5000, BUG_RATE = 33 , DOC_DOC = "測試資料3"}
                  , new PrintModel2(){ ACT_AMT_DIFF = 2300, ACT_BUG_AMT = 3600, ACT_EXP_INIT = 5000, ACT_ID = "1124", ACT_NAME = "會計科目4" , ACT_PREV_AMT = 6000, BUG_RATE = 41 , DOC_DOC = "測試資料4"}
                  , new PrintModel2(){ ACT_AMT_DIFF = 2400, ACT_BUG_AMT = 3700, ACT_EXP_INIT = 5000, ACT_ID = "1125", ACT_NAME = "會計科目5" , ACT_PREV_AMT = 7000, BUG_RATE = 42 , DOC_DOC = "測試資料5"}
                  , new PrintModel2(){ ACT_AMT_DIFF = 2500, ACT_BUG_AMT = 3900, ACT_EXP_INIT = 5000, ACT_ID = "1126", ACT_NAME = "會計科目6" , ACT_PREV_AMT = 8000, BUG_RATE = 43 , DOC_DOC = "測試資料6"}
            };


            string subHeader = String.Format("{0} 預算明細表", 106);
            string[] fileTitle = { subHeader };

            var pdfFile = (new EPPlusExporter()).GenerateXlsx<IXlsHeaderSchema, PrintModel2>(fileTitle, headers, data.OrderBy(a => a.ACT_ID), null);
            return File(pdfFile, "application/vnd.ms-excel", Server.UrlPathEncode("EPPlustXlsFile2.xlsx"));
        }
    }
}
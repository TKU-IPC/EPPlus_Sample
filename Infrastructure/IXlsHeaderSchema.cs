using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPPlusExam.Infrastructure
{
    interface IXlsHeaderSchema
    {
        string name { set; get; }
        string type { set; get; }
        int? width { set; get; }
        int? height { set; get; }
        int? align { set; get; }  //0:右,1:中,2:右
        int x1 { set; get; }
        int y1 { set; get; }
        int x2 { set; get; }
        int y2 { set; get; }
    }
}
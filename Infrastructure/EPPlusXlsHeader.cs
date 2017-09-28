using EPPlusExam.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPPlusExam.Infrastructure
{
    public class EPPlusXlsHeader : IXlsHeaderSchema
    {
        public string name { set; get; }
        public string type { set; get; }
        public int? width { set; get; }
        public int? height { set; get; }
        public int? align { set; get; }  //0:右,1:中,2:右
        public int x1 { set; get; }
        public int y1 { set; get; }
        public int x2 { set; get; }
        public int y2 { set; get; }
    }
}
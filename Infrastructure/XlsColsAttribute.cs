using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPPlusExam.Infrastructure
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method | AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class XlsColsCfgAttribute : Attribute
    {
        private bool _isGenerate = true;
        private bool _isShowZero = true;
        private bool _isPercentSymbol = true;   //是否需要百分比符號(預設是需要的)
        private int _colWidth = 20;        //欄位寬度
        private int _pointLenNum = 2;   //小數位數(預設２位)
        private int _colPosIdx = 0;   //欄位顯示的順序(由左至右, 0 表示由schema由決定)
        private int _colAlign = 0;   //0(左),1(中),2(右) 預設靠左

        public bool IsGenerate { set { this._isGenerate = value; } get { return this._isGenerate; } }
        public bool IsShowZero { set { this._isShowZero = value; } get { return this._isShowZero; } }
        public bool IsPercentSymbol { set { this._isPercentSymbol = value; } get { return this._isPercentSymbol; } }
        public int ColWidth { set { this._colWidth = value; } get { return this._colWidth; } }
        public int PointLenNum { set { this._pointLenNum = value; } get { return this._pointLenNum; } }
        public int ColPosIdx { set { this._colPosIdx = value; } get { return this._colPosIdx; } }
        public int ColAlign { set { this._colAlign = value; } get { return this._colAlign; } }

        public string ColName { set; get; }
        public string ColTypeName { set; get; }
    }

    public class ColHeaderCfg
    {
        public bool IsGenerate { set; get; }
        public bool IsShowZero { set; get; }
        public bool IsPercentSymbol { set; get; }
        public int ColWidth { set; get; }
        public int PointLenNum { set; get; }
        public int ColPosIdx { set; get; }
        public int ColAlign { set; get; }
        public string ColName { set; get; }
        public string ColTypeName { set; get; }

    }

}
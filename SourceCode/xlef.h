/*******************************************************************
 *                                                                 *
 * 此class由kx設計，並發佈初版                                     *
 * 二版則由Edison.Shih.提供函式，補足初版之不足                    *
 *                                                                 *
 * 由Edison.Shih.提供的函式補足，會做edisonx的記號                 *
 *                                                                 *
 * Class由 Visual C++ 6 開發                                       *
 * 適用於Microsoft Excel 2003 於 Microsoft Excel 2003 SP3 測試無誤 *
 * 測試平台 Microsoft Windows XP SP3                               *
 *                                                   2011/7/11     *
 *******************************************************************/
#ifndef __XLEF_H__
#define __XLEF_H__

#include "excel.h"
#include "stdafx.h"

enum Boarder_Style
{
    BS_NONE    = 0,      //無框線
    BS_SOLIDLINE,        //一般線
    BS_BIGDASH,          //小間隔虛線- - - - - -有粗細
    BS_SMALLDASH,        //大間隔虛線- - - - - -無粗細
    BS_DOTDASH,          //虛線-.-.-.-.-.-.
    BS_DASHDOTDOT,       //虛線.-..-..-..-..-.
    BS_DOUBLSOLID = 9,   //雙線============（不受粗細改變）
    BS_SLASHDASH  = 13   //雙線-/-/-/-/-/-/（不受粗細改變）
};
enum Boarder_Weight//（粗細）
{
    BA_HAITLINE = 1,     //比一般小（所以用虛線表示）
    BA_THIN,             //一般
    BA_MEDIUM,           //粗
    BA_THICK             //厚
};
enum Horizontal_Alignment
{
    HA_GENERAL = 1,
    HA_LEFT,                //edisonx
    HA_CENTER,
    HA_RIGHT,                //edisonx
    HA_FILL,                //重複至填滿    //edisonx
    HA_JUSTIFYPARA,            //段落重排（有留白邊，有自動斷行）
    HA_CENTERACROSS,        //跨欄置中（不合拼儲存格）
    HA_JUSTIFY,                //分散對齊（縮排）
};

enum Vertical_Alignment
{
    VA_TOP = 1,        //edisonx
    VA_CENTER,            //edisonx
    VA_BOTTOM,            //edisonx
    VA_JUSTIFYPARA,    //段落重排（有留白邊，有自動斷行）
    VA_JUSTIFY        //分散對齊
};

enum Histogram_Chart_Type
{
    CT_AREA = 0,     //區域
    CT_COLUMN,         //方柱
    CT_CONE,        //圓錐
    CT_CYLINDER,    //圓柱
    CT_PYRAMID        //金字塔
};

enum Stock_Type
{
    ST_HLC = 0,    //最高-最低-收盤
    ST_OHLC,    //開盤-最高-最低-收盤
    ST_VHLC,    //成交量-最高-最低-收盤
    ST_VOHLC    //成交量-開盤-最高-最低-收盤
};
///////////////////////////////////
//Boarder
//Set Boarder Style
// #define BS_NONE          0  //無框線
// #define BS_SOLIDLINE     1  //一般線
// #define BS_BIGDASH       2  //小間隔虛線- - - - - -有粗細
// #define BS_SMALLDASH     3  //大間隔虛線- - - - - -無粗細
// #define BS_DOTDASH       4  //虛線-.-.-.-.-.-.
// #define BS_DASHDOTDOT    5  //虛線.-..-..-..-..-.
// //6, 7, 8 = 1
// #define BS_DOUBLSOLID    9  //雙線============（不受粗細改變）
// //10, 11, 12 = 1
// #define BS_SLASHDASH    13  //雙線-/-/-/-/-/-/（不受粗細改變）

//Set Boarder Weight（粗細）
// #define BA_HAITLINE    1  //比一般小（所以用虛線表示）
// #define BA_THIN        2  //一般
// #define BA_MEDIUM      3  //粗
// #define BA_THICK       4  //厚
//Set Boarder Color
//0-56

///////////////////////////////////
//Alignment
//Set Horizontal Alignment
// #define HA_GENERAL            1//通用格式
// #define HA_LEFT                2                //edisonx
// #define HA_CENTER            3
// #define HA_RIGHT            4                //edisonx
// #define HA_FILL                5//重複至填滿    //edisonx
// #define HA_JUSTIFYPARA         6//段落重排（有留白邊，有自動斷行）
// #define HA_CENTERACROSS        7//跨欄置中（不合拼儲存格）
// #define HA_JUSTIFY            8//分散對齊（縮排）
//Set Vertical Alignment    
// #define VA_TOP                1                //edisonx
// #define VA_CENTER            2                //edisonx
// #define VA_BOTTOM            3                //edisonx
// #define VA_JUSTIFYPARA        4//段落重排（有留白邊，有自動斷行）
// #define VA_JUSTIFY            5//分散對齊

///////////////////////////////////
//Histogram Chart Type
// #define CT_AREA        0    //區域
// #define CT_COLUMN    1    //方柱
// #define CT_CONE        2    //圓錐
// #define CT_CYLINDER    3    //圓柱
// #define CT_PYRAMID    4    //金字塔
//Stock Type
// #define ST_HLC        0    //最高-最低-收盤
// #define ST_OHLC        1    //開盤-最高-最低-收盤
// #define ST_VHLC        2    //成交量-最高-最低-收盤
// #define ST_VOHLC    3    //成交量-開盤-最高-最低-收盤

class xlsFile
{
    COleVariant VOptional, VTRUE, VFALSE;  
    //VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR)
    //VFalse((short)FALSE)
    //VTrue((short)TRUE)
    _Application objApp;
    Workbooks objBooks;
    _Workbook objBook;
    Sheets objSheets;
    _Worksheet objSheet,objSheetT;
    Range range,col,row;//,oCell;//,range2,range3;
    Interior cell;
    Font font;
    COleException e;

    LPDISPATCH lpDisp;
    ChartObjects chartobjects;
    ChartObject chartobject;
    _Chart xlsChart;
    VARIANT var;
    
    Shapes shapes;

    char buf[200];  //暫存的字串
    char buf1[200];
    char buf2[200];
      
public:
    xlsFile();
    ~xlsFile();
    //xlsFile& //開了檔案之後可以繼續選擇Sheet和命名
    xlsFile& New();
    xlsFile& Open();
    xlsFile& Open(const char*);
    void SaveAs(const char*);
    void Save();
    //void Quit(CString FileName);
    void Quit();
    void SetVisible(bool);//設定視窗為看得見，並把控制權交給使用者
    //----------------------------------------------------
    //Sheet操作
    long SheetTotal();                        //取得 Sheet 個數
    void SetSheetName(short, const char*);    //由SheetNumber    指定SheetName
    CString GetSheetName(short);              //由SheetNumber    取得SheetName
    
    xlsFile& SelectSheet(const char*);        //由SheetName      選擇Sheet
    xlsFile& SelectSheet(short);              //由SheetNumber    選擇Sheet
    void CopySheet(const char*);              //複製SheetName    指定插入Sheet的位置，並指定新Sheet名稱
    void CopySheet(short);                    //複製SheetNumber  指定插入Sheet的位置，並指定名稱
    void DelSheet(const char*);               //選SheetName      刪除Sheet
    void DelSheet(short);                     //選SheetNumber    刪除Sheet
    //-----------------------------------------------------
    //
    long GetHorztlStartCell(); // 起始行
    long GetVrticlStartCell(); // 起始列
    long GetHorztlTotalCell(); // 總行數
    long GetVrticlTotalCell(); // 總列數
    //-----------------------------------------------------
    //xlsFile& 選了格子之後可以繼續下「讀」「寫」的成員函數
    //選一格
    xlsFile& SelectCell(const char* );    
    xlsFile& SelectCell(const char* , int );
    xlsFile& SelectCell(char,int);    
    xlsFile& SelectCell(char,char,int);    
    //選一個範圍
    xlsFile& SelectCell(const char* , const char* );    
    xlsFile& SelectCell(const char* , int ,const char* , int );
    xlsFile& SelectCell(char,int,char,int);    
    xlsFile& SelectCell(char,char,int,char,char,int);    
    //--------------------------------------------
    void ClearCell();                                //清除儲存格
    xlsFile& SetMergeCells(short vMerge = TRUE, 
                    bool isCenterAcross = true);     //合併儲存格（通常會配跨欄置中）
    //--------------------------------------------
    //對齊
    xlsFile& SetHorztlAlgmet(short);    //水平對齊
    xlsFile& SetVrticlAlgmet(short);    //垂直對齊
    xlsFile& SetTextAngle(short Angle); //方向-文字角度
    xlsFile& AutoNewLine(bool NewLine); //自動換行
    //---------------------------------------------
    //格線
    xlsFile& SetCellBorder(long BoarderStyle = 1, 
        int BoarderWeight = 2, long BoarderColor = 1);  //設定框線粗細和顏色
    //---------------------------------------
    //儲存格大小
    void AutoFitHight();            //自動調整列高
    void AutoFitWidth();            //自動調整欄寬
    xlsFile& SetCellHeight(float);    //設定列高
    xlsFile& SetCellWidth(float);    //設定欄寬
    //---------------------------------------------
    //字
    xlsFile& SetFont(const char* fontType = "新細明體");    //設定字型（預設新細明體）
    xlsFile& SetFontBold(bool isBold = true);               //粗體
    xlsFile& SetFontStrkthrgh(bool isBold = true);          //刪除線
    xlsFile& SetFontSize(short fontSize = 12);              //設定字體大小（預設12pt）
    xlsFile& SetFontColor(short colorIndex = 1);            //字型顏色（預設黑色）
    //---------------------------------------------
    xlsFile& SetCellColor(short);//設定底色
    //---------------------------------------------
    //（17-32隱藏版也有收進來）
    //Microsoft Excel 的顏色排序是依
    //紅、橙、黃、綠、藍、靛、紫、灰（y），由深到淺（x）
    //不過絕對RGB並沒有規律的存在這個表裡
    short SelectColor(short x = 8, short y = 7);  //依excel介面的座標選擇顏色
    short SelectColor(const char ColorChar = 'W');    //快速版（黑D、白W、紅R、綠G、藍B、黃Y）
    //---------------------------------------------
    //設定資料進儲存格（存成字串）
    //一般版
    void SetCell(int);
    void SetCell(double);
    void SetCell(long);    
    void SetCell(const char* );    
    void SetCell(CString );    
    //自訂細部格式版
    void SetCell(const char*, int);
    void SetCell(const char*, double);
    void SetCell(const char*, long);
    //--------------------------------------------
    //取值
    int GetCell2Int();
    CString GetCell2CStr();
    double GetCell2Double();
    //--------------------------------------------
    //排序（依列排序）//edisonx
    void Sort(CString IndexCell1, long DeCrement1,
              CString IndexCell2 = "", long DeCrement2 = 1,
              CString IndexCell3 = "", long DeCrement3 = 1);
    //--------------------------------------------
    //皆由edisonx提供函數資料

    //儲存圖表圖片.bmp（.jpg亦可以）
    void SaveChart(CString FullBmpPathName);
    //圖表（三類型的函數在每次建立都要使用）
    //使用前必須選擇貼上Chart的儲存格範圍
    
    //選擇資料範圍
    xlsFile& SelectChartRange(const char* , const char* );    
    xlsFile& SelectChartRange(const char* , int ,const char* , int );
    xlsFile& SelectChartRange(char,int,char,int);    
    xlsFile& SelectChartRange(char,char,int,char,char,int);
    //設定Chart參數
    xlsFile& SetChart(short XaxisByToporLeft = 2, bool isLabelVisable = 1, 
        CString = "" , CString = "" , CString = "" );
    //區域、直方、方柱、圓柱、圓錐、金字塔
    void InsertHistogramChart(int shapeType = CT_COLUMN, bool is3D = 0, 
                          int isVrticlorHorztlorOther = 0, 
                          int isNone_Stack_Percent = 0);
    //其它（特殊圖表）   
    void InsertBubleChart(bool is3D = 0);                                                 //泡泡圖
    void InsertDoughnutChart(bool Explode = 0);                                           //圓環圖
    void InsertSurfaceChart(bool is3D = 0, bool isWire = 0);                              //曲面圖
    void InsertRadarChart(bool isWire = 0, bool isDot = 1);                               //雷達圖
    void InsertPieChart(bool Explode = 0, int type2Dor3DorOf = 0);                        //圓餅圖
    void InsertLineChart(bool isDot = 1, bool is3D = 0, int isNone_Stack_Percent = 0);    //折線圖
    void InsertXYScatterChart(bool isDot, bool isLine, bool Smooth);                      //離散圖
    void InsertStockChart(int);                                                           //股票圖
    //--------------------------------------------
    void InsertImage(const char* , float , float );  //插入圖片
    void InsertImage(const char* );                  //插入圖片（先選取範圍，圖檔必失真）

private:
    void xlsFile::NewChart();  //在Sheet新增圖表
    //防止任何運算
    void operator+(const xlsFile&);
    void operator-(const xlsFile&);
    void operator*(const xlsFile&);
    void operator/(const xlsFile&);
    void operator%(const xlsFile&);
    void operator=(const xlsFile&);
    
    bool operator<(const xlsFile&);
    bool operator>(const xlsFile&);
    bool operator>=(const xlsFile&);
    bool operator<=(const xlsFile&);
    bool operator==(const xlsFile&);
    bool operator!=(const xlsFile&);

    
    bool operator&&(const xlsFile&);
    bool operator&(const xlsFile&);
    bool operator||(const xlsFile&);
    bool operator|(const xlsFile&);
    
    bool operator>>(const xlsFile&);
    bool operator<<(const xlsFile&);
};
/*

  range.SetFormula(COleVariant("=RAND()*100000"));    //套公式
  range.setSetValue(COleVariant("Last Name"));        //輸入值
  range.SetNumberFormat(COleVariant("$0.00"));        //數字格式

  //插圖
  Shapes shapes = objSheet.GetShapes(); 
  range = objSheet.GetRange(COleVariant("J7"),COleVariant("R21")); 
  
    //range.AttachDispatch(pRange);
    shapes.AddPicture(
        "c:\\CHILIN.bmp",                //LPCTSTR Filename
        false,                            //long LinkToFile
        true,                            //long SaveWithDocument
        (float)range.GetLeft().dblVal,    //float Left
        (float)range.GetTop().dblVal,   //float Top
        (float)range.GetWidth().dblVal, //float Width
        (float)range.GetHeight().dblVal //float Height
    );
    range.Sort(
    key1,        //  key1
    DeCrement1,    // long Order1, [ 1(ascending order) or 2(descending order) ]
    key2,        // key2, 
    VOptional,    // type, [xlSortLabels, xlSortValues]
    DeCrement2,    // long Order2, [ 1(升冪) or 2( 降) ]
    key3,        // key3
    DeCrement3,    // long Order3, [ 1(升冪) or 2( 降) ]
    2,            // Header, [0,1 : 不含 title 2 : title （選取範圍）一起排
    //進階
    VOptional,                    // OrderCustom [從1開始，自定義排序順序列表中之索引號，省略使用常規]
    _variant_t((short)TRUE),    // MatchCase [TRUE分大小寫排]
    1,                            // Orientation : [排序方向, 1:按列, 2:按行)
    1,                            // SortMethod : [1:按字符漢語拼音順序, 2:按字符筆畫數]
    //未知選項
    1, // DataOption1 可選 0 與 1
    1, // DataOption2 可選 0 與 1
    1  // DataOption3 可選 0 與 1
        );
*/
#endif

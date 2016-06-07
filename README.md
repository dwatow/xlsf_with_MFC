# Class-of-xls-File
A Class of xls funtion file with C++(MFC)

# 聲明
這是一份學MFC沒多久，而寫出來的程式碼。

最大目的希望可以讓目前使用MFC的人，可以更愉快一點！^^"

希望這個Class可以拋磚引玉，引起大家對於MFC使用心得的話題，討論聊MFC的使用心得。

# 作用
從MFC加入Excel的type library的function太多太亂太難用

所以建一個class將function包成一些常使用的動作

# 使用平台
* OS: Windows 7
* Tool: Visual C++ 6
* Type Library: Excel 2000、2003（其他版本沒有測試過）

# 檔案說明
* Source Code
  * xlef.h 宣告檔
  * xlef.cpp 定義檔
* Reference Doc
  * Chart Style.ppt 
    * 圖表的參考資料: 原本只是要試著分類，結果還沒有做出來(而且超複雜)
  * Chart and Cell Setup.xls
    * sheet圖表: 
      * UI圖表列表與圖表代碼的關係
    * sheet底色: 
      * 調色盤色碼
      * 用excel UI的座標, 取代原本api的數字
      * 用 黑D、白W、紅R、綠G、藍B、黃Y
    * sheet框線: 
      * Boarder_Style(左)
      * Boarder_Weight(右)
  * cexcel.h  
    MFC生成檔, 可以拿來比對你的檔案和我的檔案差在哪, 會影響xlef的支援度。

# 使用前...要做的事
1. MFC的.exe專案
2. 加入Excel type library（EXCEL11）

## 參考資料
* [\~流浪小築\~](http://www.intra.idv.tw/)的[自動產生Excel]( http://www.intra.idv.tw/data/c_school_4/mfc/auto_excel.htm)

# Sample Code
用起來的code會像這樣
~~~
#include "xlsFile.h"

xlsFile CrossTalkFrom;

CrossTalkFrom.New();  //開新檔案
CrossTalkFrom.SetSheetName(1, "CrossTalk值");  //新增sheet, 排在第1位(1起始), 命名為 CrossTalk值

//////////////////////////////////////////////////////////////////////////
//填值
//SelectCell: 選擇儲存格, 填入座標
//SetCell: 填入值

CrossTalkFrom.SelectCell("B1").SetCell("Lv");
CrossTalkFrom.SelectCell("F1").SetCell("x");
CrossTalkFrom.SelectCell("J1").SetCell("y");
CrossTalkFrom.SelectCell("C2").SetCell(vChain1[0].GetStrLv());
CrossTalkFrom.SelectCell("B3").SetCell(vChain1[1].GetStrLv());
CrossTalkFrom.SelectCell("D3").SetCell(vChain1[2].GetStrLv());
CrossTalkFrom.SelectCell("C4").SetCell(vChain1[3].GetStrLv());
CrossTalkFrom.SelectCell("C6").SetCell(vChain1[4].GetStrLv());
CrossTalkFrom.SelectCell("B7").SetCell(vChain1[5].GetStrLv());
CrossTalkFrom.SelectCell("D7").SetCell(vChain1[6].GetStrLv());
CrossTalkFrom.SelectCell("C8").SetCell(vChain1[7].GetStrLv());
CrossTalkFrom.SelectCell("C10").SetCell(vChain1[8].GetStrLv());
CrossTalkFrom.SelectCell("B11").SetCell(vChain1[9].GetStrLv());
CrossTalkFrom.SelectCell("D11").SetCell(vChain1[10].GetStrLv());
CrossTalkFrom.SelectCell("C12").SetCell(vChain1[11].GetStrLv());
CrossTalkFrom.SelectCell("G2").SetCell(vChain1[0].GetStrSx());
CrossTalkFrom.SelectCell("F3").SetCell(vChain1[1].GetStrSx());
CrossTalkFrom.SelectCell("H3").SetCell(vChain1[2].GetStrSx());
CrossTalkFrom.SelectCell("G4").SetCell(vChain1[3].GetStrSx());
CrossTalkFrom.SelectCell("G6").SetCell(vChain1[4].GetStrSx());
CrossTalkFrom.SelectCell("F7").SetCell(vChain1[5].GetStrSx());
CrossTalkFrom.SelectCell("H7").SetCell(vChain1[6].GetStrSx());
CrossTalkFrom.SelectCell("G8").SetCell(vChain1[7].GetStrSx());
CrossTalkFrom.SelectCell("G10").SetCell(vChain1[8].GetStrSx());
CrossTalkFrom.SelectCell("F11").SetCell(vChain1[9].GetStrSx());
CrossTalkFrom.SelectCell("H11").SetCell(vChain1[10].GetStrSx());
CrossTalkFrom.SelectCell("G12").SetCell(vChain1[11].GetStrSx());
CrossTalkFrom.SelectCell("K2").SetCell(vChain1[0].GetStrSy());
CrossTalkFrom.SelectCell("J3").SetCell(vChain1[1].GetStrSy());
CrossTalkFrom.SelectCell("L3").SetCell(vChain1[2].GetStrSy());
CrossTalkFrom.SelectCell("K4").SetCell(vChain1[3].GetStrSy());
CrossTalkFrom.SelectCell("K6").SetCell(vChain1[4].GetStrSy());
CrossTalkFrom.SelectCell("J7").SetCell(vChain1[5].GetStrSy());
CrossTalkFrom.SelectCell("L7").SetCell(vChain1[6].GetStrSy());
CrossTalkFrom.SelectCell("K8").SetCell(vChain1[7].GetStrSy());
CrossTalkFrom.SelectCell("K10").SetCell(vChain1[8].GetStrSy());
CrossTalkFrom.SelectCell("J11").SetCell(vChain1[9].GetStrSy());
CrossTalkFrom.SelectCell("L11").SetCell(vChain1[10].GetStrSy());
CrossTalkFrom.SelectCell("K12").SetCell(vChain1[11].GetStrSy());

//////////////////////////////////////////////////////////////////////////
//畫背景和框線
//SelectCell: 選取儲存格範圍, 英文和數字可以分開填入，並且參數化
//SetCellColor: 設定底色
//SetCellBorder: 設定框(填預設值)

char cCell = 'B';
int iCell = 2;
CrossTalkFrom.SelectCell(cCell, iCell   , cCell + 2, iCell + 2).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell, iCell +4, cCell + 2, iCell + 6).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell + 1, iCell +5).SetCellColor(2);
CrossTalkFrom.SelectCell(cCell, iCell +8, cCell + 2, iCell +10).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell + 1, iCell +9).SetCellColor(1);

cCell = 'F';
CrossTalkFrom.SelectCell(cCell, iCell   , cCell + 2, iCell + 2).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell, iCell +4, cCell + 2, iCell + 6).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell + 1, iCell +5).SetCellColor(2);
CrossTalkFrom.SelectCell(cCell, iCell +8, cCell + 2, iCell +10).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell + 1, iCell +9).SetCellColor(1);

cCell = 'J';
CrossTalkFrom.SelectCell(cCell, iCell   , cCell + 2, iCell + 2).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell, iCell +4, cCell + 2, iCell + 6).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell + 1, iCell +5).SetCellColor(2);
CrossTalkFrom.SelectCell(cCell, iCell +8, cCell + 2, iCell +10).SetCellColor(15).SetCellBorder();
CrossTalkFrom.SelectCell(cCell + 1, iCell +9).SetCellColor(1);

//顯示 操控權還給使用者
CrossTalkFrom.SetVisible(true);
~~~

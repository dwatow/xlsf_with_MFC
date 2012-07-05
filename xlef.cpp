#include "stdafx.h"
#include "xlef.h"
#include <comdef.h>

xlsFile::xlsFile(): 
VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR), VFALSE((short)FALSE), VTRUE((short)TRUE)
{
	ZeroMemory(buf,sizeof(buf));
	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	//Step 1.叫Excel應用程式
	if(!objApp.CreateDispatch("Excel.Application",&e))
	{
		CString str;
		str.Format("Excel CreateDispatch() failed w/err 0x%08lx", e.m_sc);
		AfxMessageBox(str, MB_SETFOREGROUND);
	}
};

xlsFile::~xlsFile()
{
	//objApp.SetUserControl(TRUE);  //移至visualable
	range.ReleaseDispatch();
	//chartobject.ReleaseDispatch();
	//chartobjects.ReleaseDispatch();
	objSheet.ReleaseDispatch();
	objSheets.ReleaseDispatch();
	objBook.ReleaseDispatch();
	objBooks.ReleaseDispatch();
	objApp.ReleaseDispatch();
}

//Open()
xlsFile& xlsFile::New()
{
	objBooks = objApp.GetWorkbooks();
    objBook = objBooks.Add(VOptional);	//開新檔案
	objSheets = objBook.GetWorksheets();
	return *this;
}

xlsFile& xlsFile::Open(const char* path)
{
	objBooks = objApp.GetWorkbooks();
    objBook.AttachDispatch(objBooks.Add(_variant_t(path))); //開啟一個已存在的檔案
	objBook.Activate();
	objSheets = objBook.GetWorksheets();
	return *this;
}

void xlsFile::SaveAs(const char* strTableName)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf, "%s", strTableName);
	objBook.SaveAs(
		COleVariant(buf),
		VOptional, VOptional, 
		VOptional, VOptional, 
		VOptional, 1,
		VOptional, VFALSE,
		VOptional, VOptional, VOptional); 
	//objBook.Close (VOptional,COleVariant(buf),VOptional);
}

void xlsFile::Save()
{
	objBook.Save();
}

/*
void xlsFile::Quit(CString FileName)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%s",FileName);
	objBook.Close(VFalse,COleVariant(buf), VOptional); //關閉不跳出視窗問問題
	objBooks.Close();
	objApp.Quit();
}
*/
void xlsFile::Quit()
{
	objBook.Close(VFALSE,VOptional, VOptional);
	objBooks.Close();
	objApp.Quit();
}

//SetVisible()
void xlsFile::SetVisible(bool a)
{
	objApp.SetVisible(a);    //顯示Excel檔
	objApp.SetUserControl(a);//使用者控制後，就不可以自動關閉
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Sheet操作

//-------------------------
////取得 Sheet 個數
long xlsFile::SheetTotal()
{
	return objSheets.GetCount();//edisonx
}
//-------------------------
//由SheetNumber 指定SheetName
void xlsFile::SetSheetName(short SheetNumber, const char* SheetName)
{
	objSheet = objSheets.GetItem(COleVariant(SheetNumber));
    objSheet.SetName(SheetName);//設定sheet名稱
}
//-------------------------
//由SheetNumber 取得SheetName
CString xlsFile::GetSheetName(short SheetNumber)
{
	objSheet = objSheets.GetItem(COleVariant(SheetNumber));
	return objSheet.GetName();//edisonx
}
//-------------------------
//選擇Sheet
//由SheetName
xlsFile& xlsFile::SelectSheet(const char* SheetName)
{
	objSheet = objSheets.GetItem(_variant_t(SheetName));
	objSheet.Activate();//edisonx
	return *this;
}
//由SheetNumber 
xlsFile& xlsFile::SelectSheet(short SheetNumber)
{
	objSheet = objSheets.GetItem(COleVariant(SheetNumber));
	objSheet.Activate();//edisonx
	return *this;
}
//-------------------------
//複製SheetName 指定插入Sheet的位置，並指定新Sheet名稱
void xlsFile::CopySheet(const char* SheetName)
{
	objSheet.AttachDispatch(objSheets.GetItem(_variant_t(SheetName)),true);
	objSheet.Copy(vtMissing,_variant_t(objSheet));
}
//複製SheetNumber 指定插入Sheet的位置，並指定名稱
void xlsFile::CopySheet(short SheetNumber)
{
	objSheet.AttachDispatch(objSheets.GetItem(COleVariant(SheetNumber)));
	objSheet.Copy(vtMissing,_variant_t(objSheet));
}
//-------------------------
//刪除Sheet
//選SheetName 
void xlsFile::DelSheet(const char* SheetName)
{	
	objSheet = objSheets.GetItem(_variant_t(SheetName));
	objSheet.Delete();//edisonx
}
//選SheetNumber
void xlsFile::DelSheet(short SheetNumber)
{
	objSheet = objSheets.GetItem(COleVariant(SheetNumber));
	objSheet.Delete();//edisonx
}
///////////////////////////////////////////////////////////////////////////////////////////
//Cell操作
//Cell計數計算
// 取得起始列
long xlsFile::GetHorztlStartCell()
{
	Range usedrange;
	usedrange.AttachDispatch(objSheet.GetUsedRange());
	return usedrange.GetColumn();
}
// 取得起始行
long xlsFile::GetVrticlStartCell()
{	
	Range usedrange;
	usedrange.AttachDispatch(objSheet.GetUsedRange());
	return usedrange.GetRow();
}
// 取得總列數
long xlsFile::GetHorztlTotalCell()
{
	Range usedrange;
	usedrange.AttachDispatch(objSheet.GetUsedRange());
	range.AttachDispatch(usedrange.GetColumns());
	return range.GetCount();
}
// 取得總行數
long xlsFile::GetVrticlTotalCell()
{
	Range usedrange;
	usedrange.AttachDispatch(objSheet.GetUsedRange());
	range.AttachDispatch(usedrange.GetRows());
	return range.GetCount();
}
//-------------------------
//Cell格式設定
//-------------------------
//選格子
//選一格
xlsFile& xlsFile::SelectCell(const char* x)
{
	range=objSheet.GetRange(COleVariant(x),COleVariant(x));
	return *this;
}

xlsFile& xlsFile::SelectCell(const char* x, int y)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%s%d",x,y);
	range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
	return *this;
}
//小於Z
xlsFile& xlsFile::SelectCell(char x, int y)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%c%d",x,y);
	range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
	return *this;
}
//大於Z，開始選AA
xlsFile& xlsFile::SelectCell(char x1,char x2,int y)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%c%c%d",x1,x2,y);
	range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
	return *this;
}
//-------------------------
//選格子
//選範圍
xlsFile& xlsFile::SelectCell(const char* x1,const char* x2)
{
	range=objSheet.GetRange(COleVariant(x1),COleVariant(x2));
	return *this;
}
xlsFile& xlsFile::SelectCell(const char* x1, int y1, const char* x2, int y2)
{
	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	sprintf(buf1,"%s%d",x1,y1);
	sprintf(buf2,"%s%d",x2,y2);
	range=objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
	return *this;
}
//小於Z
xlsFile& xlsFile::SelectCell(char x1, int y1, char x2, int y2)
{
	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	sprintf(buf1,"%c%d",x1,y1);
	sprintf(buf2,"%c%d",x2,y2);
	range=objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
	return *this;
}
//大於Z，開始選AA
xlsFile& xlsFile::SelectCell(char xA1, char xB1, int y1, char xA2, char xB2, int y2)
{
	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	sprintf(buf1,"%c%c%d",xA1,xB1,y1);
	sprintf(buf2,"%c%c%d",xA2,xB2,y2);

	range=objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
	return *this;
}
//-------------------------
//-------------------------
//清除儲存格
void xlsFile::ClearCell()
{
	//先選取一個範圍的儲存格
	range.Clear();//edisonx
}
//合併儲存格
xlsFile& xlsFile::SetMergeCells(short vMerge, bool isCenterAcross)
{
	//先選取一個範圍的儲存格
    range.SetMergeCells(_variant_t(vMerge));
	if(isCenterAcross) SetHorztlAlgmet(HA_CENTERACROSS);
	return *this;
}
//-------------------------
//-------------------------
//對齊方式
//水平對齊
xlsFile& xlsFile::SetHorztlAlgmet(short position)
{
	range.SetHorizontalAlignment(COleVariant(position));
	return *this;
}

//垂直對齊
xlsFile& xlsFile::SetVrticlAlgmet(short position)
{
	range.SetVerticalAlignment(COleVariant(position));
	return *this;
}

//對齊方式的方向幾度（+90~-90）
xlsFile& xlsFile::SetTextAngle(short Angle)
{
	range.SetOrientation(COleVariant(Angle)); 
	return *this;
}
//設定文字自動換行
xlsFile& xlsFile::AutoNewLine(bool NewLine)
{
	if(NewLine)		range.SetWrapText(VTRUE);
	else			range.SetWrapText(VFALSE);
	return *this;
}
//-------------------------
//-------------------------
//設定框線、框線顏色
xlsFile& xlsFile::SetCellBorder(long BoarderStyle, int BoarderWeight, long BoarderColor)
{
	range.BorderAround(_variant_t(BoarderStyle), BoarderWeight, BoarderColor,_variant_t((long)RGB(0,0,0)));
	return *this;
}
//-------------------------
//-------------------------
//設定欄寬列高
//自動調整列高
void xlsFile::AutoFitWidth()
{
	col = range.GetEntireColumn();	//選取某個範圍的一整排
	col.AutoFit();					//自動調整一整排的欄寬
}
//自動調整欄寬
void xlsFile::AutoFitHight()
{
	row = range.GetEntireRow();		//選取某個範圍的一整排
	row.AutoFit();					//自動調整一整排的列高
}
//設定列高
xlsFile& xlsFile::SetCellHeight(float height)
{
	range.SetRowHeight(_variant_t(height));
	return *this;
}
//設定欄寬
xlsFile& xlsFile::SetCellWidth(float height)
{
	range.SetColumnWidth(_variant_t(height));
	return *this;
}
//-------------------------
//-------------------------
//設定字型
xlsFile& xlsFile::SetFont(const char* fontType)
{
	font = range.GetFont();
    font.SetName(_variant_t(fontType));//原本是韓文字型
	return *this;
}
//粗體
xlsFile& xlsFile::SetFontBold(bool isBold)
{
	font = range.GetFont();
	if (isBold)		font.SetBold(VTRUE);
	else			font.SetBold(VFALSE);
	//font.SetBold(_variant_t(isBold)); //粗體
	return *this;
}
//刪除線
xlsFile& xlsFile::SetFontStrkthrgh(bool isStrike)
{
	font = range.GetFont();
	if (isStrike)	font.SetStrikethrough(VTRUE);	//edisonx
	else			font.SetStrikethrough(VFALSE);	//edisonx
	//font.SetStrikethrough(_variant_t((short)STRIKE));
	return *this;
}
//字型大小
xlsFile& xlsFile::SetFontSize(short fontSize)
{
	font = range.GetFont();
    font.SetSize(_variant_t(fontSize));//字型大小pt
	return *this;
}
//字型顏色
xlsFile& xlsFile::SetFontColor(short colorIndex)
{
	font = range.GetFont();
	font.SetColorIndex(_variant_t(colorIndex)); //字色(預設黑色)
	return *this;
}
//-------------------------
//-------------------------
//設定底色
xlsFile& xlsFile::SetCellColor(short colorIndex)
{
	cell = range.GetInterior();                   //取得選取範圍，設定儲存格的記憶體位址
    cell.SetColorIndex(_variant_t(colorIndex));   //設定底色（查表）
	//cell.SetColor(_variant_t(colorIndex));
	return *this;
}
//選擇顏色（適合字色和底色）依excel介面的座標選擇顏色
short xlsFile::SelectColor(short x, short y)
{
//Microsoft Excel 的顏色排序是依
//紅、橙、黃、綠、藍、靛、紫、灰（y）
//由深到淺（x）
	switch(x)
	{
	case 1:
			 if(y == 1)	return 1;
		else if(y == 2) return 9;
		else if(y == 3) return 3;
		else if(y == 4) return 7;
		else if(y == 5) return 38;

		else if(y == 6) return 17;
		else if(y == 7) return 38;
		break;
	case 2:
			 if(y == 1)	return 53;
		else if(y == 2) return 46;
		else if(y == 3) return 45;
		else if(y == 4) return 44;
		else if(y == 5) return 40;
		
		else if(y == 6) return 18;
		else if(y == 7) return 26;
		break;
	case 3:
			 if(y == 1)	return 52;
		else if(y == 2) return 12;
		else if(y == 3) return 43;
		else if(y == 4) return  6;
		else if(y == 5) return 36;
		
		else if(y == 6) return 19;
		else if(y == 7) return 27;
		break;
	case 4:
			 if(y == 1)	return 51;
		else if(y == 2) return 10;
		else if(y == 3) return 50;
		else if(y == 4) return  4;
		else if(y == 5) return 35;
		
		else if(y == 6) return 20;
		else if(y == 7) return 28;
		break;
	case 5:
			 if(y == 1)	return 49;
		else if(y == 2) return 14;
		else if(y == 3) return 42;
		else if(y == 4) return  8;
		else if(y == 5) return 34;
		
		else if(y == 6) return 21;
		else if(y == 7) return 29;
		break;
	case 6:
			 if(y == 1)	return 11;
		else if(y == 2) return  5;
		else if(y == 3) return 41;
		else if(y == 4) return 33;
		else if(y == 5) return 37;
		
		else if(y == 6) return 22;
		else if(y == 7) return 30;
		break;
	case 7:
			 if(y == 1)	return 55;
		else if(y == 2) return 47;
		else if(y == 3) return 13;
		else if(y == 4) return 54;
		else if(y == 5) return 39;
		
		else if(y == 6) return 23;
		else if(y == 7) return 31;
		break;
	case 8:
			 if(y == 1)	return 56;
		else if(y == 2) return 16;
		else if(y == 3) return 48;
		else if(y == 4) return 15;
		else if(y == 5) return  2;
		
		else if(y == 6) return 24;
		else if(y == 7) return 32;
		break;
	}
	return 2;//預設白色
}
short xlsFile::SelectColor(const char ColorChar)
{
	switch(ColorChar)
	{
	//黑色
	case 'D':
	case 'd':
		return 1;
		break;
	//白色
	case 'W':
	case 'w':
		return 2;
		break;
	//紅色
	case 'R':
	case 'r':
		return 3;
		break;
	//綠色
	case 'G':
	case 'g':
		return 4;
		break;
	//藍色
	case 'B':
	case 'b':
		return 5;
		break;
	//黃色
	case 'Y':
	case 'y':
		return 6;
		break;	}
	return 2;//預設白色
}
///////////////////////////////////////////////////////////////////////////////////////////
//Cell操作
//Cell填值
//-------------------------
//-------------------------
//SetCell()
void xlsFile::SetCell(int Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%d",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(long Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%d",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(double Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%f",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Data)
{
	ZeroMemory(buf,sizeof(buf));
	strcpy(buf,Data);
	//sprintf(buf,"%s",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(CString Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%s",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, int Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, double Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, long Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}
//-------------------------
//-------------------------
CString xlsFile::GetCell2CStr()
{
// 	COleVariant vResult = range.GetValue2();//edisonx 
// 	vResult.ChangeType(VT_BSTR);			//edisonx 
// 	return vResult.bstrVal;					//edisonx 

	//      VARIANT cellvalue;
	//      cellvalue = ;
    //cellvalue = range.GetText();
    return (char*)_bstr_t(range.GetItem(_variant_t((long)1), _variant_t((long)1)));
}
int xlsFile::GetCell2Int()
{
	COleVariant vResult = range.GetValue2();//edisonx 
	vResult.ChangeType(VT_INT);				//edisonx 
	return vResult.intVal;					//edisonx 
}
double xlsFile::GetCell2Double()
{
	COleVariant vResult = range.GetValue2();//edisonx 
	vResult.ChangeType(VT_R8);				//edisonx 
	return vResult.dblVal;					//edisonx 
}
///////////////////////////////////////////////////////////////////////////////////////////
//演算法操作
//排序
void xlsFile::Sort(CString IndexCell1, long DeCrement1,
				   CString IndexCell2, long DeCrement2,
				   CString IndexCell3, long DeCrement3)
{
	VARIANT key1, key2, key3;

	V_VT(&key1) = VT_DISPATCH;
	V_DISPATCH(&key1)=objSheet.GetRange(COleVariant(IndexCell1),COleVariant(IndexCell1));

	if(IndexCell2.IsEmpty())
	{
		range.Sort( key1, DeCrement1, VOptional, VOptional, 1, VOptional, 1, 2,//一般選項
					VOptional, _variant_t((short)TRUE),//進階 
					1, 1, 1, 1, 1);//未知選項//edisonx
	}
	else
	{
		V_VT(&key2) = VT_DISPATCH;
		V_DISPATCH(&key2)=objSheet.GetRange(COleVariant(IndexCell2),COleVariant(IndexCell2));
		
		if(IndexCell3.IsEmpty())
		{
			range.Sort( key1, DeCrement1, key2,	VOptional, DeCrement2, VOptional, 1, 2,
						VOptional, _variant_t((short)TRUE),//進階 
						1, 1, 1, 1, 1);//未知選項//edisonx
		}
		else
		{
			V_VT(&key3) = VT_DISPATCH;
			V_DISPATCH(&key3)=objSheet.GetRange(COleVariant(IndexCell3),COleVariant(IndexCell3));
			
			range.Sort(	key1, DeCrement1, key2, VOptional, DeCrement2, key3, DeCrement3, 2,//一般選項
						VOptional, _variant_t((short)TRUE),//進階 
						1, 1, 1, 1, 1);//未知選項//edisonx
		}
	}
}
///////////////////////////////////////////////////////////////////////////////////////////
//圖表操作
//插入圖表（一條龍code）
/*
void xlsFile::DrawChart(CString DataRangeStart, CString DataRangeEnd, 
					   long ChartType, short PlotBy, 
					   short StartFrom, CString TitleString, 
					   UINT ChartStartX, UINT ChartStartY, UINT width, UINT height
					   ) // 畫表格
{	
	//在Sheet新增圖表
	lpDisp = objSheet.ChartObjects(VOptional);
	chartobjects.AttachDispatch(lpDisp);
	
	//圖表符合儲存格範圍的大小
	chartobject = chartobjects.Add( (float)range.GetLeft().dblVal,  (float)range.GetTop().dblVal, 
									(float)range.GetWidth().dblVal, (float)range.GetHeight().dblVal);
	//資料來源（一筆）
	xlsChart.AttachDispatch(chartobject.GetChart());
	lpDisp = objSheet.GetRange(COleVariant(DataRangeStart), COleVariant(DataRangeEnd));
	range.AttachDispatch(lpDisp);
	
	var.vt = VT_DISPATCH;
	var.pdispVal = lpDisp;

	xlsChart.ChartWizard(var,					// const VARIANT& Source.
		COleVariant((short)11),					// const VARIANT& fix please, Gallery: 3d Column. 1 or 11 是否轉動3D（3D類適用, 1轉，11不轉）
		COleVariant((short)1),					// const VARIANT& fix please, Format, use default
		COleVariant((short)PlotBy),				// const VARIANT& PlotBy: 1.X  2.Y 圖表的x軸要使用 表格的1:X-top還是2:Y-left
		COleVariant((short)1),					// const VARIANT& Category Labels fix please 不當軸的那個資料，從第幾個格子開始算（比較群組資料數量）
		COleVariant((short)StartFrom),			// const VARIANT& Series Labels. Start X, 不當軸的那個資料，資料名稱要用幾個格子（更改名字）
		COleVariant((short)TRUE),				// const VARIANT& HasLegend. 是否要顯示群組資料標籤
		//以下可不填
		_variant_t(COleVariant(TitleString)),	// const VARITNT& Title
		_variant_t(COleVariant(X_String)),		// const VARIANT& CategoryTitle
		_variant_t(COleVariant(Y_String)),		// const VARIANT& ValueTitle
		VOptional								// const VARIANT& ExtraTitle
		);
	xlsChart.SetChartType((long)ChartType);		
}
*/


void xlsFile::NewChart()
{
	//在Sheet新增圖表
	lpDisp = objSheet.ChartObjects(VOptional);
	chartobjects.AttachDispatch(lpDisp);	
	//圖表符合儲存格範圍的大小
	chartobject = chartobjects.Add(
		(float)range.GetLeft().dblVal,
		(float)range.GetTop().dblVal, 
		(float)range.GetWidth().dblVal, 
		(float)range.GetHeight().dblVal );
	//資料來源（範圍left, top預設為 比較Item和Group）
	xlsChart.AttachDispatch(chartobject.GetChart());
}
///////////////////////////////////////////////////////////////////////////////////////////
//圖表操作
//儲存圖表
//edisonx
void xlsFile::SaveChart(CString FullBmpPathName)
{
	xlsChart.Export(LPCTSTR(FullBmpPathName),VOptional,VOptional);
}
//選擇表格資料的範圍
xlsFile& xlsFile::SelectChartRange(const char* x1,const char* x2)
{
	NewChart();

	lpDisp = objSheet.GetRange(COleVariant(x1),COleVariant(x2));
	range.AttachDispatch(lpDisp);
	
	return *this;
}

xlsFile& xlsFile::SelectChartRange(const char* x1, int y1, const char* x2, int y2)
{
	NewChart();	

	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	sprintf(buf1,"%s%d",x1,y1);
	sprintf(buf2,"%s%d",x2,y2);

	lpDisp = objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
	range.AttachDispatch(lpDisp);
	return *this;
}
//小於Z
xlsFile& xlsFile::SelectChartRange(char x1, int y1, char x2, int y2)
{
	NewChart();

	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	sprintf(buf1,"%c%d",x1,y2);
	sprintf(buf2,"%c%d",x1,y2);

	lpDisp = objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
	range.AttachDispatch(lpDisp);
	return *this;
}
//大於Z，開始選AA
xlsFile& xlsFile::SelectChartRange(char xA1, char xB1, int y1, char xA2, char xB2, int y2)
{
	NewChart();
	ZeroMemory(buf1,sizeof(buf1));
	ZeroMemory(buf2,sizeof(buf2));
	sprintf(buf1,"%c%c%d",xA1,xB1,y1);
	sprintf(buf2,"%c%c%d",xA2,xB2,y2);	
	lpDisp = objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
	range.AttachDispatch(lpDisp);
	return *this;
}

// 設定表格參數（預設會顯示立體直方圖）
xlsFile& xlsFile::SetChart(short XaxisByToporLeft, bool isLabelVisable, CString TitleString, CString XaxisTitle, CString YaxisTitle) 
{	
	var.vt = VT_DISPATCH;
	var.pdispVal = lpDisp;

	short LabelVisable = (isLabelVisable) ? TRUE : FALSE ;
		
	xlsChart.ChartWizard(var,					// const VARIANT& Source.
		COleVariant((short)11),					// const VARIANT& fix please, Gallery: 3d Column. 1 or 11 是否轉動3D（3D類適用, 1轉，11不轉）
		COleVariant((short)1),					// const VARIANT& fix please, Format, use default
		COleVariant(XaxisByToporLeft),			// const VARIANT& PlotBy: 1.X  2.Y 圖表的x軸要使用 表格的1:X-top還是2:Y-left
		COleVariant((short)1),					// const VARIANT& Category Labels fix please 不當軸的那個資料，從第幾個格子開始算（比較群組資料數量）
		COleVariant((short)1),					// const VARIANT& Series Labels. Start X, 不當軸的那個資料，資料名稱要用幾排格子（更改名字）
		COleVariant(LabelVisable),				// const VARIANT& HasLegend. 是否要顯示群組資料標籤
		//以下可不填
		_variant_t(COleVariant(TitleString)),	// const VARITNT& Title
		_variant_t(COleVariant(XaxisTitle)),	// const VARIANT& CategoryTitle
		_variant_t(COleVariant(YaxisTitle)),	// const VARIANT& ValueTitle
		VOptional								// const VARIANT& ExtraTitle
		);
	return *this;
}
//插入圖表
void xlsFile::InsertHistogramChart(int shapeType, bool is3D, 
						  int isVrticl_Horztl_Other, 
						  int isNone_Stack_Percent)
{
	long ChartType = 51;
	if (shapeType == 0)//Area
	{
		if(!is3D)//2D
		{
			if(isNone_Stack_Percent == 0)		 ChartType = 1;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 77;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 76;//有百分比
		}
		else		//3D
		{
			if(isNone_Stack_Percent == 0) 		 ChartType = -4098;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 78;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 79;//百分比
		}
	} 
	else if (shapeType == 1)//直方圖
	{
		if(isVrticl_Horztl_Other == 0)//直的
		{
			if(!is3D)//2D
			{
				if(isNone_Stack_Percent == 0) 	   ChartType = 51;//無堆疊
				else if (isNone_Stack_Percent == 1) ChartType = 52;//有堆疊
				else if (isNone_Stack_Percent == 2) ChartType = 53;//有百分比
			}
			else		//3D
			{
				if(isNone_Stack_Percent == 0)		ChartType = 54;//無堆疊
				else if (isNone_Stack_Percent == 1) ChartType = 55;//有堆疊
				else if (isNone_Stack_Percent == 2) ChartType = 56;//百分比
			}
		}
		else if(isVrticl_Horztl_Other == 1)//橫的
		{
			if(!is3D)//2D
			{
				if(isNone_Stack_Percent == 0)		ChartType = 57;
				else if (isNone_Stack_Percent == 1) ChartType = 58;
				else if (isNone_Stack_Percent == 2) ChartType = 59;
			}
			else		//3D
			{
				if(isNone_Stack_Percent == 0)		ChartType = 60;
				else if (isNone_Stack_Percent == 1) ChartType = 61;
				else if (isNone_Stack_Percent == 2) ChartType = 62;
			}
		}
		else						ChartType = -4100;	//平面 必3D
	}
	else if (shapeType == 2)//CONE
	{
		if(isVrticl_Horztl_Other == 0)//直的
		{
			if(isNone_Stack_Percent == 0)		ChartType = 92;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 93;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 94;//百分比
		}
		else if(isVrticl_Horztl_Other == 1)//橫的
		{
			if(isNone_Stack_Percent == 0)		ChartType = 95;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 96;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 97;//百分比
		}
		else									ChartType = 98;//平面 必3D
	}
	else if (shapeType == 3)
	{
		if(isVrticl_Horztl_Other == 0)//直的
		{
			if(isNone_Stack_Percent == 0)		ChartType = 99;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 100;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 101;//百分比
		}
		else if(isVrticl_Horztl_Other == 1)//橫的
		{
			if(isNone_Stack_Percent == 0)		ChartType = 102;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 103;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 104;//百分比
		}
		else					ChartType = 105;//平面 必3D
	}
	else if (shapeType == 4)
	{
		if(isVrticl_Horztl_Other == 0)//直的
		{
			if(isNone_Stack_Percent == 0)		ChartType = 106;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 107;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 108;//百分比
		}
		else if(isVrticl_Horztl_Other == 1)//橫的
		{
			if(isNone_Stack_Percent == 0)		ChartType = 109;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 110;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 111;//百分比
		}
		else					ChartType = 112;	//平面 必3D
	}

	xlsChart.SetChartType((long)ChartType);
}

//泡泡
void xlsFile::InsertBubleChart(bool is3D)
{
	long ChartType = 51;

		if(is3D)		ChartType = 15;
		else			ChartType = 87;

	xlsChart.SetChartType((long)ChartType);
}
//圓環
void xlsFile::InsertDoughnutChart(bool Explode)
{
	long ChartType = 51;
	
	if(!Explode)	ChartType = -4120;
	else			ChartType = 80;

	xlsChart.SetChartType((long)ChartType);

}
//曲面
void xlsFile::InsertSurfaceChart(bool is3D, bool isWire)
{
	long ChartType = 51;
	
	if (is3D)
	{
		if (!isWire)	ChartType = 83;
		else			ChartType = 84;
	} 
	else
	{
		if (!isWire)	ChartType = 85;
		else			ChartType = 86;
	}

	xlsChart.SetChartType((long)ChartType);
}
//雷達
void xlsFile::InsertRadarChart(bool isWire, bool isDot)
{
	long ChartType = 51;
	
	if (isWire)
	{
		if (!isDot)	ChartType = -4151;
		else		ChartType = 81;
	} 
	else			ChartType = 82;

	xlsChart.SetChartType((long)ChartType);
}
//圓餅
void xlsFile::InsertPieChart(bool Explode, int type2Dor3DorOf)
{
	long ChartType = 51;

	if(!Explode)
	{
		if (type2Dor3DorOf == 0)			ChartType = 5;
		else if (type2Dor3DorOf == 1)		ChartType = -1402;
		else if (type2Dor3DorOf == 2)		ChartType = 68;
	}
	else
	{
		if (type2Dor3DorOf == 0)			ChartType = 69;
		else if (type2Dor3DorOf == 1)		ChartType = 70;
		else if (type2Dor3DorOf == 2)		ChartType = 71;
	}
	
	xlsChart.SetChartType(ChartType);
}

void xlsFile::InsertLineChart(bool isDot, bool is3D, int isNone_Stack_Percent)
{
	long ChartType = 51;
	
	if(!is3D)//3D
	{
		if(!isDot)
		{
			if(isNone_Stack_Percent == 0) 	   ChartType = 4;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 63;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 64;//有百分比
		}
		else
		{
			if(isNone_Stack_Percent == 0) 	   ChartType = 65;//無堆疊
			else if (isNone_Stack_Percent == 1) ChartType = 66;//有堆疊
			else if (isNone_Stack_Percent == 2) ChartType = 67;//有百分比
		}
	}
	else						ChartType = -4101;//3D

	xlsChart.SetChartType((long)ChartType);
}

//離散圖
void xlsFile::InsertXYScatterChart(bool isDot, bool isLine, bool Smooth)
{
	long ChartType = 51;
	if(!isLine)			ChartType = -4169;//3D
	else
	{
		if(Smooth)
		{
			if(isDot) 	ChartType = 72;
			else	 	ChartType = 73;
		}
		else
		{
			if(isDot)	ChartType = 74;
			else		ChartType = 75;
		}
	}
	xlsChart.SetChartType((long)ChartType);
}

//股票圖
void xlsFile::InsertStockChart(int StockType)
{
	long ChartType = 51;
	
	if (StockType == 0)			ChartType = 88;
	else if (StockType == 1)	ChartType = 89;
	else if (StockType == 2)	ChartType = 90;
	else if (StockType == 3)	ChartType = 91;

	xlsChart.SetChartType((long)ChartType);
}
//--------------------------------------------
//--------------------------------------------
//插入圖（從檔案）
void xlsFile::InsertImage(const char* FileNamePath, float Width, float Height)
{
	shapes = objSheet.GetShapes(); 
	shapes.AddPicture(
		FileNamePath,					//LPCTSTR Filename
		false,							//long LinkToFile
		true,							//long SaveWithDocument
		(float)range.GetLeft().dblVal,	//float Left
		(float)range.GetTop().dblVal,   //float Top
		Width,							//float Width
		Height							//float Height
		);
}

void xlsFile::InsertImage(const char* FileNamePath)
{
	shapes = objSheet.GetShapes(); 
	shapes.AddPicture(
		FileNamePath,					//LPCTSTR Filename
		false,							//long LinkToFile
		true,							//long SaveWithDocument
		(float)range.GetLeft().dblVal,	//float Left
		(float)range.GetTop().dblVal,   //float Top
		(float)range.GetWidth().dblVal, //float Width
		(float)range.GetHeight().dblVal //float Height
		);
}

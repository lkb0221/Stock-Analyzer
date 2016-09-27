/*------------------------------------------------------------------------------*
 * File Name:				 													*
 * Creation: 																	*
 * Purpose: OriginC Source C file												*
 * Copyright (c) ABCD Corp.	2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010		*
 * All Rights Reserved															*
 * 																				*
 * Modification Log:															*
 *------------------------------------------------------------------------------*/
 
////////////////////////////////////////////////////////////////////////////////////
// Including the system header file Origin.h should be sufficient for most Origin
// applications and is recommended. Origin.h includes many of the most common system
// header files and is automatically pre-compiled when Origin runs the first time.
// Programs including Origin.h subsequently compile much more quickly as long as
// the size and number of other included header files is minimized. All NAG header
// files are now included in Origin.h and no longer need be separately included.
//
// Right-click on the line below and select 'Open "Origin.h"' to open the Origin.h
// system header file.
#include <Origin.h>
#include <GetNBox.h>
#include <..\OriginLab\OCHttpUtils.h>
////////////////////////////////////////////////////////////////////////////////////
#define STR_YAHOO_API_BASE "http://ichart.finance.yahoo.com/table.csv?s=%s&a=%d&b=%d&c=%d&d=%d&e=%d&f=%d&g=d&ignore=.csv"
#define STR_FILE_NAME_DOWNLOAD "Table.csv"
#define STR_TEMPLATE_NAME_WBK "Stock Analyzer.ogw"

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main

void StockAnalyzer_Main() {
	// Get Config Tree
	Tree tr;
	tr = StockAnalyzer_GUI();
	
	// Download File
	string strDownload = StockAnalyzer_Download(tr.Data.SymbolName.strVal, tr.Data.StartDate.dVal, tr.Data.EndDate.dVal);
	
	// Setup workbook
	StockAnalyzer_Setup(tr, strDownload);
	
	// Delete downloaded file
	StockAnalyzer_Delete(strDownload);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DIalog GUI

Tree StockAnalyzer_GUI() {
	double dToday = StockAnalyzer_GetToday();
	
	GETN_BOX(tr)
	
	GETN_BEGIN_BRANCH(Data, "Data")
	GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_RADIO_INDEX_EX(Method, "Data Source", 0, "Download through Yahoo API|Select from columns")
		GETN_STR(SymbolName, "Stock Name", "AAPL")
		GETN_RADIO_INDEX_EX(Range, "Range", 0, "Last Year|Last Three Year|All|Custom")
		GETN_DATE(StartDate, "Date From", dToday-366)
		GETN_DATE(EndDate, "Date To", dToday-1)
	GETN_END_BRANCH(Data)
	
	GETN_BEGIN_BRANCH(Overlays, "Technical Overlays")
		GETN_BEGIN_BRANCH(MA, "Moving Averages - Simple and Exponential")
			GETN_NUM(Period1, "1st Period", 20)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Period", 50)				GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "3rd Period", 100)				GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(MA)
		
		GETN_BEGIN_BRANCH(BB, "Bollinger Bands")
			GETN_NUM(Period, "Period Window (N)", 20)	GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Band, "Multiplier (K)", 2)					GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(BB)
		
	GETN_END_BRANCH(Overlays)
	
	GETN_BEGIN_BRANCH(Indicators, "Technical Indicators")
		//GETN_BEGIN_BRANCH(ADL, "Accumulation Distribution")
		//GETN_END_BRANCH(ADL)
		
		GETN_BEGIN_BRANCH(Arron, "Aroon")
			GETN_NUM(Period, "Look-back Period Window", 25)	GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(Arron)
		
	GETN_END_BRANCH(Indicators)
	
	
	
	if( GetNBox(tr, StockAnalyzer_GUI_event, "Stock Analyzer") ) {
		return tr;
	}
	
}

static double StockAnalyzer_GetToday() {
	double dToday;
	SYSTEMTIME sysTime;
	GetSystemTime(&sysTime);
	SystemTimeToJulianDate(&dToday, &sysTime);
	return dToday;
}

//////////////////////////////////////////////////////////////////////////////////////////////////
// GUI events

static int StockAnalyzer_GUI_event(TreeNode& tr, int nRow, int nEvent, DWORD& dwEnables, LPCSTR lpcszNodeName, WndContainer& getNContainer, string& strAux, string& strErrMsg) {
	StockAnalyzer_GUI_event_Input(tr.Data);
	return 0;
}

static bool StockAnalyzer_GUI_event_Input(Tree tr) {
	switch(tr.Method.nVal) {
	case 0:	// yahoo api
		break;
	case 1:	// column
		break;
	}
	
	return true;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Download csv form Yahoo API

static string StockAnalyzer_Download(string strSymbolName, double dStartDate, double dEndDate) {
	// Get day, month, year integers
	vector<int> vsStart(3), vsEnd(3);
	vsStart = StockAnalyzer_ParseDate(dStartDate);
	vsEnd = StockAnalyzer_ParseDate(dEndDate);
	
	// Build url string
	string strURL;
	strURL.Format(STR_YAHOO_API_BASE, strSymbolName, vsStart[1], vsStart[2], vsStart[0], vsEnd[1], vsEnd[2], vsEnd[0]);
	
	// Biuld Local File Path
	string strFilePath = StockAnalyzer_GetTemplatePath(STR_FILE_NAME_DOWNLOAD);
	
	// Download
	bool bGet = http_get_file(strURL, strFilePath);
	
	if (bGet == 0) {
		return strFilePath;
	}
	else {
		return NULL;
	}
}

static void StockAnalyzer_Delete(string strFilePath) {
	DeleteFile(strFilePath);
}

static vector<int> StockAnalyzer_ParseDate(double dDate/*, int& nYear, int& nMonth, int& nDay*/) {
	vector<int> vs(3);	// {Year, Month, Day}
	
	SYSTEMTIME st;
	JulianDateToSystemTime(&dDate, &st);
	vs[0] = st.wYear;	// year
	vs[1] = st.wMonth - 1;	// month
	vs[2] = st.wDay;		// day
	
	return vs;

}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Template Setup

static void StockAnalyzer_Setup(Tree tr, string strFilePath) {
	WorksheetPage wp;
	string strTemplatePath = StockAnalyzer_GetTemplatePath(STR_TEMPLATE_NAME_WBK);
	wp.Create(strTemplatePath);
	set_user_info(wp, "Overlays", tr.Overlays);
	set_user_info(wp, "Indicators", tr.Indicators);
	static Worksheet wksData = wp.Layers("Data");
	
	StockAnalyzer_Import(strFilePath, wksData);
	
	
}

static void StockAnalyzer_Import(string strFilePath, Worksheet wks) {
	ASCIMP	ai;
	if(AscImpReadFileStruct(strFilePath, &ai)==0) {
		ai.iAutoSubHeaderLines = 0;
		ai.iHeaderLines = 1;
		ai.iSubHeaderLines = 0;
		ai.iRenameWks = 0;
		if(0 == wks.ImportASCII(strFilePath, ai)) {
			wks.LT_execute("wks.col1.SetFormat(4, 22, yyyy-MM-dd);");
			wks.Sort();
			wks.LT_execute("run -p au;");
		};
	};
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// OverLays & Indicaters

void StockAnalyzer_MainProcess(int nUID) {
	// Get Configrations
	WorksheetPage wp;		wp = StockAnalyzer_UID2WB(nUID);
	Worksheet wksData = wp.Layers("Data");
	Worksheet wksOverLay = wp.Layers("Overlays");
	Worksheet wksIndicator = wp.Layers("Indicators");
	Tree trOverlays;		get_user_info(wp, "Overlays", trOverlays);
	Tree trIndicators;		get_user_info(wp, "Indicators", trIndicators);
	
	int nIndexOverlay, nIndexIndicator;
	vector<int> vsIndex;
	
	// SMA
	vsIndex = StockAnalyzer_Offset(nIndexOverlay, 3);	
	StockAnalyzer_SMA_Main(wksOverLay, wksData, vsIndex, trOverlays.MA);
	
	// EMA
	vsIndex = StockAnalyzer_Offset(nIndexOverlay, 3);	
	StockAnalyzer_EMA_Main(wksOverLay, wksData, vsIndex, trOverlays.MA);
	
	// Bollinger Bands
	vsIndex = StockAnalyzer_Offset(nIndexOverlay, 3);	
	StockAnalyzer_Bollinger_Main(wksOverLay, wksData, vsIndex, trOverlays.BB);
	
	// Money Flow, ADL & CMF
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_MoneyFlow_Main(wksIndicator, wksData, vsIndex);
	
	// Aroon
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_Aroon_Main(wksIndicator, wksData, vsIndex, trIndicators.Aroon);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// SMA

static void StockAnalyzer_SMA_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> vsIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	for (int ii = 0; ii < 3; ii++) {
		Column cc(wksTarget, vsIndex[ii]);
		Dataset ds(cc);
		int nWin = tree_get_node(tr, ii).nVal;
		
		if ((is_missing_value(nWin)) || (nWin <= 0)) {
			continue;
		}
		
		ds = StockAnalyzer_SMA(dsClose, nWin);
		cc.SetComments("SMA(" + ftoa(nWin) + ")");
	}
}

static vector<double> StockAnalyzer_SMA(vector<double> vsClose, int nWin) {
	int nNum = vsClose.TrimLeft(true);
	int nSize = vsClose.GetSize();
	vector<double> vsSMA(nSize);
	if (is_missing_value(nWin)) {
		return NULL;
	}
	
	for (int ii = 0; ii < nSize; ii++) {
		if (ii + 1 - nWin <= 0) {
			vsSMA[ii] = NANUM;
		}
		else {
			vector<double> vsSubRange;
			vsClose.GetSubVector(vsSubRange, ii-nWin, ii-1);
			double dSum;
			vsSubRange.Sum(dSum);
			vsSMA[ii] = dSum / nWin;
		};
	};
	
	if (nNum > 0)	vsSMA.InsertAt(0, NANUM, nNum);
	return vsSMA;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// EMA

static void StockAnalyzer_EMA_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> vsIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	for (int ii = 0; ii < 3; ii++) {
		Column cc(wksTarget, vsIndex[ii]);
		Dataset ds(cc);
		int nWin = tree_get_node(tr, ii).nVal;
		
		if ((is_missing_value(nWin)) || (nWin <= 0)) {
			continue;
		}
		
		ds = StockAnalyzer_EMA(dsClose, nWin);
		cc.SetComments("EMA(" + ftoa(nWin) + ")");
	}
}

static vector<double> StockAnalyzer_EMA(vector<double> vsClose, int nWin) {
	int nSize = vsClose.GetSize();
	vector<double> vsEMA(nSize), vsSMA(nSize);
	
	vsSMA = StockAnalyzer_SMA(vsClose, nWin);
	double dMultiplier = 2 / ((double)nWin + 1);
	
	// Calulate EMA
	bool bStart;
	for (int ii = 0; ii < nSize; ii++) {
		if (is_missing_value(vsSMA[ii])) {
			vsEMA[ii] = NANUM;
		}
		else if (!bStart){
			vsEMA[ii] = vsSMA[ii];
			bStart = true;
		}
		else {
			vsEMA[ii] = ((vsClose[ii] - vsEMA[ii-1]) * dMultiplier) + vsEMA[ii-1];
		}
	};
	
	return vsEMA;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Bollinger Bands

static void StockAnalyzer_Bollinger_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> vsIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	Column ccUpper(wksTarget, vsIndex[0]);			Dataset dsUpper(ccUpper);
	Column ccLower(wksTarget, vsIndex[1]);			Dataset dsLower(ccLower);
	Column ccCenter(wksTarget, vsIndex[2]);		Dataset dsCenter(ccCenter);
	int nPeriod = tr.GetNode("Period").nVal;
	int nWidth = tr.GetNode("Band").nVal;
	StockAnalyzer_Bollinger(dsClose, nPeriod, nWidth, dsCenter, dsUpper, dsLower);
	ccUpper.SetComments("Bollinger(" + ftoa(nPeriod) + ", " + ftoa(nWidth) + ")");
}

static bool StockAnalyzer_Bollinger(vector<double> vsClose, int nPeriod, int nWidth, vector<double> &vsCenter, vector<double> &vsUpper, vector<double> &vsLower) {
	int nSize = vsClose.GetSize();
	vsCenter.SetSize(nSize);
	vsUpper.SetSize(nSize);
	vsLower.SetSize(nSize);
	
	for (int ii = 0; ii < nSize; ii++) {
		if (ii + 1 - nPeriod <= 0) {
			vsCenter[ii] = NANUM;
			vsUpper[ii] = NANUM;
			vsLower[ii] = NANUM;
		}
		else {
			vector<double> vsSubRange;
			vsClose.GetSubVector(vsSubRange, ii-nPeriod, ii-1);
			double dMean, dSD;
			ocmath_basic_summary_stats(nPeriod, vsSubRange, NULL, &dMean, &dSD);
			vsCenter[ii] = dMean;
			vsUpper[ii] = dMean + nWidth * dSD;
			vsLower[ii] = dMean - nWidth * dSD;
		};
	};
	
	return true;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Money Flow ==> Accumulation Distribution Line, Chaikin Money Flow

static void StockAnalyzer_MoneyFlow_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex) {
	// Source Data
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// 
	Dataset dsVolume(wksSource, 5);
	vector<double> vsMoneyFlowVolume;
	
	// ADL
	Column ccADL(wksTarget, nIndex[0]);
	Dataset dsADL(ccADL);		dsADL = StockAnalyzer_ADL(dsHigh, dsLow, dsClose, dsVolume, vsMoneyFlowVolume);
	ccADL.SetComments("Accum/Dist");
	ccADL.SetLongName("Accumulation Distribution Line");
	
	// CMF
	Column ccCMF(wksTarget, nIndex[1]);
	int dCMFp = 20;
	Dataset dsCMF(ccCMF);		dsCMF = StockAnalyzer_CMF(vsMoneyFlowVolume, dsVolume, dCMFp);
	ccCMF.SetComments("CMF("+ ftoa(dCMFp) + ")");
	ccCMF.SetLongName("Chaikin Money Flow");
	
}

static vector<double> StockAnalyzer_ADL(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, vector<double> vsVolume, vector<double> &vsMoneyFlowVolume) {
	int nSize = vsClose.GetSize();
	vector<double> vsADL(nSize);		vsADL = NANUM;
	
	vector<double> vsMoneyFlowMultiplier(nSize);
	vsMoneyFlowVolume.SetSize(nSize);
	vsMoneyFlowMultiplier = ((vsClose-vsLow)-(vsHigh-vsClose)) / (vsHigh-vsLow);
	for (int ii = 0; ii < nSize; ii++) {
		vsMoneyFlowVolume[ii] = is_missing_value(vsMoneyFlowMultiplier[ii])? 0:vsMoneyFlowMultiplier[ii]*vsVolume[ii];
	};
	
	ocmath_d_cumulative_sum(vsMoneyFlowVolume, 0, nSize, vsADL);
	
	return vsADL;
}

static vector<double> StockAnalyzer_CMF(vector<double> vsMoneyFlowVolume, vector<double> vsVolume, int dCMFp = 20) {
	int nSize = vsMoneyFlowVolume.GetSize();
	vector<double> vsCMF(nSize);		vsCMF = NANUM;
	for (int ii = dCMFp-1; ii < nSize; ii++) {
		vector<double> vsSubF, vsSubV;
		vsMoneyFlowVolume.GetSubVector(vsSubF, ii - dCMFp + 1, ii);
		vsVolume.GetSubVector(vsSubV, ii - dCMFp + 1, ii);
		double dSumF, dSumV;
		vsSubF.Sum(dSumF);	vsSubV.Sum(dSumV);
		vsCMF[ii] = dSumF/dSumV;
	};
	return vsCMF;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Aroon

static void StockAnalyzer_Aroon_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	int nWin = tr.GetNode("Period").nVal;
	Column ccUp(wksTarget, nIndex[0]);		Dataset dsUp(ccUp);
	Column ccDown(wksTarget, nIndex[1]);	Dataset dsDown(ccDown);
	ccUp.SetLongName("Aroon");
	ccDown.SetLongName("Aroon");
	ccUp.SetComments("Arron Up");
	ccDown.SetComments("Arron Down");
	
	StockAnalyzer_Aroon(dsClose, dsUp, dsDown, nWin);
}

static void StockAnalyzer_Aroon(vector<double> vsClose, vector<double>& vsUp, vector<double>& vsDown, int nWin = 25) {
	int nSize = vsClose.GetSize();
	vsUp.SetSize(nSize);
	vsDown.SetSize(nSize);
	if (nSize <= nWin) {
		return;
	};
	
	for (int ii = nWin - 1; ii < nSize; ii++) {
		vector<double> vsSub;
		vsClose.GetSubVector(vsSub, ii - nWin + 1, ii);
		double dMin, dMax;
		uint nIndexMin, nIndexMax;
		vsSub.GetMinMax(dMin, dMax, &nIndexMin, &nIndexMax);
		vsUp[ii] = ((double)nIndexMax) / 25 * 100;
		vsDown[ii] = ((double)nIndexMin) / 25 * 100;
	};
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Misc

vector<int> StockAnalyzer_Offset(int &nIndex, int nSize) {
	vector<int> vs(nSize);
	for (int ii = 0; ii < nSize; ii++, nIndex++) {
		vs[ii] = nIndex + 1;
	};
	return vs;
}

static WorksheetPage StockAnalyzer_UID2WB(int nUID) {
	//Worksheet wks;				wks = (Worksheet)Project.GetObject(nUID);   
	//WorksheetPage wp;		wp = wks.GetPage();
	WorksheetPage wp;
	wp = (WorksheetPage)Project.GetObject(nUID);   
	return wp;
}

static string StockAnalyzer_GetTemplatePath (string strName){
	string strFilePath = __FILE__;
	strFilePath = GetFilePath(strFilePath) + strName;
	return strFilePath;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Graph Control

void SA_OverLay_Type(int nType) {
	// nType = 1  ==> SMA
	// nType = 2  ==> EMA
	// nType = 3  ==> Bollinger
	// nType = 4  ==> Chandelier Exit
	
	GraphLayer gl = Project.ActiveLayer().GetPage().Layers(0);		// Get 1st layer
	
	switch (nType) {
	case 1:		// SMA
		gl.DataPlots(0).Show = true;
		gl.DataPlots(1).Show = true;
		gl.DataPlots(2).Show = true;
		gl.DataPlots(3).Show = true;
		gl.DataPlots(4).Show = true;
		gl.DataPlots(5).Show = false;
		gl.DataPlots(6).Show = false;
		gl.DataPlots(7).Show = false;
		gl.DataPlots(8).Show = false;
		gl.DataPlots(9).Show = false;
		gl.DataPlots(10).Show = false;
		gl.DataPlots(11).Show = false;
		gl.GraphObjects("Legend_SMA").Show = true;
		gl.GraphObjects("Legend_EMA").Show = false;
		gl.GraphObjects("Legend_Bollinger").Show = false;
		break;
	case 2:		// EMA
		gl.DataPlots(0).Show = true;
		gl.DataPlots(1).Show = false;
		gl.DataPlots(2).Show = false;
		gl.DataPlots(3).Show = false;
		gl.DataPlots(4).Show = false;
		gl.DataPlots(5).Show = true;
		gl.DataPlots(6).Show = true;
		gl.DataPlots(7).Show = true;
		gl.DataPlots(8).Show = true;
		gl.DataPlots(9).Show = false;
		gl.DataPlots(10).Show = false;
		gl.DataPlots(11).Show = false;
		gl.GraphObjects("Legend_SMA").Show = false;
		gl.GraphObjects("Legend_EMA").Show = true;
		gl.GraphObjects("Legend_Bollinger").Show = false;
		break;
	case 3:	// Bollinger
		gl.DataPlots(0).Show = true;
		gl.DataPlots(1).Show = false;
		gl.DataPlots(2).Show = false;
		gl.DataPlots(3).Show = false;
		gl.DataPlots(4).Show = false;
		gl.DataPlots(5).Show = false;
		gl.DataPlots(6).Show = false;
		gl.DataPlots(7).Show = false;
		gl.DataPlots(8).Show = false;
		gl.DataPlots(9).Show = true;
		gl.DataPlots(10).Show = true;
		gl.DataPlots(11).Show = true;
		gl.GraphObjects("Legend_SMA").Show = false;
		gl.GraphObjects("Legend_EMA").Show = false;
		gl.GraphObjects("Legend_Bollinger").Show = true;
		break;
	default:		// default is SMA
		gl.DataPlots(0).Show = true;
		gl.DataPlots(1).Show = true;
		gl.DataPlots(2).Show = true;
		gl.DataPlots(3).Show = true;
		gl.DataPlots(4).Show = true;
		gl.DataPlots(5).Show = false;
		gl.DataPlots(6).Show = false;
		gl.DataPlots(7).Show = false;
		gl.DataPlots(8).Show = false;
		gl.DataPlots(9).Show = false;
		gl.DataPlots(10).Show = false;
		gl.DataPlots(11).Show = false;
		gl.GraphObjects("Legend_SMA").Show = true;
		gl.GraphObjects("Legend_EMA").Show = false;
		gl.GraphObjects("Legend_Bollinger").Show = false;
		break;
	};
	
	gl.GetPage().Refresh();
}
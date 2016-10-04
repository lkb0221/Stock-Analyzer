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
#define STR_YAHOO_API_ALL "http://ichart.finance.yahoo.com/table.csv?s=%s&g=d&ignore=.csv"
#define STR_FILE_NAME_DOWNLOAD "Table.csv"
#define STR_TEMPLATE_NAME_WBK "Stock Analyzer.ogw"

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main

void StockAnalyzer_Main() {
	// Get Config Tree
	Tree tr;
	tr = StockAnalyzer_GUI();
	if (tr.IsEmpty()) {
		return;
	}
	
	// Download File
	string strDownload = StockAnalyzer_Download(tr.Data.SymbolName.strVal, tr.Data.StartDate.dVal, tr.Data.EndDate.dVal);
	
	// Setup workbook
	WorksheetPage wp;
	wp = StockAnalyzer_Setup(tr, strDownload);
	
	// Delete downloaded file
	StockAnalyzer_Delete(strDownload);
	
	// Force refresh
	wp.LT_execute("run -p au;");
	set_active_layer(wp.Layers(1));
	set_active_layer(wp.Layers(0));
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
		GETN_RADIO_INDEX_EX(Range, "Range", 1, "Last Year|Last Three Year|All|Custom")
		GETN_DATE(StartDate, "Date From", dToday-365*2)	GETN_CURRENT_SUBNODE.Show = false;
		GETN_DATE(EndDate, "Date To", dToday-1)					GETN_CURRENT_SUBNODE.Show = false;
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
		
		GETN_BEGIN_BRANCH(ADL, "Accumulation Distribution")
		GETN_END_BRANCH(ADL)
		
		GETN_BEGIN_BRANCH(Aroon, "Aroon")
			GETN_NUM(Period, "Look-back Period Window", 25)	GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(Aroon)
		
		GETN_BEGIN_BRANCH(AroonO, "Aroon Oscillator")
		GETN_END_BRANCH(AroonO)
		
		GETN_BEGIN_BRANCH(ADX, "Average Directional Index (ADX)")
			GETN_NUM(Period, "Look-back Period Window", 14)	GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(ADX)
		
		GETN_BEGIN_BRANCH(ATR, "Average True Range (ATR)")
			GETN_RADIO_INDEX_EX(Method, "Ture Range Method", 0, "Current High less the current Low|Current High less the previous Close (absolute value)|Current Low less the previous Close (absolute value)")
			GETN_NUM(Period, "Period Window", 14)
		GETN_END_BRANCH(ATR)
		
		GETN_BEGIN_BRANCH(BBW, "Bollinger Band Width")
		GETN_END_BRANCH(BBW)
		
		GETN_BEGIN_BRANCH(pB, "%B Indicator")
		GETN_END_BRANCH(pB)
		
		GETN_BEGIN_BRANCH(CCI, "Commodity Channel Index (CCI)")
			GETN_NUM(Period, "Look-back Period Window", 20)	GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(CCI)
		
		GETN_BEGIN_BRANCH(COPP, "Coppock Curve")
			GETN_NUM(Period1, "1st ROC Period", 14)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd ROC Period", 11)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "WMA Period", 10)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(COPP)
		
		GETN_BEGIN_BRANCH(CMF, "Chaikin Money Flow (CMF)")
		GETN_END_BRANCH(CMF)
		
		GETN_BEGIN_BRANCH(ChiOsc, "Chaikin Oscillator")
			GETN_NUM(Period1, "1st ADL EMA Window", 3)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd ADL EMA Window", 10)				GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(ChiOsc)
		
		GETN_BEGIN_BRANCH(PMO, "DecisionPoint Price Momentum Oscillator (PMO)")
			GETN_NUM(Period1, "1st Time Period", 35)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Time Period", 20)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(SignalPeriod, "Signal EMA Period", 10)		GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(PMO)
		
		GETN_BEGIN_BRANCH(DPO, "Detrended Price Oscillator (DPO)")
			GETN_NUM(Period, "Window", 20)					GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(DPO)
		
		GETN_BEGIN_BRANCH(EMV, "Ease of Movement")
			GETN_NUM(Period, "Period", 14)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(EMV)
		
		GETN_BEGIN_BRANCH(Force, "Force Index")
			GETN_NUM(Period, "Period", 13)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(Force)
		
		GETN_BEGIN_BRANCH(Mass, "Mass Index")
			GETN_NUM(Period, "Period", 25)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(Mass)
		
		GETN_BEGIN_BRANCH(MACD, "Moving Average Convergence/Divergence Oscillator (MACD)")
			GETN_NUM(WinEMA1, "EMA Window 1", 12)		GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(WinEMA2, "EMA Window 2", 26)		GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(SignalPeriod, "Signal Period", 9)			GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(MACD)
		
		GETN_BEGIN_BRANCH(MFI, "Money Flow Index (MFI)")
			GETN_NUM(Period, "Period", 14)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(MFI)
		
		GETN_BEGIN_BRANCH(NVI, "Negative Volume Index (NVI)")
			GETN_NUM(Period, "Period", 255)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(NVI)
		
		GETN_BEGIN_BRANCH(OBV, "On Balance Volume (OBV)")
		GETN_END_BRANCH(OBV)
		
		GETN_BEGIN_BRANCH(PPO, "Percentage Price Oscillator (PPO)")
			GETN_NUM(Period1, "1st Time Period", 12)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Time Period", 26)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(SignalPeriod, "Signal EMA Period", 9)		GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(PPO)
		
		GETN_BEGIN_BRANCH(PVO, "Percentage Volume Oscillator (PVO)")
			GETN_NUM(Period1, "1st Time Period", 12)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Time Period", 26)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(SignalPeriod, "Signal EMA Period", 9)		GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(PVO)
		
		GETN_BEGIN_BRANCH(McClellan, "McClellan Oscillator")
			GETN_NUM(Period1, "Period Window 1", 19)		GETN_CURRENT_SUBNODE.Enable = false;
			GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "Period Window 2", 39)		GETN_CURRENT_SUBNODE.Enable = false;
			GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(McClellan)
		
		GETN_BEGIN_BRANCH(StochasticOscillator, "Stochastic Oscillator")
			GETN_NUM(PeriodK, "K Period", 14)	GETN_OPTION_NUM_FORMAT( ".0" )		// %K
			GETN_NUM(PeriodD, "D Period", 3)	GETN_OPTION_NUM_FORMAT( ".0" )		// %D
			//GETN_CHECK(Slow, "Include Slow Stochastic Oscillator", false)
			//GETN_CHECK(Full, "Include Full Stochastic Oscillator", false)
		GETN_END_BRANCH(StochasticOscillator)
		
		GETN_BEGIN_BRANCH(ROC, "Rate of Change (ROC)")
			GETN_NUM(Period1, "1st Period", 250)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Period", 125)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "3rd Period", 63)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period4, "4th Period", 21)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(ROC)
		
		
		
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
	
	if (0 == lstrcmp(lpcszNodeName, "Range")) {
		StockAnalyzer_GUI_event_Input(tr.Data);
	};
	
	
	
	return 0;
}

static bool StockAnalyzer_GUI_event_Input(Tree &tr) {
	switch(tr.Method.nVal) {
	case 0:	// yahoo api
		StockAnalyzer_GUI_event_InputAPI(tr);
		break;
	case 1:	// column
		break;
	}
	
	return true;
}

static bool StockAnalyzer_GUI_event_InputAPI(Tree &tr) {
	switch (tr.Range.nVal) {
	case 0:	// last year
		tr.StartDate.Show = false;
		tr.EndDate.Show = false;
		tr.EndDate.dVal = StockAnalyzer_GetToday() - 1;
		tr.StartDate.dVal = tr.EndDate.dVal - 365;
		break;
	case 1:	// Last 3 years
		tr.StartDate.Show = false;
		tr.EndDate.Show = false;
		tr.EndDate.dVal = StockAnalyzer_GetToday() - 1;
		tr.StartDate.dVal = tr.EndDate.dVal - 365*3;
		break;
	case 2:	// All
		tr.StartDate.Show = false;
		tr.EndDate.Show = false;
		tr.EndDate.dVal = NANUM;
		tr.StartDate.dVal = NANUM;
		break;
	case 3:
		tr.StartDate.Show = true;
		tr.EndDate.Show = true;
		if (is_missing_value(tr.EndDate.dVal)) {
			tr.EndDate.dVal = StockAnalyzer_GetToday() - 1;
		};
		if (is_missing_value(tr.StartDate.dVal)) {
			tr.StartDate.dVal = tr.EndDate.dVal - 365;
		}
		break;
	};
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
	if (is_missing_value(dStartDate) || is_missing_value(dEndDate)) {
		strURL.Format(STR_YAHOO_API_BASE, strSymbolName);
	}
	else {
		strURL.Format(STR_YAHOO_API_BASE, strSymbolName, vsStart[1], vsStart[2], vsStart[0], vsEnd[1], vsEnd[2], vsEnd[0]);
	}
	
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

static WorksheetPage StockAnalyzer_Setup(Tree tr, string strFilePath) {
	WorksheetPage wp;
	string strTemplatePath = StockAnalyzer_GetTemplatePath(STR_TEMPLATE_NAME_WBK);
	wp.Create(strTemplatePath);
	set_user_info(wp, "Overlays", tr.Overlays);
	set_user_info(wp, "Indicators", tr.Indicators);
	Worksheet wksData = wp.Layers("Data");
	
	StockAnalyzer_Import(strFilePath, wksData);
	
	return wp;
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
			//wks.LT_execute("run -p au;");
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
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);	
	StockAnalyzer_Aroon_Main(wksIndicator, wksData, vsIndex, trIndicators.Aroon);
	
	// ATR
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_ATR_Main(wksIndicator, wksData, vsIndex, trIndicators.ATR);
	
	// ADX
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);
	StockAnalyzer_ADX_Main(wksIndicator, wksData, vsIndex, trIndicators.ADX);
	
	// Band Width
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_BBWidth_Main(wksIndicator, wksOverLay, vsIndex, trOverlays.BB);
	
	// %B
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_PerB_Main(wksIndicator, wksOverLay, vsIndex, trOverlays.BB);
	
	// CCI
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);	
	StockAnalyzer_CCI_Main(wksIndicator, wksData, vsIndex, trIndicators.CCI);
	
	// ROC
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 4);	
	StockAnalyzer_ROC_Main(wksIndicator, wksData, vsIndex, trIndicators.ROC);
	
	// COPP
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_COPP_Main(wksIndicator, wksData, vsIndex, trIndicators.COPP);
	
	// ChiOsc
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_ChiOsc_Main(wksIndicator, wksIndicator, vsIndex, trIndicators.ChiOsc);
	
	// PMO
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_PMO_Main(wksIndicator, wksData, vsIndex, trIndicators.PMO);
	
	// DPO
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_DPO_Main(wksIndicator, wksData, vsIndex, trIndicators.DPO);
	
	// EMV
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_EMV_Main(wksIndicator, wksData, vsIndex, trIndicators.EMV);
	
	// Force
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_Force_Main(wksIndicator, wksData, vsIndex, trIndicators.Force);
	
	// Mass
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_Mass_Main(wksIndicator, wksData, vsIndex, trIndicators.Mass);
	
	// MACD
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);	
	StockAnalyzer_MACD_Main(wksIndicator, wksData, vsIndex, trIndicators.MACD);
	
	// MFI
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_MFI_Main(wksIndicator, wksData, vsIndex, trIndicators.MFI);
	
	// NVI
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_NVI_Main(wksIndicator, wksData, vsIndex, trIndicators.NVI);
	
	// OBV
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_OBV_Main(wksIndicator, wksData, vsIndex);
	
	// PPO
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);	
	StockAnalyzer_PPO_Main(wksIndicator, wksData, vsIndex, trIndicators.PPO);
	
	// PVO
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);	
	StockAnalyzer_PVO_Main(wksIndicator, wksData, vsIndex, trIndicators.PVO);
	
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
	vector<double> vsSMA(nSize);	vsSMA[0] = vsClose[0];
	if (is_missing_value(nWin)) {
		return NULL;
	}
	
	for (int ii = 1; ii < nSize; ii++) {
		if (ii + 1 - nWin <= 0) {
			//vsSMA[ii] = NANUM;
			vsSMA[ii] = ((vsSMA[ii-1] * ii) + vsClose[ii]) / (ii+1);
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
	
	//vsSMA = StockAnalyzer_SMA(vsClose, nWin);
	double dMultiplier = 2 / ((double)nWin + 1);
	
	// Calulate EMA
	//bool bStart;
	vsEMA[0] = vsClose[0];
	for (int ii = 1; ii < nSize; ii++) {
		//if (is_missing_value(vsSMA[ii])) {
			//vsEMA[ii] = NANUM;
		//}
		//else if (!bStart){
			//vsEMA[ii] = vsSMA[ii];
			//bStart = true;
		//}
		if (ii + 1 - nWin <= 0) {
			vsEMA[ii] = ((vsEMA[ii-1] * ii) + vsClose[ii]) / (ii+1);
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
	
	Dataset dsUp(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Aroon", "Arron Up"));
	Dataset dsDown(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Aroon", "Arron Down"));
	Dataset dsOscillator(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Aroon", "Aroon(" + ftoa(nWin) + ") Oscillator"));
	
	//Column ccUp(wksTarget, nIndex[0]);				Dataset dsUp(ccUp);
	//Column ccDown(wksTarget, nIndex[1]);			Dataset dsDown(ccDown);
	//Column ccOscillator(wksTarget, nIndex[2]); 	Dataset dsOscillator(ccOscillator);
	//ccUp.SetLongName("Aroon");								ccUp.SetComments("Arron Up");
	//ccDown.SetLongName("Aroon");							ccDown.SetComments("Arron Down");
	//ccOscillator.SetLongName("Aroon");	ccOscillator.SetComments("Aroon(" + ftoa(nWin) + ") Oscillator");
	
	StockAnalyzer_Aroon(dsClose, dsUp, dsDown, nWin);
	dsOscillator = dsUp - dsDown;
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

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Average True Range (ATR)

static void StockAnalyzer_ATR_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// 
	
	int nMethodTR = tr.GetNode("Method").nVal;
	int nPeriod = tr.GetNode("Period").nVal;
	
	//Dataset dsTR(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "True Range", StockAnalyzer_Legend("TR", nPeriod)));
	Dataset dsATR(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Average True Range", StockAnalyzer_Legend("ATR", nPeriod)));
	
	//dsTR = StockAnalyzer_TR(dsClose, dsHigh, dsLow, nMethodTR);
	dsATR = StockAnalyzer_ATR(dsClose, dsHigh, dsLow, nPeriod, nMethodTR);
}

static vector<double> StockAnalyzer_ATR(vector<double> vsClose, vector<double> vsHigh, vector<double> vsLow, int nPeriod, int nMode = 0) {
	int nSize = vsClose.GetSize();
	vector<double> vsATR(nSize);
	vector<double> vsTR(nSize);
	
	vsTR = StockAnalyzer_TR(vsClose, vsHigh, vsLow, nMode);
	
	// Starting point
	vector<double> vsSub(nPeriod);
	vsTR.GetSubVector(vsSub, 0, nPeriod-1);
	double dMean;		ocmath_basic_summary_stats(nPeriod, vsSub, NULL, &dMean);
	
	// Calculation
	for (int ii = 0; ii < nSize; ii++) {
		if (ii < nPeriod-1) {
			vsATR[ii] = NANUM;
		}
		else if (ii == nPeriod-1) {
			vsATR[ii] = dMean;
		}
		else {
			vsATR[ii] = ((vsATR[ii-1] * (nPeriod - 1)) + vsTR[ii]) / nPeriod;
		}
	};
	
	return vsATR;
}

static vector<double> StockAnalyzer_TR(vector<double> vsClose, vector<double> vsHigh, vector<double> vsLow, int nMode = 0) {
	int nSize = vsClose.GetSize();
	vector<double> vsTR(nSize);
	
	switch (nMode) {
	case 0:
		vsTR = vsHigh - vsLow;
		break;
	case 1:
		vsTR[0] = NANUM;
		for (int ii = 1; ii < nSize; ii++) {
			vsTR[ii] = abs(vsHigh[ii] - vsClose[ii-1]);
		};
		break;
	case 2:
		vsTR[0] = NANUM;
		for (int ii = 1; ii < nSize; ii++) {
			vsTR[ii] = abs(vsLow[ii] - vsClose[ii-1]);
		};
		break;
	};
	
	return vsTR;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Average Directional Index (ADX)

static void StockAnalyzer_ADX_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// 
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsADX(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Average Directional Index", StockAnalyzer_Legend("ADX", nPeriod)));
	Dataset dsDIp(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Average Directional Index", "+DI"));
	Dataset dsDIn(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Average Directional Index", "-DI"));
	
	dsADX = StockAnalyzer_ADX(dsHigh, dsLow, dsClose, dsDIp, dsDIn, nPeriod);
	
	
}

static vector<double> StockAnalyzer_ADX(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, vector<double>& vsDIp, vector<double>& vsDIn, int nWin = 14) {
	vector<double> vsTR;
	vector<double> vsDMp, vsDMn;
	vector<double> vsADX;
	
	vsTR = StockAnalyzer_TR(vsClose, vsHigh, vsLow);
	StockAnalyzer_DM(vsHigh, vsLow, vsDMp, vsDMn);
	vsTR = StockAnalyzer_EMA(vsTR, nWin);
	vsDMp = StockAnalyzer_EMA(vsDMp, nWin);
	vsDMn = StockAnalyzer_EMA(vsDMn, nWin);
	vsDIp = vsDMp / vsTR * 100;
	vsDIn = vsDMn / vsTR * 100;
	
	vsADX = abs(vsDIp - vsDIn) / (vsDIp + vsDIn) * 100;
	vsADX = StockAnalyzer_EMA(vsADX, nWin);
	
	return vsADX;
}

static void StockAnalyzer_DM(vector<double> vsHigh, vector<double> vsLow, vector<double>& vsDMp, vector<double>& vsDMn) {
	int nSize = vsHigh.GetSize();
	vsDMp.SetSize(nSize);
	vsDMn.SetSize(nSize);
	for (int ii = 1; ii < nSize; ii++) {
		double dDIp = vsHigh[ii] - vsHigh[ii-1];
		double dDIn = vsLow[ii-1] - vsLow[ii];
		if (dDIp >= dDIn) {
			vsDMp[ii] = dDIp > 0? dDIp:0;
			vsDMn[ii] = 0;
		}
		else {
			vsDMp[ii] = 0;
			vsDMn[ii] = dDIn > 0? dDIn:0;
		}
	}
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Bollinger BandWidth

static void StockAnalyzer_BBWidth_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsUpper(wksSource, 7);		
	Dataset dsLower(wksSource, 8);	
	Dataset dsCenter(wksSource, 9);	
	
	int nPeriod = tr.GetNode("Period").nVal;
	int nWidth = tr.GetNode("Band").nVal;
	
	Dataset dsBBW(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Bollinger BandWidth", StockAnalyzer_Legend("BB Width", nPeriod, nWidth)));
	
	dsBBW = StockAnalyzer_BBWidth(dsUpper, dsLower, dsCenter);
}

static vector<double> StockAnalyzer_BBWidth(vector<double> vsUpper, vector<double> vsLower, vector<double> vsMiddle) {
	return (vsUpper - vsLower) / vsMiddle * 100;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// %B Indicator

static void StockAnalyzer_PerB_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsUpper(wksSource, 7);		
	Dataset dsLower(wksSource, 8);	
	Dataset dsClose(StockAnalyzer_GetSourceColumn(wksSource, 4));	// Close value
	
	int nPeriod = tr.GetNode("Period").nVal;
	int nWidth = tr.GetNode("Band").nVal;
	
	Dataset dsB(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "%B Indicator", StockAnalyzer_Legend("%B", nPeriod, nWidth)));
	dsB = StockAnalyzer_PercentageB(dsUpper, dsLower, dsClose);
}

static vector<double> StockAnalyzer_PercentageB(vector<double> vsUpper, vector<double> vsLower, vector<double> vsPrice) {
	return (vsPrice - vsLower) / (vsUpper - vsLower);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Commodity Channel Index (CCI)

static void StockAnalyzer_CCI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// 
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset ds100p(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Commodity Channel Index", ""));
	Dataset dsCCI(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Commodity Channel Index", StockAnalyzer_Legend("CCI", nPeriod)));
	Dataset ds100n(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Commodity Channel Index", ""));
	
	dsCCI = StockAnalyzer_CCI(dsHigh, dsLow, dsClose, nPeriod);
	int nSize = dsCCI.GetSize();
	ds100p.SetSize(nSize);		ds100p = 100;
	ds100n.SetSize(nSize);		ds100n = -100;
}

static vector<double> StockAnalyzer_CCI(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, int nWin) {
	int nSize = vsClose.GetSize();
	vector<double> vsCCI(nSize);
	vector<double> vsTP(nSize), vsTPMA(nSize), vsMD(nSize);
	
	double dConstant = 0.015;
	vsTP = (vsHigh + vsLow + vsClose) / 3;
	vsTPMA = StockAnalyzer_SMA(vsTP, 20);
	
	// Calculate Mean Deviation
	vsMD = StockAnalyzer_Moving_Mean_Deviation(vsTP, vsTPMA, 20);
	
	vsCCI = (vsTP - vsTPMA) / (dConstant * vsMD);
	
	return vsCCI;
}

static vector<double> StockAnalyzer_Moving_Mean_Deviation(vector<double> vsValue, vector<double> vsSMA, int nWin = 20) {
	int nSize = vsValue.GetSize();
	vector<double> vsMD(nSize);
	for(int ii = nWin - 1; ii < nSize; ii++) {
		vector<double> vsSub;
		vsValue.GetSubVector(vsSub, ii-nWin+1, ii);
		double dMean = vsSMA[ii];
		vsSub -= dMean;
		vsSub.Abs();
		double dSum;
		vsSub.Sum(dSum);
		vsMD[ii] = dSum / nWin;
	};
	return vsMD;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Rate of Change (ROC)

static void StockAnalyzer_ROC_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	for (int ii = 0; ii < 4; ii++) {
		int nWin = tree_get_node(tr, ii).nVal;
		Dataset dsROC(StockAnalyzer_SetColumn(wksTarget, nIndex[ii], "Rate of Change", StockAnalyzer_Legend("ROC", nWin)));
		
		if ((is_missing_value(nWin)) || (nWin <= 0)) {
			continue;
		}
		
		dsROC = StockAnalyzer_ROC(dsClose, nWin);
	}
}

static vector StockAnalyzer_ROC(vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsROC(nSize);	vsROC = NANUM;
	for (int ii = nPeriod; ii < nSize; ii++) {
		vsROC[ii] = (vsClose[ii] - vsClose[ii - nPeriod]) / vsClose[ii - nPeriod] * 100;
	}
	return vsROC;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Coppock Curve

static void StockAnalyzer_COPP_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;		// 1st ROC
	int nPeriod2 = tr.GetNode("Period2").nVal;		// 2nd ROC
	int nPeriod3 = tr.GetNode("Period3").nVal;		// WMA
	
	Dataset dsCOPP(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Coppock Curve", StockAnalyzer_Legend("COPP", nPeriod1, nPeriod2, nPeriod3)));
	dsCOPP = StockAnalyzer_Coppock(dsClose, nPeriod1, nPeriod2, nPeriod3);
	
}

static vector<double> StockAnalyzer_Coppock(vector<double> vsClose, int nPeriod1, int nPeriod2, int nPeriod3) {
	// Coppock Curve = period3 WMA of (period1 RoC + perod2 RoC)
	vector<double> vsRoC1, vsRoC2;
	vector<double> vsWMA;
	vector<double> vsCoppock;		vsCoppock = NANUM;
	if (is_missing_value(nPeriod1) || is_missing_value(nPeriod2) || is_missing_value(nPeriod3)) {
		return vsCoppock;
	};
	vector<int> vsCheck;
	vsCheck.Add(nPeriod1);
	vsCheck.Add(nPeriod2);
	vsCheck.Add(nPeriod3);
	if (vsClose.GetSize() <= min(vsCheck)) {
		return vsCoppock;
	};
	
	vsRoC1 = StockAnalyzer_ROC(vsClose, nPeriod1);
	vsRoC2 = StockAnalyzer_ROC(vsClose, nPeriod2);
	
	vsCoppock = StockAnalyzer_WMA(vsRoC1 + vsRoC2, nPeriod3);
	
	return vsCoppock;
}

static vector<double> StockAnalyzer_WMA(vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsWMA(nSize);		vsWMA = NANUM;
	
	vector<double> vsMultiplier;			vsMultiplier.Data(1, nPeriod, 1);
	double dTotal;								vsMultiplier.Sum(dTotal);
	for (int ii = nPeriod-1; ii < nSize; ii++) {
		vector<double> vsSub;
		vsClose.GetSubVector(vsSub, ii - nPeriod + 1, ii);
		vsSub = vsSub * vsMultiplier;
		double dSum;	vsSub.Sum(dSum);
		vsWMA[ii] = dSum/dTotal;
	}
	
	return vsWMA;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Chaikin Oscillator

static void StockAnalyzer_ChiOsc_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsADL(wksSource, 1);		// ADL column that been calculated before
	
	int nPeriod1 = tr.GetNode("Period1").nVal;		// 1st EMA
	int nPeriod2 = tr.GetNode("Period2").nVal;		// 2nd EMA
	
	Dataset dsChiOsc(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Chaikin Oscillator", StockAnalyzer_Legend("ChiOsc", nPeriod1, nPeriod2)));
	dsChiOsc = StockAnalyzer_ChiOsc(dsADL, nPeriod1, nPeriod2);
}

static vector<double> StockAnalyzer_ChiOsc(vector<double> vsADL, int nPeriod1 = 3, int nPeriod2 = 10) {
	return StockAnalyzer_EMA(vsADL, nPeriod1) - StockAnalyzer_EMA(vsADL, nPeriod2);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//DecisionPoint Price Momentum Oscillator (PMO)

static void StockAnalyzer_PMO_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;				// 1st EMA
	int nPeriod2 = tr.GetNode("Period2").nVal;				// 2nd EMA
	int nPeriod3 = tr.GetNode("SignalPeriod").nVal;		// Signal window
	
	Dataset dsPMO(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "DecisionPoint Price Momentum Oscillator", StockAnalyzer_Legend("PMO", nPeriod1, nPeriod2, nPeriod3)));
	Dataset dsPMOS(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "DecisionPoint Price Momentum Oscillator", ""));
	
	dsPMO = StockAnalyzer_PMO(dsClose, dsPMOS, nPeriod1, nPeriod2, nPeriod3);
}

static vector<double> StockAnalyzer_PMO(vector<double> vsPrice, vector<double> &vsSignal, int nPeriod1 = 35, int nPeriod2 = 20, int nPeriod3 = 10) {
	vector<double> vsPMO;
	double dSmoothFatctor1 = 2 / (double)nPeriod1;
	double dSmoothFatctor2 = 2 / (double)nPeriod2;
	
	vector<double> vsROC_1, vsSmooth1;
	vsROC_1 = StockAnalyzer_ROC(vsPrice, 1);
	vsSmooth1 = StockAnalyzer_PMO_SmoothEMA(vsROC_1, nPeriod1, dSmoothFatctor1);
	vsPMO = StockAnalyzer_PMO_SmoothEMA(vsSmooth1 * 10, nPeriod2, dSmoothFatctor2);
	vsSignal = StockAnalyzer_EMA(vsPMO, nPeriod3);
	
	return vsPMO;
}

static vector<double> StockAnalyzer_PMO_SmoothEMA(vector<double> vs, int nPeriod, double dSmoothFatctor) {
	int nSize = vs.GetSize();
	vector<double> vsSmooth(nSize);
	
	bool bStart = false;
	double dSum;
	int nCount;
	for (int ii = 0; ii < nSize; ii++) {
		if (is_missing_value(vs[ii])) {
			vsSmooth[ii] = NANUM;
		}
		else if (!bStart) {
			if (nCount == nPeriod - 1) {
				vsSmooth[ii] = (dSum + vs[ii]) / nPeriod;
				bStart = true;
			}
			else {
				dSum += vs[ii];
				vsSmooth[ii] = NANUM;
				nCount++;
			}
		}
		else {
			vsSmooth[ii] = (vs[ii] - vsSmooth[ii-1]) * dSmoothFatctor + vsSmooth[ii-1];
		}
	};
	return vsSmooth;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Detrended Price Oscillator (DPO)

static void StockAnalyzer_DPO_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod = tr.GetNode("Period").nVal;				
	
	Dataset dsDPO(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Detrended Price Oscillator", StockAnalyzer_Legend("DPO", nPeriod)));
	dsDPO = StockAnalyzer_DPO(dsClose, nPeriod);
}

static vector<double> StockAnalyzer_DPO(vector<double> vsPrice, int nPeriod) {
	int nOffset = nPeriod / 2 + 1;
	
	vector<double> vsMA, vsDPO;
	vsMA = StockAnalyzer_SMA(vsPrice, nPeriod);		// Get N-period SMA
	vector<int> vsOffsetRemove;	vsOffsetRemove.Data(0, nOffset-1, 1);
	vsMA.RemoveAt(vsOffsetRemove);	vsPrice.SetSize(vsMA.GetSize());
	vsDPO = vsPrice - vsMA;
	
	return vsDPO;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Ease of Movement (EMV)

static void StockAnalyzer_EMV_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsVolume(wksSource, 5);
	
	int nPeriod = tr.GetNode("Period").nVal;			
	
	Dataset dsEMV(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Ease of Movement", StockAnalyzer_Legend("EMV", nPeriod)));
	dsEMV = StockAnalyzer_EMV(dsHigh, dsLow, dsVolume, nPeriod);
}

static vector<double> StockAnalyzer_EMV(vector<double> vsHigh, vector<double> vsLow, vector<double> vsVol, int nPeriod) {
	int nSize = vsHigh.GetSize();
	vector<double> vsEMV(nSize);
	
	// 1-Period EMV = ((H + L)/2 - (Prior H + Prior L)/2) / ((V/100,000,000)/(H - L))
	for (int ii = 1; ii < nSize; ii++) {					
		vsEMV[ii] = ((vsHigh[ii] + vsLow[ii]) / 2 - (vsHigh[ii-1] + vsLow[ii-1]) / 2) / ((vsVol[ii] / 100000000) / (vsHigh[ii] - vsLow[ii]));
	};
	
	vsEMV = StockAnalyzer_EMA(vsEMV, nPeriod);
	
	return vsEMV;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Force Index

static void StockAnalyzer_Force_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	Dataset dsVolume(wksSource, 5);
	
	int nPeriod = tr.GetNode("Period").nVal;			
	
	Dataset dsForce(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Force Index", StockAnalyzer_Legend("FORCE", nPeriod)));
	dsForce = StockAnalyzer_Force(dsClose, dsVolume, nPeriod);
}

static vector<double> StockAnalyzer_Force(vector<double> vsClose, vector<double> vsVol, int nPeriod = 14) {
	// Force Index(1) = {Close (today)  -  Close (yesterday)} x Volume
	// Force Index(N) = N-period EMA of Force Index(1)
	
	vector<double> vsForce;
	
	vector<double> vsPreviousClose;
	vsClose.GetSubVector(vsPreviousClose, 0, vsClose.GetSize()-2);
	vsPreviousClose.InsertAt(0, NANUM);
	vsForce = (vsClose - vsPreviousClose) * vsVol;
	vsForce = StockAnalyzer_EMA(vsForce, nPeriod);
	
	return vsForce;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Mass Index

static void StockAnalyzer_Mass_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	
	int nPeriod = tr.GetNode("Period").nVal;	
	
	Dataset dsMass(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Mass Index", StockAnalyzer_Legend("MASS", nPeriod)));
	dsMass = StockAnalyzer_Mass(dsHigh, dsLow, nPeriod);
}

static vector<double> StockAnalyzer_Mass(vector<double> vsHigh, vector<double> vsLow, int nPeriod = 25) {
	vector<double> vsMass;
	
	vector<double> vsDiff, vsSingleEMA, vsDoubleEMA, vsRatioEMA;
	vsDiff = vsHigh - vsLow;
	vsSingleEMA = StockAnalyzer_EMA(vsDiff, 9);
	vsDoubleEMA = StockAnalyzer_EMA(vsSingleEMA, 9);
	vsRatioEMA = vsSingleEMA / vsDoubleEMA;
	vsMass = StockAnalyzer_SMA(vsRatioEMA, nPeriod) * nPeriod;
	
	return vsMass;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//MACD (Moving Average Convergence/Divergence Oscillator)

static void StockAnalyzer_MACD_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("WinEMA1").nVal;
	int nPeriod2 = tr.GetNode("WinEMA2").nVal;
	int nPeriod3 = tr.GetNode("SignalPeriod").nVal;
	
	Dataset dsMACD(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Moving Average Convergence/Divergence Oscillator", StockAnalyzer_Legend("MACD", nPeriod1, nPeriod2,nPeriod3)));
	Dataset dsSignal(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Moving Average Convergence/Divergence Oscillator", ""));
	Dataset dsHisto(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Moving Average Convergence/Divergence Oscillator", ""));
	
	dsMACD = StockAnalyzer_MACD(dsClose, nPeriod1, nPeriod2);
	dsSignal = StockAnalyzer_EMA(dsMACD, nPeriod3);
	dsHisto = dsMACD - dsSignal;
}

static vector<double> StockAnalyzer_MACD(vector<double> vsClose, int nPeriod1 = 12, int nPeriod2 = 26) {
	vector<double> vsMACD;
	
	vsMACD = StockAnalyzer_EMA(vsClose, nPeriod1) - StockAnalyzer_EMA(vsClose, nPeriod2);
	
	return vsMACD;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Money Flow Index (MFI)

static void StockAnalyzer_MFI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// real close
	Dataset dsVolume(wksSource, 5);
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsMFI(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Money Flow Index", StockAnalyzer_Legend("MFI", nPeriod)));
	dsMFI = StockAnalyzer_MFI(dsHigh, dsLow, dsClose, dsVolume, nPeriod);
}

static vector<double> StockAnalyzer_MFI(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, vector<double> vsVolume, int nPeriod) {
	vector<double> vsMFI;
	
	vector<int> vsDir;
	vector<double> vsTP, vsMF, vsMFr;
	vsDir = StockAnalyzer_UpDown(vsClose);
	vsTP = StockAnalyzer_TypicalPrice(vsHigh, vsLow, vsClose);
	vsMF = StockAnalyzer_MoneyFlow(vsTP, vsVolume);
	vsMFr = StockAnalyzer_MoneyFlowRatio(vsMF, vsDir, nPeriod);
	vsMFI = StockAnalyzer_MoneyFlowIndex(vsMFr);
	
	return vsMFI;
}

static vector<int> StockAnalyzer_UpDown(vector<double> vsClose) {
	int nSize = vsClose.GetSize();
	vector<int> vsDir(nSize);
	vsDir = NANUM;
	
	for (int ii = 1; ii < nSize; ii++) {
		if (vsClose[ii] > vsClose[ii-1]) {
			vsDir[ii] = 1;
		}
		else {
			vsDir[ii] = -1;
		}
	};
	
	return vsDir;
}

static vector<double> StockAnalyzer_TypicalPrice(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose) {
	return (vsHigh + vsLow + vsClose) / 3;
}

static vector<double> StockAnalyzer_MoneyFlow(vector<double> vsTypicalPrice, vector<double> vsVolume) {
	return vsTypicalPrice * vsVolume;
}

static vector<double> StockAnalyzer_MoneyFlowRatio(vector<double> vsMF, vector<int> vsDir, int nPeriod) {
	int nSize = vsMF.GetSize();
	vector<double> vsMFp(nSize), vsMFn(nSize), vsMFr(nSize);
	StockAnalyzer_SeperatePnN(vsMF, vsDir, vsMFp, vsMFn);
	vsMFp = StockAnalyzer_SMA(vsMFp, nPeriod) * nPeriod;
	vsMFn = StockAnalyzer_SMA(vsMFn, nPeriod) * nPeriod;
	vsMFr = vsMFp / vsMFn;
	
	return vsMFr;
}

static vector<double> StockAnalyzer_MoneyFlowIndex(vector<double> vsMoneyFlowRatio) {
	return 100 - 100 / (1 + vsMoneyFlowRatio);
}

static void StockAnalyzer_SeperatePnN(vector<double> vs, vector<int> vsDir, vector<double> &vsP, vector<double> &vsN) {
	int nSize = vs.GetSize();
	vsP = NANUM;
	vsN = NANUM;
	for (int ii = 0; ii < nSize; ii++) {
		if (vsDir[ii] >= 0) {
			vsP[ii] = vs[ii];
		}
		else {
			vsN[ii] = vs[ii];
		}
	}
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Negative Volume Index (NVI)

static void StockAnalyzer_NVI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	Dataset dsVolume(wksSource, 5);
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsNVI(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Negative Volume Index", StockAnalyzer_Legend("NVI, EMA", nPeriod)));
	Dataset dsNVIEMA(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Negative Volume Index", ""));
	
	dsNVI = StockAnalyzer_NVI(dsClose, dsVolume);
	dsNVIEMA = StockAnalyzer_EMA(dsNVI, nPeriod);
}

static vector<double> StockAnalyzer_NVI(vector<double> vsPrice, vector<double> vsVol) {
	int nSize = vsPrice.GetSize();
	vector<double> vsNVI(nSize);		vsNVI[0] = 1000;
	
	for(int ii = 1; ii < nSize; ii++) {
		if (vsVol[ii] < vsVol[ii-1]) {
			double dSPX = (vsPrice[ii] - vsPrice[ii-1]) / vsPrice[ii-1] * 100;
			vsNVI[ii] = vsNVI[ii-1] + dSPX;
		}
		else {
			vsNVI[ii] = vsNVI[ii-1];
		}
	};
	
	return vsNVI;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// On Balance Volume (OBV)

static void StockAnalyzer_OBV_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex) {
	Dataset dsClose(wksSource, 6);	// adj close
	Dataset dsVolume(wksSource, 5);
	
	Dataset dsOBV(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "On Balance Volume", "OBV"));
	dsOBV = StockAnalyzer_OBV(dsClose, dsVolume);
}

static vector<double> StockAnalyzer_OBV(vector<double> vsPrice, vector<double> vsVol) {
	int nSize = vsPrice.GetSize();
	vector<double> vsOBV(nSize);		
	vsOBV[0] = 0;
	
	for (int ii = 1; ii < nSize; ii++) {
		if (vsPrice[ii] > vsPrice[ii-1]) {
			vsOBV[ii] = vsOBV[ii-1] + vsVol[ii];
		}
		else if (vsPrice[ii] < vsPrice[ii-1]) {
			vsOBV[ii] =vsOBV[ii-1] - vsVol[ii];
		}
		else {
			vsOBV[ii] = vsOBV[ii-1];
		};
	};
	
	return vsOBV;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Percentage Price Oscillator (PPO)

static void StockAnalyzer_PPO_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;
	int nPeriod2 = tr.GetNode("Period2").nVal;
	int nPeriod3 = tr.GetNode("SignalPeriod").nVal;
	
	Dataset dsPPO(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Percentage Price Oscillator", StockAnalyzer_Legend("PPO", nPeriod1, nPeriod2, nPeriod3)));
	Dataset dsSignal(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Percentage Price Oscillator", ""));
	Dataset dsHisto(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Percentage Price Oscillator", ""));
	
	dsPPO = StockAnalyzer_PPO(dsClose, nPeriod1, nPeriod2);
	dsSignal = StockAnalyzer_EMA(dsPPO, nPeriod3);
	dsHisto = dsPPO - dsSignal;
}

static vector<double> StockAnalyzer_PPO(vector<double> vsPrice, int nPeriod1 = 12, int nPeriod2 = 26) {
	vector<double> vsPPO;
	
	vector<double> vsEMA1, vsEMA2;
	vsEMA1 = StockAnalyzer_EMA(vsPrice, nPeriod1);
	vsEMA2 = StockAnalyzer_EMA(vsPrice, nPeriod2);
	
	vsPPO = (vsEMA1 - vsEMA2) / vsEMA2 * 100;
	
	return vsPPO;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Percentage Volume Oscillator (PVO)

static void StockAnalyzer_PVO_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsVol(wksSource, 5);	// Volumne
	
	int nPeriod1 = tr.GetNode("Period1").nVal;
	int nPeriod2 = tr.GetNode("Period2").nVal;
	int nPeriod3 = tr.GetNode("SignalPeriod").nVal;
	
	Dataset dsPVO(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Percentage Volume  Oscillator", StockAnalyzer_Legend("PVO", nPeriod1, nPeriod2, nPeriod3)));
	Dataset dsSignal(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Percentage Volume  Oscillator", ""));
	Dataset dsHisto(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Percentage Volume  Oscillator", ""));
	
	dsPVO = StockAnalyzer_PVO(dsVol, nPeriod1, nPeriod2);
	dsSignal = StockAnalyzer_EMA(dsPVO, nPeriod3);
	dsHisto = dsPVO - dsSignal;
}

static vector<double> StockAnalyzer_PVO(vector<double> vsVol, int nPeriod1 = 12, int nPeriod2 = 26) {
	vector<double> vsPVO;
	
	vector<double> vsEMA1, vsEMA2;
	vsEMA1 = StockAnalyzer_EMA(vsVol, nPeriod1);
	vsEMA2 = StockAnalyzer_EMA(vsVol, nPeriod2);
	
	vsPVO = (vsEMA1 - vsEMA2) / vsEMA2 * 100;
	
	return vsPVO;
}






////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Misc

static Column StockAnalyzer_GetSourceColumn(Worksheet wks, int nIndex) {
	WorksheetPage wp;
	wp = wks.GetPage();
	Worksheet wksSource = wp.Layers("Data");
	Column cc(wksSource, nIndex);
	return cc;
}

static vector<int> StockAnalyzer_Offset(int &nIndex, int nSize) {
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

static Column StockAnalyzer_SetColumn(Worksheet wks, int nIndex, string strLongName, string strComments) {
	Column cc(wks, nIndex);
	cc.SetLongName(strLongName);
	cc.SetComments(strComments);
	return cc;
}

static string StockAnalyzer_Legend(string strName, int nParam1, int nParam2 = -1, int nParam3 = -1) {
	string str = strName + "(" + ftoa(nParam1);
	
	if (nParam2 != -1) {
		str = str + ", " + ftoa(nParam2);
	};
	
	if (nParam3 != -1) {
		str = str + ", " + ftoa(nParam3);
	};
	
	str = str + ")";
	
	return str;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Graph Control

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Overlay

void SA_OverLay_Type(int nType) {
	string strLegendName;
	vector<int> vsIndex;	
	SA_OverLayList(nType, vsIndex, strLegendName);
	
	GraphLayer gl = Project.ActiveLayer().GetPage().Layers(0);		// Get 1st layer
	
	SA_OverLay_ShowHide(gl, vsIndex, strLegendName);
	
	gl.GetPage().Refresh();
}

static void SA_OverLay_ShowHide(GraphLayer gl, vector<int> vsIndex, string strLegendName) {
	vector<uint> vecIndex;
	
	foreach (DataPlot dp in gl.DataPlots) {
		dp.Show = (vsIndex.Find(vecIndex, dp.GetIndex()) >= 1);
	};
	
	foreach (GraphObject grobj in gl.GraphObjects) {
		string strName = grobj.GetName();
		grobj.Show = (strncmp(strName, "Legend", 6) != 0) || (strcmp(strName, strLegendName) == 0);
	};
}

static bool SA_OverLayList(int nType, vector<int> &vsIndex, string &strLegendName) {
	switch (nType) {
	case 1:	// SMA
		vector<int> vsTemp = {0, 1, 2, 3};
		vsIndex = vsTemp;
		strLegendName = "LegendSMA";
		break;
	case 2:	// EMA
		vector<int> vsTemp = {0, 4, 5, 6};
		vsIndex = vsTemp;
		strLegendName = "LegendEMA";
		break;
	case 3:	// BB
		vector<int> vsTemp = {0, 7, 8, 9};
		vsIndex = vsTemp;
		strLegendName = "LegendBB";
		break;
	default:
		vector<int> vsTemp = {0, 1, 2, 3};
		vsIndex = vsTemp;
		strLegendName = "LegendSMA";
		break;
	}
	return true;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Indicator

void SA_Indicator_Type(int nType) {
	vector<string> saList;	saList = SA_IndicatorList();
	GraphPage gp = Project.ActiveLayer().GetPage();
	SA_Indicator_ShowHide(gp, saList[nType-1]);
	gp.Layers(saList[nType-1]).Rescale();
}

static void SA_Indicator_ShowHide(GraphPage gp, string strLayerName) {
	foreach(GraphLayer gl in gp.Layers) {
		string strName = gl.GetName();
		bool bHit = (strcmp(strName, "Stock") == 0) || (strcmp(strName, "Volume") == 0) || (strcmp(strName, strLayerName) == 0);
		gl.Show = bHit;
	};
}

static vector<string> SA_IndicatorList() {
	vector<string> sa;
	
	sa.Add("Accumulation Distribution Line");
	sa.Add("Aroon");
	sa.Add("Aroon Oscillator");
	sa.Add("Average True Range");
	sa.Add("Average Directional Index");
	sa.Add("Bollinger BandWidth");
	sa.Add("B Indicator");
	sa.Add("Commodity Channel Index");
	sa.Add("Coppock Curve");
	sa.Add("Chaikin Money Flow");
	sa.Add("Chaikin Oscillator");
	sa.Add("PMO");
	sa.Add("DPO");
	sa.Add("Ease of Movement");
	sa.Add("Force Index");
	sa.Add("Mass Index");
	sa.Add("MACD");
	sa.Add("Money Flow Index");
	sa.Add("Negative Volume Index");
	sa.Add("On Balance Volume");
	sa.Add("PPO");
	sa.Add("PVO");
	sa.Add("");
	sa.Add("");
	sa.Add("Rate of Change");
	
	return sa;
}


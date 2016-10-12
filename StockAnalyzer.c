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
#include <../OriginLab/OCHttpUtils.h>
#include <../OriginLab/DialogEx.h>
#include <../OriginLab/HTMLDlg.h>
#include <../OriginLab/GraphPageControl.h>
////////////////////////////////////////////////////////////////////////////////////
#define STR_HTML_NAME "Stock Analyzer"
#define STR_YAHOO_API_BASE "http://ichart.finance.yahoo.com/table.csv?s=%s&a=%d&b=%d&c=%d&d=%d&e=%d&f=%d&g=d&ignore=.csv"
#define STR_YAHOO_API_ALL "http://ichart.finance.yahoo.com/table.csv?s=%s&g=d&ignore=.csv"
#define STR_FILE_NAME_DOWNLOAD "Table.csv"
#define STR_TEMPLATE_NAME_WBK "Stock Analyzer.ogw"
#define GRAPH_CONTROL_ID 200
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main

class StockAnalyzerDlg: public HTMLDlg {
	public:
		StockAnalyzerDlg() {}
		~StockAnalyzerDlg() {}
		string GetInitURL();
		string GetDialogTitle();
		//void SetGraphTarget(string strGraphName);
	protected:
		bool OnInitDialog();
		bool OnDlgResize(int nType, int cx, int cy);
		bool GetInitDlgSize(int& width, int& height);
		bool OnDestroy() {return true;}
		void UpdateIndicator(int nIndex);
		void UpdateOverlay(int nIndex);
	public:
		EVENTS_BEGIN_DERIV(HTMLDlg)
			ON_INIT(OnInitDialog)
			ON_DESTROY(OnDestroy)
			ON_SIZE(OnDlgResize)
		EVENTS_END_DERIV
		
		DECLARE_DISPATCH_MAP
	private:	
		BOOL GetGraphRECT(RECT& gcCntrlRect)  {				//this is the function to call JavaScript and get the position of GraphControl
			if (!m_dhtml)
				return false;
			Object jsscript = m_dhtml.GetScript();
			
			if(!jsscript) //check the validity of returned COM object is always recommended
				return false;
			
			string str = jsscript.getGraphRECT();	
			JSON.FromString(gcCntrlRect, str); //convert a string to a structure
			return true;
		}
	private:
		GraphPage m_gp;
		GraphControl m_gpCntrl;
};

BEGIN_DISPATCH_MAP(StockAnalyzerDlg, HTMLDlg)
	DISP_FUNCTION(StockAnalyzerDlg, UpdateIndicator, VTS_VOID, VTS_I4)    
	DISP_FUNCTION(StockAnalyzerDlg, UpdateOverlay, VTS_VOID, VTS_I4)
END_DISPATCH_MAP

string StockAnalyzerDlg::GetInitURL() {
	string strPath = GetFilePath(__FILE__);
	strPath += STR_HTML_NAME + ".html";    //html file in same folder as this cpp file
	return strPath;
}

string StockAnalyzerDlg::GetDialogTitle() {
	string strTitle("Stock Analyzer");
	return strTitle;
}

/*
void StockAnalyzerDlg::SetGraphTarget(string strGraphName) {
	m_gp = Project.GraphPages(strGraphName);
}
*/

bool StockAnalyzerDlg::OnInitDialog() {
	HTMLDlg::OnInitDialog();
	//ModifyStyle(0, WS_MAXIMIZEBOX);
	//GraphControl gcCntrl;
	RECT rr;
	m_gpCntrl.CreateControl(GetSafeHwnd(), &rr, GRAPH_CONTROL_ID, WS_CHILD|WS_VISIBLE|WS_BORDER);
	DWORD 	dwNoClicks = OC_PREVIEW_NOCLICK_NOZOOMPAN | NOCLICK_LABEL | NOCLICK_BUTTONS;
	//GraphPage m_gp = Project.GraphPages("Graph1");
	m_gp = GetStockAnalyzerGraph();
	bool bb = m_gpCntrl.AttachPage(m_gp, dwNoClicks);
	return true;
}

bool StockAnalyzerDlg::GetInitDlgSize(int& width, int& height) {
	width = 1024;
	height = 500;
	return true;
}

///*
bool StockAnalyzerDlg::OnDlgResize(int nType, int cx, int cy) {
	if ( !IsInitReady() )
			return false;
		// MoveControlsHelper	_temp(this); // you can uncomment this line, if the dialog flickers when you resize it
		HTMLDlg::OnDlgResize(nType, cx, cy); //place html control in dialog
		
		if ( !IsHTMLDocumentCompleted() ) //check the state of HTML control
			return false;
		
		RECT rectGraph;
		if ( !GetGraphRECT(rectGraph))
			return false;
		//overlap the GraphControl on the top of HTML control
		m_gpCntrl.SetWindowPos(HWND_TOP, rectGraph.left, rectGraph.top,RECT_WIDTH(rectGraph),RECT_HEIGHT(rectGraph), 0); 
		
		return true;
}
//*/

/*
bool StockAnalyzerDlg::OnDlgResize(int nType, int cx, int cy) {
	if ( !IsInitReady() )
			return TRUE;
	
		MoveControlsHelper	_temp(this);
		
		RECT rect;
		GetClientRect(&rect);
		RECT rectHTML, rectGraph;
		rectHTML = rectGraph = rect;
		
		int nGap = GetControlGap();
		
		rectHTML.top = nGap;
		rectHTML.left = nGap;
		rectHTML.right = RECT_X(rect) - nGap;
		rectHTML.bottom -= nGap;
		MoveControl(m_dhtml, rectHTML);
		
		rectGraph.top = nGap;
		rectGraph.left = RECT_X(rect) + nGap;
		rectGraph.right -= nGap;
		rectGraph.bottom -= nGap;
		//MoveControl(m_gpCntrl.GetControl(), rectGraph);
		m_gpCntrl.SetWindowPos(HWND_TOP, rectGraph.left, rectGraph.top,RECT_WIDTH(rectGraph),RECT_HEIGHT(rectGraph), 0); 
		m_gpCntrl.GetPage().Refresh();
		
		return TRUE;
}
*/

void StockAnalyzerDlg::UpdateIndicator(int nIndex) {
	//out_int("Indicator = ", nIndex);
	SA_Indicator_Type(nIndex, m_gp);
}

void StockAnalyzerDlg::UpdateOverlay(int nIndex) {
	//out_int("Overlay = ", nIndex);
	SA_OverLay_Type(nIndex, m_gp);
}

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
	
	// Run Calculation
	StockAnalyzer_MainProcess(wp.GetUID());
	
	// Force refresh
	wp.LT_execute("run -p au;");
	set_active_layer(wp.Layers(1));
	set_active_layer(wp.Layers(0));
}

void StockAnalyzer_OpenHTML() {
	// Open HTML
	StockAnalyzerDlg dlg;
	//s_pDlg = &dlg;
	dlg.DoModalEx(GetWindow());
	//s_pDlg = NULL;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DIalog GUI

static Tree StockAnalyzer_GUI() {
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
		
		GETN_BEGIN_BRANCH(Chandlr, "Chandelier Exit")
			GETN_NUM(Period, "Look-back ATR Period Window", 22)	GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Factor, "ATR Mulitplier", 3)	
		GETN_END_BRANCH(Chandlr)
		
		GETN_BEGIN_BRANCH(Ichimoku, "Ichimoku Clouds")
			GETN_NUM(Period1, "Conversion Period", 9)				GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "Base Line Period", 26)				GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "Leading Span Period", 52)		GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(Ichimoku)
		
		
		
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
		
		GETN_BEGIN_BRANCH(KST, "Know Sure Thing (KST)")
			GETN_NUM(ROCp1, "1st ROC Period", 10)										GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCw1, "1st ROC Average Window", 10)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCp2, "2nd ROC Period", 15)										GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCw2, "2nd ROC Average Window", 10)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCp3, "3rd ROC Period", 20)										GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCw3, "3rd ROC Average Window", 10)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCp4, "4th ROC Period", 30)										GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(ROCw4, "4th ROC Average Window", 15)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(SignalPeriod, "Signal SMA Period", 9)							GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(KST)
		
		GETN_BEGIN_BRANCH(SPECIALK, "Martin Pring's Special K")
			GETN_NUM(Period1, "SMA Period", 100)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "SMA Smoothing Period", 100)		GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(SPECIALK)
		
		GETN_BEGIN_BRANCH(ROC, "Rate of Change (ROC)")
			GETN_NUM(Period1, "1st Period", 250)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Period", 125)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "3rd Period", 63)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period4, "4th Period", 21)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(ROC)
		
		GETN_BEGIN_BRANCH(RSI, "Relative Strength Index (RSI & StochRSI)")
			GETN_NUM(Period, "Period", 14)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(RSI)
		
		GETN_BEGIN_BRANCH(Slope, "Moving Slope")
			GETN_NUM(Period, "Period", 52)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(Slope)
		
		GETN_BEGIN_BRANCH(StdDev, "Standard Deviation (Volatility)")
			GETN_NUM(Period, "Period", 52)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(StdDev)
		
		GETN_BEGIN_BRANCH(StochasticOscillator, "Stochastic Oscillator")
			GETN_NUM(PeriodK, "K Period", 14)	GETN_OPTION_NUM_FORMAT( ".0" )		// %K
			GETN_NUM(PeriodD, "D Period", 3)	GETN_OPTION_NUM_FORMAT( ".0" )		// %D
			GETN_RADIO_INDEX_EX(Method, "Type", 0, "Fast|Slow|Full")	
		GETN_END_BRANCH(StochasticOscillator)
		
		GETN_BEGIN_BRANCH(TRIX, "TRIX")
			GETN_NUM(Period1, "EMA Period", 15)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "Signal Period", 9)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(TRIX)
		
		GETN_BEGIN_BRANCH(TSI, "True Strength Index (TSI)")
			GETN_NUM(Period1, "1st smooth Period", 25)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd smooth Period", 13)						GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "Signal Period", 7)									GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(TSI)
		
		GETN_BEGIN_BRANCH(UI, "Ulcer Index")
			GETN_NUM(Period, "Period", 14)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(UI)
		
		GETN_BEGIN_BRANCH(ULT, "Ultimate Oscillator")
			GETN_NUM(Period1, "1st Period", 7)					GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "2nd Period", 14)				GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period3, "3rd Period", 28)				GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(ULT)
		
		GETN_BEGIN_BRANCH(VI, "Vortex Indicator")
			GETN_NUM(Period, "Period", 14)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(VI)
		
		GETN_BEGIN_BRANCH(WmR, "William %R")
			GETN_NUM(Period, "Period", 14)						GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(WmR)
		
		/*
		GETN_BEGIN_BRANCH(McClellan, "McClellan Oscillator")
			GETN_NUM(Period1, "Period Window 1", 19)		GETN_CURRENT_SUBNODE.Enable = false;
			GETN_OPTION_NUM_FORMAT( ".0" )
			GETN_NUM(Period2, "Period Window 2", 39)		GETN_CURRENT_SUBNODE.Enable = false;
			GETN_OPTION_NUM_FORMAT( ".0" )
		GETN_END_BRANCH(McClellan)
		*/
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
	
	// KST
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_KST_Main(wksIndicator, wksData, vsIndex, trIndicators.KST);
	
	// SPECIALK
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_SPECIALK_Main(wksIndicator, wksData, vsIndex, trIndicators.SPECIALK);
	
	// RSI
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 3);	
	StockAnalyzer_RSI_Main(wksIndicator, wksData, vsIndex, trIndicators.RSI);
	
	// Slope
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_Slope_Main(wksIndicator, wksData, vsIndex, trIndicators.Slope);
	
	// StdDev
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_StdDev_Main(wksIndicator, wksData, vsIndex, trIndicators.StdDev);
	
	// Stochastic Oscillator (K & D)
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_Sto_Main(wksIndicator, wksData, vsIndex, trIndicators.StochasticOscillator);
	
	// StochRSI
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_StochRSI_Main(wksIndicator, wksIndicator, vsIndex, trIndicators.RSI);
	
	// TRIX
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_TRIX_Main(wksIndicator, wksData, vsIndex, trIndicators.TRIX);
	
	// True Strength Index (TSI)
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_TSI_Main(wksIndicator, wksData, vsIndex, trIndicators.TSI);
	
	// Ulcer Index
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_UI_Main(wksIndicator, wksData, vsIndex, trIndicators.UI);
	
	// Ultimate Oscillator
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_ULT_Main(wksIndicator, wksData, vsIndex, trIndicators.ULT);
	
	// Vortex Indicator
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 2);	
	StockAnalyzer_VI_Main(wksIndicator, wksData, vsIndex, trIndicators.VI);
	
	// William %R
	vsIndex = StockAnalyzer_Offset(nIndexIndicator, 1);	
	StockAnalyzer_WilliamR_Main(wksIndicator, wksData, vsIndex, trIndicators.WmR);
	
	// Chandelier Exit
	vsIndex = StockAnalyzer_Offset(nIndexOverlay, 2);
	StockAnalyzer_ChandelierExit_Main(wksOverLay, wksData, vsIndex, trOverlays.Chandlr);
	
	// Ichimoku Clouds
	vsIndex = StockAnalyzer_Offset(nIndexOverlay, 5);
	StockAnalyzer_IchimokuCloud_Main(wksOverLay, wksData, vsIndex, trOverlays.Ichimoku);
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
	vsEMA[0] = is_missing_value(vsClose[0])? 0:vsClose[0];
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
	vsADX[0] = 20;
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
	vector<double> vsROC(nSize);	vsROC = 0;
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
	//vsPreviousClose.InsertAt(0, NANUM);
	vsPreviousClose.InsertAt(0, vsClose[0]);		// 1st force bing 0 for setup
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

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Know Sure Thing (KST)

static void StockAnalyzer_KST_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nP1 = tr.GetNode("ROCp1").nVal;
	int nW1 = tr.GetNode("ROCw1").nVal;
	int nP2 = tr.GetNode("ROCp2").nVal;
	int nW2 = tr.GetNode("ROCw2").nVal;
	int nP3 = tr.GetNode("ROCp3").nVal;
	int nW3 = tr.GetNode("ROCw3").nVal;
	int nP4 = tr.GetNode("ROCp4").nVal;
	int nW4 = tr.GetNode("ROCw4").nVal;
	int nPS = tr.GetNode("SignalPeriod").nVal;
	
	string strLegend;
	strLegend.Format("KST(%d, %d,%d,%d,%d,%d,%d,%d,%d)", nP1, nP2, nP3, nP4, nW1, nW2, nW3, nW4, nPS);
	Dataset dsKST(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Know Sure Thing", strLegend));
	Dataset dsSignal(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Know Sure Thing", ""));
	
	dsKST = StockAnalyzer_KST(dsClose, nP1, nP2, nP3, nP4, nW1, nW2, nW3, nW4);
	dsSignal = StockAnalyzer_SMA(dsKST, nPS);
}

static vector<double> StockAnalyzer_KST(vector<double> vsClose, int nP1, int nP2, int nP3, int nP4, int nW1, int nW2, int nW3, int nW4) {
	vector<double> vsKST;
	
	vector<double> vsRCMA1, vsRCMA2, vsRCMA3, vsRCMA4;
	vsRCMA1 = StockAnalyzer_RCMA(vsClose, nP1, nW1);
	vsRCMA2 = StockAnalyzer_RCMA(vsClose, nP2, nW2);
	vsRCMA3 = StockAnalyzer_RCMA(vsClose, nP3, nW3);
	vsRCMA4 = StockAnalyzer_RCMA(vsClose, nP4, nW4);
	vsKST = vsRCMA1 + 2 * vsRCMA2 + 3 * vsRCMA3 + 4 * vsRCMA4;
	
	return vsKST;
}

static vector<double> StockAnalyzer_RCMA(vector<double> vsClose, int nP, int nW) {
	vector<double> vsRCMA;
	
	vsRCMA = StockAnalyzer_ROC(vsClose, nP);
	vsRCMA = StockAnalyzer_SMA(vsRCMA, nW);
	
	return vsRCMA;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Martin Pring's Special K

static void StockAnalyzer_SPECIALK_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriodSMA = tr.GetNode("Period1").nVal;	//
	int nPeriodSMASmooth = tr.GetNode("Period2").nVal;	// 
	
	Dataset dsSPECIALK(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Martin Pring's Special K", StockAnalyzer_Legend("SPECIALK", nPeriodSMA, nPeriodSMASmooth)));
	Dataset dsSmooth(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Martin Pring's Special K", ""));
	
	dsSPECIALK = StockAnalyzer_SPECIALK(dsClose);
	dsSmooth = StockAnalyzer_SPECIALK_Smooth(dsSPECIALK, nPeriodSMA, nPeriodSMASmooth);
}

static vector<double> StockAnalyzer_SPECIALK(vector<double> vsClose) {
	vector<double> vsSpecialK;
	int nSize = vsClose.GetSize();
	vsSpecialK.SetSize(nSize);
	vsSpecialK = 0;
	
	if (nSize < 725) {
		return vsSpecialK;
	};
	
	vector<int> vsListROC = {10, 15, 20, 30, 40, 65, 75, 100, 198, 265, 390, 530};
	vector<int> vsListWin = {10, 10, 10, 15, 50, 65, 75, 100, 130, 130, 130, 195};
	vector<int> vsListMulti = {1, 2, 3, 4, 1, 2, 3, 4, 1, 2, 3, 4};
	
	for (int ii = 0; ii < 12; ii++) {
		vector<double> vsROC;
		vsROC = StockAnalyzer_ROC(vsClose, vsListROC[ii]);
		vsSpecialK += StockAnalyzer_SMA(vsROC, vsListWin[ii]) * vsListMulti[ii];
	};
	
	return vsSpecialK;
}

static vector<double> StockAnalyzer_SPECIALK_Smooth(vector<double> vsK, int nPeriodSMA, int nPeriodSMASmooth) {
	vector<double> vsSmooth;
	
	vsSmooth = StockAnalyzer_SMA(vsK, nPeriodSMA);
	
	return vsSmooth;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Relative Strength Index (RSI)

static void StockAnalyzer_RSI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod = tr.GetNode("Period").nVal;
	int nSize = dsClose.GetSize();
	
	Dataset ds70(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Relative Strength Index", ""));
	Dataset dsRSI(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Relative Strength Index", StockAnalyzer_Legend("RSI", nPeriod)));
	Dataset ds30(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Relative Strength Index", ""));
	
	ds70.SetSize(nSize); 	ds70 = 70;
	dsRSI = StockAnalyzer_RSI(dsClose, nPeriod);
	ds30.SetSize(nSize); 	ds30 = 30;
}

static vector<double> StockAnalyzer_RSI(vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsRSI(nSize), vsRS(nSize);
	
	vector<double> vsChange;
	vsChange = StockAnalyzer_Change(vsClose);	// Get Change
	
	vector<double> vsGain(nSize), vsLoss(nSize);
	StockAnalyzer_SeperateGainAndLoss(vsChange, vsGain, vsLoss);		// Seperate gain and loss form change
	
	vsRS = StockAnalyzer_RS(vsGain, vsLoss, nPeriod);		// RS
	
	vsRSI = StockAnalyzer_RSI(vsRS);		// RSI
	
	return vsRSI;
}

static vector<double> StockAnalyzer_Change(vector<double> vsSource) {
	vector<double> vsChange;
	vsSource.Difference(vsChange);
	vsChange.InsertAt(0, 0);
	return vsChange;
}

static bool StockAnalyzer_SeperateGainAndLoss(vector<double> vsChange, vector<double> &vsGain, vector<double> &vsLoss) {
	for (int ii = 0; ii < vsChange.GetSize(); ii++) {
		if (vsChange[ii] > 0) {
			vsGain[ii] = vsChange[ii];
			vsLoss[ii] = NANUM;
		}
		else if (vsChange[ii] < 0) {
			vsGain[ii] = NANUM;
			vsLoss[ii] = abs(vsChange[ii]);
		}
		else {
			vsGain[ii] = NANUM;
			vsLoss[ii] = NANUM;
		};
	};
	return true;
}

static vector<double> StockAnalyzer_RS(vector<double> vsGain, vector<double> vsLoss, int nPeriod) {
	
	// The very first calculations for average gain and average loss are simple N period averages.
	int nSize = vsGain.GetSize();
	vector<double> vsMovingGain(nSize), vsMovingLoss(nSize);
	//vsMovingGain = StockAnalyzer_SMA(vsGain, nPeriod);
	//vsMovingLoss = StockAnalyzer_SMA(vsLoss, nPeriod);
	vsMovingGain[nPeriod] = StockAnalyzer_RS_First(vsGain, nPeriod);
	vsMovingLoss[nPeriod] = StockAnalyzer_RS_First(vsLoss, nPeriod);
	
	// The second, and subsequent, calculations are based on the prior averages and the current gain loss:
	for (int ii = nPeriod + 1; ii < nSize; ii++) {
		if (is_missing_value(vsGain[ii])) {
			vsMovingGain[ii] = vsMovingGain[ii-1] * (nPeriod - 1) / nPeriod;
		}
		else {
			vsMovingGain[ii] = ((vsMovingGain[ii-1] * (nPeriod - 1)) + vsGain[ii]) / nPeriod;
		}
		
		if (is_missing_value(vsLoss[ii])) {
			vsMovingLoss[ii] = vsMovingLoss[ii-1] * (nPeriod - 1) / nPeriod;
		}
		else {
			vsMovingLoss[ii] = ((vsMovingLoss[ii-1] * (nPeriod - 1)) + vsLoss[ii]) / nPeriod;
		}
	};
	
	//return vsRS;
	return vsMovingGain / vsMovingLoss;
}

static vector<double> StockAnalyzer_RSI(vector<double> vsRS) {
	return 100 - 100 / (1 + vsRS);
}

static double StockAnalyzer_RS_First(vector<double> vs, int nPeriod) {
	vector<double> vsSubRange;
	vs.GetSubVector(vsSubRange, 0, nPeriod-1);
	double dSum;
	vsSubRange.Sum(dSum);
	return dSum / nPeriod;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Slope

static void StockAnalyzer_Slope_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsSlope(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Slope", StockAnalyzer_Legend("Slope", nPeriod)));
	dsSlope = StockAnalyzer_Slope(dsClose, nPeriod);
}

static vector<double> StockAnalyzer_Slope(vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsSlope(nSize);
	vsSlope = NANUM;
	
	for (int ii = 10; ii < nSize; ii++) {					// Start form 10 to reduce the starting noise
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		int nNum = nEnd - nStart + 1;
		vector<double> vx, vy;
		vx.Data(1, nNum, 1);
		vsClose.GetSubVector(vy, nStart, nEnd);
		FitParameter sFitParameter[2];
		ocmath_linear_fit(vx, vy, nNum, sFitParameter);
		vsSlope[ii] = sFitParameter[1].Value;
	}
	
	return vsSlope;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Standard Deviation (Volatility)

static void StockAnalyzer_StdDev_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsStdDev(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Standard Deviation", StockAnalyzer_Legend("STDDEV", nPeriod)));
	dsStdDev = StockAnalyzer_StdDev(dsClose, nPeriod);
}

static vector<double> StockAnalyzer_StdDev(vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsStdDev(nSize);
	vsStdDev = NANUM;
	
	for (int ii = 10; ii < nSize; ii++) {					// Start form 10 to reduce the starting noise
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		int nNum = nEnd - nStart + 1;
		vector<double> vs;
		vsClose.GetSubVector(vs, nStart, nEnd);
		double  dSD;
		ocmath_basic_summary_stats(nNum, vs, NULL, NULL, &dSD);
		vsStdDev[ii] = dSD;
	}
	
	
	return vsStdDev;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Stochastic Oscillator

static void StockAnalyzer_Sto_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// real close
	
	int nPeriodK = tr.GetNode("PeriodK").nVal;
	int nPeriodD = tr.GetNode("PeriodD").nVal;
	
	Dataset dsK(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Stochastic Oscillator", StockAnalyzer_Legend("%K", nPeriodK)));
	Dataset dsD(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Stochastic Oscillator", StockAnalyzer_Legend("%D", nPeriodD)));
	
	dsK = StockAnalyzer_Sto_K(dsHigh, dsLow, dsClose, nPeriodK);
	
	switch (tr.Method.nVal) {
	case 0:		// Fast
		dsD = StockAnalyzer_SMA(dsK, 3);
		break;
	case 1:		// Slow
		dsK = StockAnalyzer_SMA(dsK, 3);
		dsD = StockAnalyzer_SMA(dsK, 3);
		break;
	case 2:		// Full
		break;
	}
	
}

static vector<double> StockAnalyzer_Sto_K(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsK(nSize);
	vsK = NANUM;
	
	for (int ii = 1; ii < nSize; ii++) {					// Start form 2nd point to reduce the starting noise
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		vector<double> vs;
		vsClose.GetSubVector(vs, nStart, nEnd);
		vsK[ii] = StockAnalyzer_StochasticOscillator(max(vs), min(vs), vsClose[ii]);
	};
	
	return vsK;
}

static double StockAnalyzer_StochasticOscillator(double dHighest, double dLowest, double dClose) {
	return (dClose - dLowest) / (dHighest - dLowest) * 100;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// StochRSI

static void StockAnalyzer_StochRSI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsRSI(wksSource, 45);
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsStochRSI(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "StochRSI", StockAnalyzer_Legend("StochRSI", nPeriod)));
	
	dsStochRSI = StockAnalyzer_StochRSI(dsRSI, nPeriod);
}

static vector<double> StockAnalyzer_StochRSI(vector<double> vsRSI, int nPeriod) {
	int nSize = vsRSI.GetSize();
	vector<double> vsStochRSI(nSize);
	vsStochRSI = NANUM;
	
	for (int ii = 1; ii < nSize; ii++) {
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		vector<double> vs;
		vsRSI.GetSubVector(vs, nStart, nEnd);
		vsStochRSI[ii] = (vsRSI[ii] - min(vs)) / (max(vs) - min(vs));
	};
	
	return vsStochRSI;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// TRIX

static void StockAnalyzer_TRIX_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;
	int nPeriod2 = tr.GetNode("Period2").nVal;
	
	Dataset dsTRIX(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "TRIX", StockAnalyzer_Legend("TRIX", nPeriod1, nPeriod2)));
	Dataset dsSignal(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "TRIX", ""));
	
	dsTRIX = StockAnalyzer_TRIX(dsClose, nPeriod1);
	dsSignal = StockAnalyzer_EMA(dsTRIX, nPeriod2);
}

static vector<double> StockAnalyzer_TRIX(vector<double> vsClose, int nPeriod) {
	vector<double> vsTRIX;
	
	vsTRIX = StockAnalyzer_EMA(vsClose, nPeriod);		// Single-Smoothed EMA
	vsTRIX = StockAnalyzer_EMA(vsTRIX, nPeriod);		// Double-Smoothed EMA
	vsTRIX = StockAnalyzer_EMA(vsTRIX, nPeriod);		// Triple-Smoothed EMA
	vsTRIX = StockAnalyzer_ROC(vsTRIX, 1);
	
	return vsTRIX;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// True Strength Index (TSI)

static void StockAnalyzer_TSI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;
	int nPeriod2 = tr.GetNode("Period2").nVal;
	int nPeriod3 = tr.GetNode("Period3").nVal;
	
	Dataset dsTSI(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "True Strength Index", StockAnalyzer_Legend("TSI", nPeriod1, nPeriod2, nPeriod3)));
	Dataset dsSignal(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "True Strength Index", ""));
	
	dsTSI = StockAnalyzer_TSI(dsClose, nPeriod1, nPeriod2);
	dsSignal = StockAnalyzer_EMA(dsTSI, nPeriod3);
}

static vector<double> StockAnalyzer_TSI(vector<double> vsClose, int nPeriod1, int nPeriod2) {
	vector<double> vsTSI, vsPC, vsAPC;
	
	// Price Change
	vsClose.Difference(vsPC);
	vsPC.InsertAt(0, 0);
	vsAPC = abs(vsPC);
	
	// 1st Smooth
	vsPC = StockAnalyzer_EMA(vsPC, nPeriod1);
	vsAPC = StockAnalyzer_EMA(vsAPC, nPeriod1);
	
	// 2nd Smooth
	vsPC = StockAnalyzer_EMA(vsPC, nPeriod2);
	vsAPC = StockAnalyzer_EMA(vsAPC, nPeriod2);
	
	vsTSI = 100 * (vsPC / vsAPC);
	
	return vsTSI;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Ulcer Index

static void StockAnalyzer_UI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsUI(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Ulcer Index", StockAnalyzer_Legend("UI", nPeriod)));
	
	dsUI = StockAnalyzer_UI(dsClose, nPeriod);
}

static vector<double> StockAnalyzer_UI(vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsUI(nSize);
	
	vector<double> vsPercentDrawdown(nSize), vsSquaredAverage(nSize);
	for (int ii = 1; ii < nSize; ii++) {
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		int nNum = nEnd - nStart + 1;
		vector<double> vs;
		vsClose.GetSubVector(vs, nStart, nEnd);
		vsPercentDrawdown[ii] = (vsClose[ii] - max(vs)) / max(vs) * 100;
		vsPercentDrawdown.GetSubVector(vs, nStart, nEnd);
		vs *=vs;
		vs.Sum(vsSquaredAverage[ii]); 
		vsSquaredAverage[ii] = vsSquaredAverage[ii] / nNum;
	};
	vsUI = sqrt(vsSquaredAverage);
	
	return vsUI;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Ultimate Oscillator

static void StockAnalyzer_ULT_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsClose(wksSource, 6);	// adj close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;
	int nPeriod2 = tr.GetNode("Period2").nVal;
	int nPeriod3 = tr.GetNode("Period3").nVal;
	
	Dataset dsULT(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Ultimate Oscillator", StockAnalyzer_Legend("ULT", nPeriod1, nPeriod2, nPeriod3)));
	
	dsULT = StockAnalyzer_ULT(dsClose, nPeriod1, nPeriod2, nPeriod3);
}

static vector<double> StockAnalyzer_ULT(vector<double> vsClose, int nPeriod1, int nPeriod2, int nPeriod3) {
	int nSize = vsClose.GetSize();
	vector<double> vsULT(nSize);
	vector<double> vsBP(nSize), vsTR(nSize), vsAvg1(nSize), vsAvg2(nSize), vsAvg3(nSize);
	
	vector<double> vsTemp;
	vsClose.GetSubVector(vsTemp, 0, 15);
	double dLow = min(vsTemp);
	double dHigh = max(vsTemp);
	
	for (int ii = 16; ii < nSize; ii++) {
		dLow = min(dLow, vsClose[ii-1]);
		dHigh = max(dHigh, vsClose[ii-1]);
		vsBP[ii] = vsClose[ii] - dLow;
		vsTR[ii] = dHigh - dLow;
	};
	
	vsAvg1 = StockAnalyzer_MovingSum(vsBP, nPeriod1) / StockAnalyzer_MovingSum(vsTR, nPeriod1);
	vsAvg2 = StockAnalyzer_MovingSum(vsBP, nPeriod2) / StockAnalyzer_MovingSum(vsTR, nPeriod2);
	vsAvg3 = StockAnalyzer_MovingSum(vsBP, nPeriod3) / StockAnalyzer_MovingSum(vsTR, nPeriod3);
	
	vsULT = 100 * ((4 * vsAvg1) + (2 * vsAvg2) + vsAvg3) / (4 + 2 + 1);
	
	return vsULT;
}

static vector<double> StockAnalyzer_MovingSum(vector<double> vs, int nPeriod) {
	int nSize = vs.GetSize();
	vector<double> vsSum(nSize);		vsSum[0] = vs[0];
	
	for (int ii = 1; ii < nSize; ii++) {
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		vector<double> vsSub;
		vs.GetSubVector(vsSub, nStart, nEnd);
		vsSub.Sum(vsSum[ii]);
	}
	
	return vsSum;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Vortex Indicator

static void StockAnalyzer_VI_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// real close
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsVIp(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Vortex Indicator", StockAnalyzer_Legend("+VI", nPeriod)));
	Dataset dsVIn(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Vortex Indicator", StockAnalyzer_Legend("-VI", nPeriod)));
	
	StockAnalyzer_VI(dsHigh, dsLow, dsClose, nPeriod, dsVIp, dsVIn);
}

static void StockAnalyzer_VI(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, int nPeriod, vector<double> &vsVIp, vector<double> &vsVIn) {
	int nSize = vsClose.GetSize();
	vector<double> vsVMp(nSize), vsVMn(nSize), vsTR(nSize);
	
	for (int ii = 1; ii < nSize; ii++) {
		vsVMp[ii] = abs(vsHigh[ii] - vsLow[ii-1]);
		vsVMn[ii] = abs(vsLow[ii] - vsHigh[ii-1]);
		vsTR[ii] = max(vsHigh[ii] - vsLow[ii], max(abs(vsHigh[ii] - vsClose[ii-1]), abs(vsLow[ii] - vsClose[ii-1])));
	}
	
	vsVMp = StockAnalyzer_MovingSum(vsVMp, nPeriod);
	vsVMn = StockAnalyzer_MovingSum(vsVMn, nPeriod);
	vsTR = StockAnalyzer_MovingSum(vsTR, nPeriod);
	
	vsVIp = vsVMp / vsTR;
	vsVIn = vsVMn / vsTR;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// William %R

static void StockAnalyzer_WilliamR_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// real close
	
	int nPeriod = tr.GetNode("Period").nVal;
	
	Dataset dsWmR(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "William", StockAnalyzer_Legend("Wm%R", nPeriod)));
	
	dsWmR = StockAnalyzer_WilliamR(dsHigh, dsLow, dsClose, nPeriod);
}

static vector<double> StockAnalyzer_WilliamR(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, int nPeriod) {
	int nSize = vsClose.GetSize();
	vector<double> vsWmR(nSize);
	
	for (int ii = 1; ii < nSize; ii++) {
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		vector<double> vsSubHigh, vsSubLow;
		vsHigh.GetSubVector(vsSubHigh, nStart, nEnd);
		vsLow.GetSubVector(vsSubLow, nStart, nEnd);
		vsWmR[ii] = (max(vsSubHigh) - vsClose[ii]) / (max(vsSubHigh) - min(vsSubLow)) * -100;
	}
	
	return vsWmR;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Chandelier Exit

static void StockAnalyzer_ChandelierExit_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// real close
	
	int nPeriod = tr.GetNode("Period").nVal;
	double nFactor = tr.GetNode("Factor").nVal;
	
	Dataset dsLong(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Chandelier Exit", StockAnalyzer_Legend("CHANDLR_Long", nPeriod, nFactor)));
	Dataset dsShort(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Chandelier Exit", StockAnalyzer_Legend("CHANDLR_Short", nPeriod, nFactor)));
	
	StockAnalyzer_ChandelierExit(dsHigh, dsLow, dsClose, nPeriod, nFactor, dsLong, dsShort);
}

static void StockAnalyzer_ChandelierExit(vector<double> vsHigh, vector<double> vsLow, vector<double> vsClose, int nPeriod, double nFactor, vector<double> &vsLong, vector<double> &vsShort) {
	int nSize = vsClose.GetSize();
	vsLong.SetSize(nSize);
	vsShort.SetSize(nSize);
	
	vector<double> vsATR(nSize);
	vsATR = StockAnalyzer_ATR(vsClose, vsHigh, vsLow, nPeriod);
	
	for (int ii = 0; ii < nSize; ii++) {
		int nStart = (ii < nPeriod)? 0:(ii - nPeriod + 1);
		int nEnd = ii;
		vector<double> vsSubHigh, vsSubLow;
		vsHigh.GetSubVector(vsSubHigh, nStart, nEnd);
		vsLow.GetSubVector(vsSubLow, nStart, nEnd);
		
		vsLong[ii] = max(vsSubHigh) - vsATR[ii] * nFactor;
		vsShort[ii] = min(vsSubLow) + vsATR[ii] * nFactor;
	};
	
	
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Ichimoku Clouds

static void StockAnalyzer_IchimokuCloud_Main(Worksheet wksTarget, Worksheet wksSource, vector<int> nIndex, Tree tr) {
	Dataset dsDate(wksSource, 0);		// 	Date(X)
	Dataset dsHigh(wksSource, 2);		// 
	Dataset dsLow(wksSource, 3);		// 
	Dataset dsClose(wksSource, 4);	// real close
	
	int nPeriod1 = tr.GetNode("Period1").nVal;
	int nPeriod2 = tr.GetNode("Period2").nVal;
	int nPeriod3 = tr.GetNode("Period3").nVal;
	
	Dataset dsTenkan(StockAnalyzer_SetColumn(wksTarget, nIndex[0], "Ichimoku Clouds", "Conversion Line"));
	Dataset dsKijun(StockAnalyzer_SetColumn(wksTarget, nIndex[1], "Ichimoku Clouds", "Base Line"));
	Dataset dsSenkouA(StockAnalyzer_SetColumn(wksTarget, nIndex[2], "Ichimoku Clouds", "Leading Span A"));
	Dataset dsSenkouB(StockAnalyzer_SetColumn(wksTarget, nIndex[3], "Ichimoku Clouds", "Leading Span B"));
	Dataset dsChikou(StockAnalyzer_SetColumn(wksTarget, nIndex[4], "Ichimoku Clouds", "Lagging Span"));
	
	dsTenkan = StockAnalyzer_IchimokuCloud_Span(dsHigh, dsLow, nPeriod1);
	dsKijun = StockAnalyzer_IchimokuCloud_Span(dsHigh, dsLow, nPeriod2);
	dsSenkouA = StockAnalyzer_Offset((dsTenkan + dsKijun) / 2, dsDate, 0);
	dsSenkouB = StockAnalyzer_Offset(StockAnalyzer_IchimokuCloud_Span(dsHigh, dsLow, nPeriod3), dsDate, 0);
	dsChikou = StockAnalyzer_Offset(dsClose, dsDate, 26);
}

static vector<double> StockAnalyzer_IchimokuCloud_Span(vector<double> vsHigh, vector<double> vsLow, int nPeriod) {
	//<double> vsSpan;
	return (StockAnalyzer_SMA(vsHigh, nPeriod) + StockAnalyzer_SMA(vsLow, nPeriod)) / 2;
}

static vector<double> StockAnalyzer_Offset(vector<double> vs, vector<double> &vsX, int nOffset) {
	// Positive offset shifts curve to earlier
	// Negative offset shifts curve to future
	
	int nSize = vs.GetSize();
	int nSizeX = vsX.GetSize();
	
	// Expand X if needed
	if (nSize - nOffset > nSizeX) {
		double dLast = vsX[nSizeX-1];
		for (int ii = 1; ii <= abs(nOffset); ii++) {
			vsX.Add(dLast + ii);
		};
	}
	
	vector<double> vsNew(max(nSize - nOffset, nSize));
	for (int ii = 0; ii < vsNew.GetSize(); ii++) {
		int nPos = ii + nOffset;
		if (nPos < 0) {
			vsNew[ii] = NANUM;
		}
		else {
			vsNew[ii] = vs[ii + nOffset];
		};
	};
	return vsNew;
}



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Misc

static GraphPage GetStockAnalyzerGraph() {
	WorksheetPage wp = Project.Pages();
	Worksheet wks = wp.Layers(0);
	GraphPage gp = wks.EmbeddedPages(0);
	return gp;
}

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

void SA_OverLay_Type(int nType, GraphPage gp) {
	string strLegendName;
	vector<int> vsIndex;	
	SA_OverLayList(nType, vsIndex, strLegendName);
	
	GraphLayer gl = gp.Layers(0);		// Get 1st layer
	
	SA_OverLay_ShowHide(gl, vsIndex, strLegendName);
	
	gp.Refresh();
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
	case 4:	// Chandelier Exit
		vector<int> vsTemp = {10, 11};
		vsIndex = vsTemp;
		strLegendName = "LegendCHANDLR";
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

void SA_Indicator_Type(int nType, GraphPage gp) {
	vector<string> saList;	saList = SA_IndicatorList();
	//GraphPage gp = Project.ActiveLayer().GetPage();
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
	
	sa.Add("Accumulation Distribution Line");		// 1
	sa.Add("Aroon");
	sa.Add("Aroon Oscillator");
	sa.Add("Average True Range");
	sa.Add("Average Directional Index");
	sa.Add("Bollinger BandWidth");
	sa.Add("B Indicator");
	sa.Add("Commodity Channel Index");
	sa.Add("Coppock Curve");
	sa.Add("Chaikin Money Flow");						// 10
	sa.Add("Chaikin Oscillator");
	sa.Add("PMO");
	sa.Add("DPO");
	sa.Add("Ease of Movement");
	sa.Add("Force Index");
	sa.Add("Mass Index");
	sa.Add("MACD");
	sa.Add("Money Flow Index");
	sa.Add("Negative Volume Index");
	sa.Add("On Balance Volume");						// 20
	sa.Add("PPO");
	sa.Add("PVO");
	sa.Add("Know Sure Thing");
	sa.Add("Special K");
	sa.Add("Rate of Change");
	sa.Add("Relative Strength Index");
	sa.Add("Slope");
	sa.Add("StdDev");
	sa.Add("Sto");
	sa.Add("StochRSI");														// 30
	sa.Add("TRIX");	
	sa.Add("TSI");	
	sa.Add("UI");	
	sa.Add("ULT");	
	sa.Add("Vortex");	
	sa.Add("WilliamR");	
	return sa;
}


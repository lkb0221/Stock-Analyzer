<<<<<<< HEAD
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


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main

void StockAnalyzer_Main() {
	Tree tr;
	tr = StockAnalyzer_GUI();
	
	
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DIalog GUI

Tree StockAnalyzer_GUI() {
	double dToday = StockAnalyzer_GetToday();
	
	GETN_BOX(tr)
	
	GETN_BEGIN_BRANCH(Data, "Data")
	GETN_OPTION_BRANCH(GETNBRANCH_OPEN)
		GETN_STR(SymbolName, "Stock Name", "AAPL")
		GETN_RADIO_INDEX_EX(Range, "Range", 0, "Last Year|Last Three Year|All|Custom")
		GETN_DATE(StartDate, "Date From", dToday-366)
		GETN_DATE(EndDate, "Date To", dToday-1)
	GETN_END_BRANCH(Data)
	
	GETN_BEGIN_BRANCH(Overlays, "Technical Overlays")
		GETN_BEGIN_BRANCH(MA, "Moving Averages - Simple and Exponential")
			GETN_NUM(Period1, "1st Period", 20)
			GETN_NUM(Period2, "2nd Period", 50)
			GETN_NUM(Period3, "3rd Period", 200)
		GETN_END_BRANCH(MA)
		
		
	GETN_END_BRANCH(Overlays)
	
	GETN_BEGIN_BRANCH(Indicators, "Technical Indicators")
		
	GETN_END_BRANCH(Indicators)
	
	if( GetNBox(tr, "Stock Analyzer") ) {
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

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Download csv form Yahoo API

static string StockAnalyzer_Download(string strSymbolName, double dStartDate, double dEndDate) {
	// Get day, month, year integers
	vector<int> vsStart(3), vsEnd(3);
	vsStart = StockAnalyzer_Parse_Date(dStartDate);
	vsEnd = StockAnalyzer_Parse_Date(dEndDate);
	
	// Build url string
	string strURL;
	strURL.Format(STR_YAHOO_API_BASE, strSymbolName, vsStart[1], vsStart[2], vsStart[0], vsEnd[1], vsEnd[2], vsEnd[0]);
	
	// Biuld Local File Path
	string strFilePath = __FILE__;
	strFilePath = GetFilePath(strFilePath) + "Table.csv";
	
	// Download
	bool bGet = http_get_file(strURL, strFilePath);
	
	if (bGet) {
		return strFilePath;
	}
	else {
		return NULL;
	}
	
	//out_str(strFilePath);
}

static vector<int> StockAnalyzer_Parse_Date(double dDate/*, int& nYear, int& nMonth, int& nDay*/) {
	vector<int> vs(3);	// {Year, Month, Day}
	
	SYSTEMTIME st;
	JulianDateToSystemTime(&dDate, &st);
	vs[0] = st.wYear;	// year
	vs[1] = st.wMonth - 1;	// month
	vs[2] = st.wDay;		// day
	
	return vs;
=======
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
#include <..\OriginLab\OCHttpUtils.h>
////////////////////////////////////////////////////////////////////////////////////
#define STR_YAHOO_API_BASE "http://ichart.finance.yahoo.com/table.csv?s=%s&a=%d&b=%d&c=%d&d=%d&e=%d&f=%d&g=d&ignore=.csv"


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main

void StockAnalyzer_Main() {
	
	
	
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Download csv form Yahoo API

static string StockAnalyzer_Download(string strSymbolName, double dStartDate, double dEndDate) {
	// Get day, month, year integers
	vector<int> vsStart(3), vsEnd(3);
	vsStart = StockAnalyzer_Parse_Date(dStartDate);
	vsEnd = StockAnalyzer_Parse_Date(dEndDate);
	
	// Build url string
	string strURL;
	strURL.Format(STR_YAHOO_API_BASE, strSymbolName, vsStart[1], vsStart[2], vsStart[0], vsEnd[1], vsEnd[2], vsEnd[0]);
	
	// Biuld Local File Path
	string strFilePath = __FILE__;
	strFilePath = GetFilePath(strFilePath) + "Table.csv";
	
	// Download
	bool bGet = http_get_file(strURL, strFilePath);
	
	if (bGet) {
		return strFilePath;
	}
	else {
		return NULL;
	}
	
	//out_str(strFilePath);
}

static vector<int> StockAnalyzer_Parse_Date(double dDate/*, int& nYear, int& nMonth, int& nDay*/) {
	vector<int> vs(3);	// {Year, Month, Day}
	
	SYSTEMTIME st;
	JulianDateToSystemTime(&dDate, &st);
	vs[0] = st.wYear;	// year
	vs[1] = st.wMonth - 1;	// month
	vs[2] = st.wDay;		// day
	
	return vs;
>>>>>>> refs/remotes/origin/master
}
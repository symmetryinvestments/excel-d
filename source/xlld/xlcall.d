/**
    Microsoft Excel Developer's Toolkit
    Version 15.0

    File:           INCLUDE\XLCALL.H
    Description:    Header file for for Excel callbacks
    import std.c.windows.windows;import std.c.windows.windows;latform:       Microsoft Windows

    DEPENDENCY:
    Include <windows.h> before you include this.

    This file defines the constants and
    data types which are used in the
    Microsoft Excel C API.

  	Ported to the D Programming Language by Laeeth Isharc (2015)
*/
module xlld.xlcall;

version(Windows) {
    import core.sys.windows.windows;
    version (UNICODE) {
        static assert(false, "Unicode not supported right now");
    } else {
        import core.sys.windows.winnt: LPSTR;
    }
} else version(unittest) {
    alias HANDLE = int;
    alias VOID = void;
    alias HWND = int;
    alias POINT = int;
    alias LPSTR = wchar*;
}

/**
   XL 12 Basic Datatypes
 */
extern(System) int Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER* opers); //pascal
extern(C)int Excel4(int xlfn, LPXLOPER operRes, int count,... );  //_cdecl

	alias BYTE=ubyte;
	alias WORD=ushort;
	alias DWORD=uint;				// guess
	alias DWORD_PTR=DWORD*;			// guess
	alias BOOL=int;
	alias XCHAR=wchar;
	alias RW=int;					// XL 12 Row
	alias COL=int;					// XL 12 Column
	alias IDSHEET=DWORD_PTR;		// XL12 Sheet ID

	/**
	   XLREF structure
	   Describes a single rectangular reference.
	*/

	struct XLREF
	{
		WORD rwFirst;
		WORD rwLast;
		BYTE colFirst;
		BYTE colLast;
	}

	alias LPXLREF=XLREF*;


	/**
	   XLMREF structure
	   Describes multiple rectangular references.
	   This is a variable size structure, default
	   size is 1 reference.
	*/

	struct XLMREF
	{
		WORD count;
		XLREF* reftbl;					/* actually reftbl[count] */
	}

	alias LPXLMREF=XLMREF*;


	/**
	   XLREF12 structure

	   Describes a single XL 12 rectangular reference.
	*/

	struct XLREF12
	{
		RW rwFirst;
		RW rwLast;
		COL colFirst;
		COL colLast;
	}
	alias LPXLREF12=XLREF12*;


	/**
	   XLMREF12 structure

	   Describes multiple rectangular XL 12 references.
	   This is a variable size structure, default
	   size is 1 reference.
	*/

	struct XLMREF12
	{
		WORD count;
		XLREF12* reftbl;					/* actually reftbl[count] */
	}
	alias LPXLMREF12=XLMREF12*;


	/**
	   FP structure

	   Describes FP structure.
	*/

	struct FP
	{
	    ushort rows;
	    ushort columns;
	    double* array;        /* Actually, array[rows][columns] */
	}

	/**
	   FP12 structure

	   Describes FP structure capable of handling the big grid.
	*/

	struct FP12
	{
	    int rows;
	    int columns;
	    double* array;        /* Actually, array[rows][columns] */
	}


	/**
	   XLOPER structure

	   Excel's fundamental data type: can hold data
	   of any type. Use "R" as the argument type in the
	   REGISTER function.
	 */
	struct XLOPER
	{
		union VAL
		{
			double num;
			LPSTR str;					/* xltypeStr */
			WORD bool_;					/* xltypeBool */
			WORD err;
			short w;					/* xltypeInt */
			struct SREF
			{
				WORD count;				/* always = 1 */
				XLREF ref_;
			}
			SREF sref;						/* xltypeSRef */
			struct MREF
			{
				XLMREF *lpmref;
				IDSHEET idSheet;
			}
			MREF mref;						/* xltypeRef */
			struct ARRAY
			{
				XLOPER *lparray;
				WORD rows;
				WORD columns;
			}
			ARRAY array;					/* xltypeMulti */
			struct FLOW
			{
				union VALFLOW
				{
					short level;		/* xlflowRestart */
					short tbctrl;		/* xlflowPause */
					IDSHEET idSheet;		/* xlflowGoto */
				}
				VALFLOW valflow;
				WORD rw;				/* xlflowGoto */
				BYTE col;				/* xlflowGoto */
				BYTE xlflow;
			}
			FLOW flow;						/* xltypeFlow */
			struct BIGDATA
			{
				union H
				{
					BYTE *lpbData;			/* data passed to XL */
					HANDLE hdata;			/* data returned from XL */
				}
				H h;
				long cbData;
			}
			BIGDATA bigdata;					/* xltypeBigData */
		}
		VAL val;
		WORD xltype;
	}
	alias LPXLOPER=XLOPER* ;

	/**
	   XLOPER12 structure

	   Excel 12's fundamental data type: can hold data
	   of any type. Use "U" as the argument type in the
	   REGISTER function.
	 */

	struct XLOPER12
	{
		union VAL
		{
			double num;				       	/* xltypeNum */
			XCHAR *str;				       	/* xltypeStr */
			BOOL bool_;				       	/* xltypeBool */
			int err;				       	/* xltypeErr */
			int w;
			struct SREF
			{
				WORD count;			       	/* always = 1 */
				XLREF12 ref_;
			}
			SREF sref;						/* xltypeSRef */
			struct MREF
			{
				XLMREF12 *lpmref;
				IDSHEET idSheet;
			}
			MREF mref;						/* xltypeRef */
			struct ARRAY
			{
				XLOPER12 *lparray;
				RW rows;
				COL columns;
			}
			ARRAY array;					/* xltypeMulti */
			struct FLOW
			{
				union VALFLOW
				{
					int level;			/* xlflowRestart */
					int tbctrl;			/* xlflowPause */
					IDSHEET idSheet;		/* xlflowGoto */
				}
				VALFLOW valflow;
				RW rw;				       	/* xlflowGoto */
				COL col;			       	/* xlflowGoto */
				BYTE xlflow;
			}
			FLOW flow;						/* xltypeFlow */
			struct BIGDATA
			{
				union H
				{
					BYTE *lpbData;			/* data passed to XL */
					HANDLE hdata;			/* data returned from XL */
				}
				H h;
				long cbData;
			}
			BIGDATA bigdata;					/* xltypeBigData */
		}
		VAL val;
		XlType xltype;
	}
	alias LPXLOPER12=XLOPER12*;

	/**
	   XLOPER and XLOPER12 data types

	   Used for xltype field of XLOPER and XLOPER12 structures
	*/

// single enums for XLOPER, proper enum for XLOPER12
	enum xltypeNum=        0x0001;
	enum xltypeStr=       0x0002;
	enum xltypeBool=      0x0004;
	enum xltypeRef=       0x0008;
	enum xltypeErr=       0x0010;
	enum xltypeFlow=      0x0020;
	enum xltypeMulti=     0x0040;
	enum xltypeMissing=   0x0080;
	enum xltypeNil=       0x0100;
	enum xltypeSRef=      0x0400;
	enum xltypeInt=       0x0800;

	enum xlbitXLFree=     0x1000;
	enum xlbitDLLFree=    0x4000;

	enum xltypeBigData=(xltypeStr| xltypeInt);

	enum XlType {
		xltypeNum=        0x0001,
		xltypeStr=       0x0002,
		xltypeBool=      0x0004,
		xltypeRef=       0x0008,
		xltypeErr=       0x0010,
		xltypeFlow=      0x0020,
		xltypeMulti=     0x0040,
		xltypeMissing=   0x0080,
		xltypeNil=       0x0100,
		xltypeSRef=      0x0400,
		xltypeInt=       0x0800,

		xlbitXLFree=     0x1000,
		xlbitDLLFree=    0x4000,

		xltypeBigData=(xltypeStr| xltypeInt),
    }


	/*
	   Error codes

	   Used for val.err field of XLOPER and XLOPER12 structures
	   when constructing error XLOPERs and XLOPER12s
	*/

	enum xlerrNull=   0;
	enum xlerrDiv0=   7;
	enum xlerrValue=  15;
	enum xlerrRef=    23;
	enum xlerrName=   29;
	enum xlerrNum=    36;
	enum xlerrNA=     42;
	enum xlerrGettingData=43;


	/*
	   Flow data types

	   Used for val.flow.xlflow field of XLOPER and XLOPER12 structures
	   when constructing flow-control XLOPERs and XLOPER12s
	 */

	enum xlflowHalt=      1;
	enum xlflowGoto=      2;
	enum xlflowRestart=   8;
	enum xlflowPause=     16;
	enum xlflowResume=    64;


	/**
	   Return codes

	   These values can be returned from Excel4(), Excel4v(), Excel12() or Excel12v().
	*/

	enum xlretSuccess=       0;    /* success */
	enum xlretAbort=         1;    /* macro halted */
	enum xlretInvXlfn=       2;    /* invalid function number */
	enum xlretInvCount=      4;    /* invalid number of arguments */
	enum xlretInvXloper=     8;    /* invalid OPER structure */
	enum xlretStackOvfl=     16;   /* stack overflow */
	enum xlretFailed=        32;   /* command failed */
	enum xlretUncalced=      64;   /* uncalced cell */
	enum xlretNotThreadSafe= 128;  /* not allowed during multi-threaded calc */
	enum xlretInvAsynchronousContext= 256;  /* invalid asynchronous function handle */
	enum xlretNotClusterSafe= 512;  /* not supported on cluster */


	/**
	   XLL events

	   Passed in to an xlEventRegister call to register a corresponding event.
	*/

	enum xleventCalculationEnded=     1;    /* Fires at the end of calculation */
	enum xleventCalculationCanceled=  2;   /* Fires when calculation is interrupted */

/**
	   Function prototypes
*/
	/* followed by count LPXLOPERs */


	int  XLCallVer(); //pascal

	long LPenHelper(int wCode, VOID *lpv); //pascal

	/* followed by count LPXLOPER12s */

	//int Excel12v(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12* opers); //pasca;


/**
	   Cluster Connector Async Callback
	*/
	// CALLBACK
	alias PXL_HPC_ASYNC_CALLBACK=int function (LPXLOPER12 operAsyncHandle, LPXLOPER12 operReturn);


	/**
	   Cluster connector entry point return codes
	*/

	enum xlHpcRetSuccess=           0;
	enum xlHpcRetSessionIdInvalid= -1;
	enum xlHpcRetCallFailed=       -2;


	/**
	   Function number bits
	*/

	enum xlCommand=   0x8000;
	enum xlSpecial=   0x4000;
	enum xlIntl=      0x2000;
	enum xlPrompt=    0x1000;


	/**
	   Auxiliary function numbers

	   These functions are available only from the C API,
	   not from the Excel macro language.
	*/

	enum xlFree=         (0  | xlSpecial);
	enum xlStack=        (1  | xlSpecial);
	enum xlCoerce=       (2  | xlSpecial);
	enum xlSet=          (3  | xlSpecial);
	enum xlSheetId=      (4  | xlSpecial);
	enum xlSheetNm=      (5  | xlSpecial);
	enum xlAbort=        (6  | xlSpecial);
	enum xlGetInst=      (7  | xlSpecial) /* Returns application's hinstance as an integer value, supported on 32-bit platform only */;
	enum xlGetHwnd=      (8  | xlSpecial);
	enum xlGetName=      (9  | xlSpecial);
	enum xlEnableXLMsgs= (10 | xlSpecial);
	enum xlDisableXLMsgs=(11 | xlSpecial);
	enum xlDefineBinaryName=(12 | xlSpecial);
	enum xlGetBinaryName	=(13| xlSpecial);
	/* GetFooInfo are valid only for calls to LPenHelper */
	enum xlGetFmlaInfo	=(14| xlSpecial);
	enum xlGetMouseInfo	=(15| xlSpecial);
	enum xlAsyncReturn	=(16| xlSpecial)	/*Set return value from an asynchronous function call*/;
	enum xlEventRegister=	(17| xlSpecial);	/*Register an XLL event*/
	enum xlRunningOnCluster=(18| xlSpecial);	/*Returns true if running on Compute Cluster*/
	enum xlGetInstPtr=(19| xlSpecial);	/* Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms */

	/* edit modes */
	enum xlModeReady	=0; //=not in edit mode;
	enum xlModeEnter	=1; //=enter mode;
	enum xlModeEdit		=2; //=edit mode;
	enum xlModePoint	= 4; //=point mode;

	/* document(page) types */
	enum dtNil=0x7f;	// window is not a sheet, macro, chart or basic OR window is not the selected window at idle state
	enum dtSheet=0;// sheet
	enum dtProc= 1;// XLM macro
	enum dtChart=2;// Chart
	enum dtBasic=6;// VBA

	/* hit test codes */
	enum htNone		=0x00;//=none of below;
	enum htClient	=0x01;//=internal for "in the client are", should never see;
	enum htVSplit	=0x02;//=vertical split area with split panes;
	enum htHSplit	=0x03;//=horizontal split area;
	enum htColWidth	=0x04;//=column width adjuster area;
	enum htRwHeight	=0x05;//=row height adjuster area;
	enum htRwColHdr	=0x06;//=the intersection of row and column headers;
	enum htObject	=0x07;//=the body of an object;
	// the following are for size handles of draw objects
	enum htTopLeft	=0x08;
	enum htBotLeft	=0x09;
	enum htLeft		=0x0A;
	enum htTopRight	=0x0B;
	enum htBotRight	=0x0C;
	enum htRight		=0x0D;
	enum htTop		=0x0E;
	enum htBot		=0x0F;
	// end size handles
	enum htRwGut	=0x10;//=row area of outline gutter;
	enum htColGut	=0x11;//=column area of outline gutter;
	enum htTextBox	=0x12;//=body of a text box (where we shouw I-Beam cursor);
	enum htRwLevels	=0x13;//=row levels buttons of outline gutter;
	enum htColLevels	=0x14;//=column levels buttons of outline gutter;
	enum htDman		=0x15;//=the drag/drop handle of the selection;
	enum htDmanFill	=0x16;//=the auto-fill handle of the selection;
	enum htXSplit	=0x17;//=the intersection of the horz & vert pane splits;
	enum htVertex	=0x18;//=a vertex of a polygon draw object;
	enum htAddVtx	=0x19;//=htVertex in add a vertex mode;
	enum htDelVtx	=0x1A;//=htVertex in delete a vertex mode;
	enum htRwHdr		=0x1B;//=row header;
	enum htColHdr	=0x1C;//=column header;
	enum htRwShow	=0x1D;//=Like htRowHeight except means grow a hidden column;
	enum htColShow	=0x1E;//=column version of htRwShow;
	enum htSizing	=0x1F;//=Internal use only;
	enum htSxpivot	=0x20;//=a drag/drop tile in a pivot table;
	enum htTabs		=0x21;//=the sheet paging tabs;
	enum htEdit		=0x22;//=Internal use only;

	struct FMLAINFO
	{
		int wPointMode;	// current edit mode.  0 => rest of struct undefined
		int cch;	// count of characters in formula
		char *lpch;	// pointer to formula characters.  READ ONLY!!!
		int ichFirst;	// char offset to start of selection
		int ichLast;	// char offset to end of selection (may be > cch)
		int ichCaret;	// char offset to blinking caret
	}

	struct MOUSEINFO
	{
		// input section
		HWND hwnd;		// window to get info on
		POINT pt;		// mouse position to get info on

		// output section
		int dt;			// document(page) type
		int ht;			// hit test code
		int rw;			// row @ mouse (-1 if #n/a)
		int col;		// col @ mouse (-1 if #n/a)
	} ;



	/*
	   User defined function

	   First argument should be a function reference.
	*/

	enum xlUDF=     255;


	/**
	   Built-in Excel functions and command equivalents
	*/


	// Excel function numbers

	enum xlfCount=0;
	enum xlfIsna=2;
	enum xlfIserror=3;
	enum xlfSum=4;
	enum xlfAverage=5;
	enum xlfMin=6;
	enum xlfMax=7;
	enum xlfRow=8;
	enum xlfColumn=9;
	enum xlfNa=10;
	enum xlfNpv=11;
	enum xlfStdev=12;
	enum xlfDollar=13;
	enum xlfFixed=14;
	enum xlfSin=15;
	enum xlfCos=16;
	enum xlfTan=17;
	enum xlfAtan=18;
	enum xlfPi=19;
	enum xlfSqrt=20;
	enum xlfExp=21;
	enum xlfLn=22;
	enum xlfLog10=23;
	enum xlfAbs=24;
	enum xlfInt=25;
	enum xlfSign=26;
	enum xlfRound=27;
	enum xlfLookup=28;
	enum xlfIndex=29;
	enum xlfRept=30;
	enum xlfMid=31;
	enum xlfLen=32;
	enum xlfValue=33;
	enum xlfTrue=34;
	enum xlfFalse=35;
	enum xlfAnd=36;
	enum xlfOr=37;
	enum xlfNot=38;
	enum xlfMod=39;
	enum xlfDcount=40;
	enum xlfDsum=41;
	enum xlfDaverage=42;
	enum xlfDmin=43;
	enum xlfDmax=44;
	enum xlfDstdev=45;
	enum xlfVar=46;
	enum xlfDvar=47;
	enum xlfText=48;
	enum xlfLinest=49;
	enum xlfTrend=50;
	enum xlfLogest=51;
	enum xlfGrowth=52;
	enum xlfGoto=53;
	enum xlfHalt=54;
	enum xlfPv=56;
	enum xlfFv=57;
	enum xlfNper=58;
	enum xlfPmt=59;
	enum xlfRate=60;
	enum xlfMirr=61;
	enum xlfIrr=62;
	enum xlfRand=63;
	enum xlfMatch=64;
	enum xlfDate=65;
	enum xlfTime=66;
	enum xlfDay=67;
	enum xlfMonth=68;
	enum xlfYear=69;
	enum xlfWeekday=70;
	enum xlfHour=71;
	enum xlfMinute=72;
	enum xlfSecond=73;
	enum xlfNow=74;
	enum xlfAreas=75;
	enum xlfRows=76;
	enum xlfColumns=77;
	enum xlfOffset=78;
	enum xlfAbsref=79;
	enum xlfRelref=80;
	enum xlfArgument=81;
	enum xlfSearch=82;
	enum xlfTranspose=83;
	enum xlfError=84;
	enum xlfStep=85;
	enum xlfType=86;
	enum xlfEcho=87;
	enum xlfSetName=88;
	enum xlfCaller=89;
	enum xlfDeref=90;
	enum xlfWindows=91;
	enum xlfSeries=92;
	enum xlfDocuments=93;
	enum xlfActiveCell=94;
	enum xlfSelection=95;
	enum xlfResult=96;
	enum xlfAtan2=97;
	enum xlfAsin=98;
	enum xlfAcos=99;
	enum xlfChoose=100;
	enum xlfHlookup=101;
	enum xlfVlookup=102;
	enum xlfLinks=103;
	enum xlfInput=104;
	enum xlfIsref=105;
	enum xlfGetFormula=106;
	enum xlfGetName=107;
	enum xlfSetValue=108;
	enum xlfLog=109;
	enum xlfExec=110;
	enum xlfChar=111;
	enum xlfLower=112;
	enum xlfUpper=113;
	enum xlfProper=114;
	enum xlfLeft=115;
	enum xlfRight=116;
	enum xlfExact=117;
	enum xlfTrim=118;
	enum xlfReplace=119;
	enum xlfSubstitute=120;
	enum xlfCode=121;
	enum xlfNames=122;
	enum xlfDirectory=123;
	enum xlfFind=124;
	enum xlfCell=125;
	enum xlfIserr=126;
	enum xlfIstext=127;
	enum xlfIsnumber=128;
	enum xlfIsblank=129;
	enum xlfT=130;
	enum xlfN=131;
	enum xlfFopen=132;
	enum xlfFclose=133;
	enum xlfFsize=134;
	enum xlfFreadln=135;
	enum xlfFread=136;
	enum xlfFwriteln=137;
	enum xlfFwrite=138;
	enum xlfFpos=139;
	enum xlfDatevalue=140;
	enum xlfTimevalue=141;
	enum xlfSln=142;
	enum xlfSyd=143;
	enum xlfDdb=144;
	enum xlfGetDef=145;
	enum xlfReftext=146;
	enum xlfTextref=147;
	enum xlfIndirect=148;
	enum xlfRegister=149;
	enum xlfCall=150;
	enum xlfAddBar=151;
	enum xlfAddMenu=152;
	enum xlfAddCommand=153;
	enum xlfEnableCommand=154;
	enum xlfCheckCommand=155;
	enum xlfRenameCommand=156;
	enum xlfShowBar=157;
	enum xlfDeleteMenu=158;
	enum xlfDeleteCommand=159;
	enum xlfGetChartItem=160;
	enum xlfDialogBox=161;
	enum xlfClean=162;
	enum xlfMdeterm=163;
	enum xlfMinverse=164;
	enum xlfMmult=165;
	enum xlfFiles=166;
	enum xlfIpmt=167;
	enum xlfPpmt=168;
	enum xlfCounta=169;
	enum xlfCancelKey=170;
	enum xlfInitiate=175;
	enum xlfRequest=176;
	enum xlfPoke=177;
	enum xlfExecute=178;
	enum xlfTerminate=179;
	enum xlfRestart=180;
	enum xlfHelp=181;
	enum xlfGetBar=182;
	enum xlfProduct=183;
	enum xlfFact=184;
	enum xlfGetCell=185;
	enum xlfGetWorkspace=186;
	enum xlfGetWindow=187;
	enum xlfGetDocument=188;
	enum xlfDproduct=189;
	enum xlfIsnontext=190;
	enum xlfGetNote=191;
	enum xlfNote=192;
	enum xlfStdevp=193;
	enum xlfVarp=194;
	enum xlfDstdevp=195;
	enum xlfDvarp=196;
	enum xlfTrunc=197;
	enum xlfIslogical=198;
	enum xlfDcounta=199;
	enum xlfDeleteBar=200;
	enum xlfUnregister=201;
	enum xlfUsdollar=204;
	enum xlfFindb=205;
	enum xlfSearchb=206;
	enum xlfReplaceb=207;
	enum xlfLeftb=208;
	enum xlfRightb=209;
	enum xlfMidb=210;
	enum xlfLenb=211;
	enum xlfRoundup=212;
	enum xlfRounddown=213;
	enum xlfAsc=214;
	enum xlfDbcs=215;
	enum xlfRank=216;
	enum xlfAddress=219;
	enum xlfDays360=220;
	enum xlfToday=221;
	enum xlfVdb=222;
	enum xlfMedian=227;
	enum xlfSumproduct=228;
	enum xlfSinh=229;
	enum xlfCosh=230;
	enum xlfTanh=231;
	enum xlfAsinh=232;
	enum xlfAcosh=233;
	enum xlfAtanh=234;
	enum xlfDget=235;
	enum xlfCreateObject=236;
	enum xlfVolatile=237;
	enum xlfLastError=238;
	enum xlfCustomUndo=239;
	enum xlfCustomRepeat=240;
	enum xlfFormulaConvert=241;
	enum xlfGetLinkInfo=242;
	enum xlfTextBox=243;
	enum xlfInfo=244;
	enum xlfGroup=245;
	enum xlfGetObject=246;
	enum xlfDb=247;
	enum xlfPause=248;
	enum xlfResume=251;
	enum xlfFrequency=252;
	enum xlfAddToolbar=253;
	enum xlfDeleteToolbar=254;
	enum xlfResetToolbar=256;
	enum xlfEvaluate=257;
	enum xlfGetToolbar=258;
	enum xlfGetTool=259;
	enum xlfSpellingCheck=260;
	enum xlfErrorType=261;
	enum xlfAppTitle=262;
	enum xlfWindowTitle=263;
	enum xlfSaveToolbar=264;
	enum xlfEnableTool=265;
	enum xlfPressTool=266;
	enum xlfRegisterId=267;
	enum xlfGetWorkbook=268;
	enum xlfAvedev=269;
	enum xlfBetadist=270;
	enum xlfGammaln=271;
	enum xlfBetainv=272;
	enum xlfBinomdist=273;
	enum xlfChidist=274;
	enum xlfChiinv=275;
	enum xlfCombin=276;
	enum xlfConfidence=277;
	enum xlfCritbinom=278;
	enum xlfEven=279;
	enum xlfExpondist=280;
	enum xlfFdist=281;
	enum xlfFinv=282;
	enum xlfFisher=283;
	enum xlfFisherinv=284;
	enum xlfFloor=285;
	enum xlfGammadist=286;
	enum xlfGammainv=287;
	enum xlfCeiling=288;
	enum xlfHypgeomdist=289;
	enum xlfLognormdist=290;
	enum xlfLoginv=291;
	enum xlfNegbinomdist=292;
	enum xlfNormdist=293;
	enum xlfNormsdist=294;
	enum xlfNorminv=295;
	enum xlfNormsinv=296;
	enum xlfStandardize=297;
	enum xlfOdd=298;
	enum xlfPermut=299;
	enum xlfPoisson=300;
	enum xlfTdist=301;
	enum xlfWeibull=302;
	enum xlfSumxmy2=303;
	enum xlfSumx2my2=304;
	enum xlfSumx2py2=305;
	enum xlfChitest=306;
	enum xlfCorrel=307;
	enum xlfCovar=308;
	enum xlfForecast=309;
	enum xlfFtest=310;
	enum xlfIntercept=311;
	enum xlfPearson=312;
	enum xlfRsq=313;
	enum xlfSteyx=314;
	enum xlfSlope=315;
	enum xlfTtest=316;
	enum xlfProb=317;
	enum xlfDevsq=318;
	enum xlfGeomean=319;
	enum xlfHarmean=320;
	enum xlfSumsq=321;
	enum xlfKurt=322;
	enum xlfSkew=323;
	enum xlfZtest=324;
	enum xlfLarge=325;
	enum xlfSmall=326;
	enum xlfQuartile=327;
	enum xlfPercentile=328;
	enum xlfPercentrank=329;
	enum xlfMode=330;
	enum xlfTrimmean=331;
	enum xlfTinv=332;
	enum xlfMovieCommand=334;
	enum xlfGetMovie=335;
	enum xlfConcatenate=336;
	enum xlfPower=337;
	enum xlfPivotAddData=338;
	enum xlfGetPivotTable=339;
	enum xlfGetPivotField=340;
	enum xlfGetPivotItem=341;
	enum xlfRadians=342;
	enum xlfDegrees=343;
	enum xlfSubtotal=344;
	enum xlfSumif=345;
	enum xlfCountif=346;
	enum xlfCountblank=347;
	enum xlfScenarioGet=348;
	enum xlfOptionsListsGet=349;
	enum xlfIspmt=350;
	enum xlfDatedif=351;
	enum xlfDatestring=352;
	enum xlfNumberstring=353;
	enum xlfRoman=354;
	enum xlfOpenDialog=355;
	enum xlfSaveDialog=356;
	enum xlfViewGet=357;
	enum xlfGetpivotdata=358;
	enum xlfHyperlink=359;
	enum xlfPhonetic=360;
	enum xlfAveragea=361;
	enum xlfMaxa=362;
	enum xlfMina=363;
	enum xlfStdevpa=364;
	enum xlfVarpa=365;
	enum xlfStdeva=366;
	enum xlfVara=367;
	enum xlfBahttext=368;
	enum xlfThaidayofweek=369;
	enum xlfThaidigit=370;
	enum xlfThaimonthofyear=371;
	enum xlfThainumsound=372;
	enum xlfThainumstring=373;
	enum xlfThaistringlength=374;
	enum xlfIsthaidigit=375;
	enum xlfRoundbahtdown=376;
	enum xlfRoundbahtup=377;
	enum xlfThaiyear=378;
	enum xlfRtd=379;
	enum xlfCubevalue=380;
	enum xlfCubemember=381;
	enum xlfCubememberproperty=382;
	enum xlfCuberankedmember=383;
	enum xlfHex2bin=384;
	enum xlfHex2dec=385;
	enum xlfHex2oct=386;
	enum xlfDec2bin=387;
	enum xlfDec2hex=388;
	enum xlfDec2oct=389;
	enum xlfOct2bin=390;
	enum xlfOct2hex=391;
	enum xlfOct2dec=392;
	enum xlfBin2dec=393;
	enum xlfBin2oct=394;
	enum xlfBin2hex=395;
	enum xlfImsub=396;
	enum xlfImdiv=397;
	enum xlfImpower=398;
	enum xlfImabs=399;
	enum xlfImsqrt=400;
	enum xlfImln=401;
	enum xlfImlog2=402;
	enum xlfImlog10=403;
	enum xlfImsin=404;
	enum xlfImcos=405;
	enum xlfImexp=406;
	enum xlfImargument=407;
	enum xlfImconjugate=408;
	enum xlfImaginary=409;
	enum xlfImreal=410;
	enum xlfComplex=411;
	enum xlfImsum=412;
	enum xlfImproduct=413;
	enum xlfSeriessum=414;
	enum xlfFactdouble=415;
	enum xlfSqrtpi=416;
	enum xlfQuotient=417;
	enum xlfDelta=418;
	enum xlfGestep=419;
	enum xlfIseven=420;
	enum xlfIsodd=421;
	enum xlfMround=422;
	enum xlfErf=423;
	enum xlfErfc=424;
	enum xlfBesselj=425;
	enum xlfBesselk=426;
	enum xlfBessely=427;
	enum xlfBesseli=428;
	enum xlfXirr=429;
	enum xlfXnpv=430;
	enum xlfPricemat=431;
	enum xlfYieldmat=432;
	enum xlfIntrate=433;
	enum xlfReceived=434;
	enum xlfDisc=435;
	enum xlfPricedisc=436;
	enum xlfYielddisc=437;
	enum xlfTbilleq=438;
	enum xlfTbillprice=439;
	enum xlfTbillyield=440;
	enum xlfPrice=441;
	enum xlfYield=442;
	enum xlfDollarde=443;
	enum xlfDollarfr=444;
	enum xlfNominal=445;
	enum xlfEffect=446;
	enum xlfCumprinc=447;
	enum xlfCumipmt=448;
	enum xlfEdate=449;
	enum xlfEomonth=450;
	enum xlfYearfrac=451;
	enum xlfCoupdaybs=452;
	enum xlfCoupdays=453;
	enum xlfCoupdaysnc=454;
	enum xlfCoupncd=455;
	enum xlfCoupnum=456;
	enum xlfCouppcd=457;
	enum xlfDuration=458;
	enum xlfMduration=459;
	enum xlfOddlprice=460;
	enum xlfOddlyield=461;
	enum xlfOddfprice=462;
	enum xlfOddfyield=463;
	enum xlfRandbetween=464;
	enum xlfWeeknum=465;
	enum xlfAmordegrc=466;
	enum xlfAmorlinc=467;
	enum xlfConvert=468;
	enum xlfAccrint=469;
	enum xlfAccrintm=470;
	enum xlfWorkday=471;
	enum xlfNetworkdays=472;
	enum xlfGcd=473;
	enum xlfMultinomial=474;
	enum xlfLcm=475;
	enum xlfFvschedule=476;
	enum xlfCubekpimember=477;
	enum xlfCubeset=478;
	enum xlfCubesetcount=479;
	enum xlfIferror=480;
	enum xlfCountifs=481;
	enum xlfSumifs=482;
	enum xlfAverageif=483;
	enum xlfAverageifs=484;
	enum xlfAggregate=485;
	enum xlfBinom_dist=486;
	enum xlfBinom_inv=487;
	enum xlfConfidence_norm=488;
	enum xlfConfidence_t=489;
	enum xlfChisq_test=490;
	enum xlfF_test=491;
	enum xlfCovariance_p=492;
	enum xlfCovariance_s=493;
	enum xlfExpon_dist=494;
	enum xlfGamma_dist=495;
	enum xlfGamma_inv=496;
	enum xlfMode_mult=497;
	enum xlfMode_sngl=498;
	enum xlfNorm_dist=499;
	enum xlfNorm_inv=500;
	enum xlfPercentile_exc=501;
	enum xlfPercentile_inc=502;
	enum xlfPercentrank_exc=503;
	enum xlfPercentrank_inc=504;
	enum xlfPoisson_dist=505;
	enum xlfQuartile_exc=506;
	enum xlfQuartile_inc=507;
	enum xlfRank_avg=508;
	enum xlfRank_eq=509;
	enum xlfStdev_s=510;
	enum xlfStdev_p=511;
	enum xlfT_dist=512;
	enum xlfT_dist_2t=513;
	enum xlfT_dist_rt=514;
	enum xlfT_inv=515;
	enum xlfT_inv_2t=516;
	enum xlfVar_s=517;
	enum xlfVar_p=518;
	enum xlfWeibull_dist=519;
	enum xlfNetworkdays_intl=520;
	enum xlfWorkday_intl=521;
	enum xlfEcma_ceiling=522;
	enum xlfIso_ceiling=523;
	enum xlfBeta_dist=525;
	enum xlfBeta_inv=526;
	enum xlfChisq_dist=527;
	enum xlfChisq_dist_rt=528;
	enum xlfChisq_inv=529;
	enum xlfChisq_inv_rt=530;
	enum xlfF_dist=531;
	enum xlfF_dist_rt=532;
	enum xlfF_inv=533;
	enum xlfF_inv_rt=534;
	enum xlfHypgeom_dist=535;
	enum xlfLognorm_dist=536;
	enum xlfLognorm_inv=537;
	enum xlfNegbinom_dist=538;
	enum xlfNorm_s_dist=539;
	enum xlfNorm_s_inv=540;
	enum xlfT_test=541;
	enum xlfZ_test=542;
	enum xlfErf_precise=543;
	enum xlfErfc_precise=544;
	enum xlfGammaln_precise=545;
	enum xlfCeiling_precise=546;
	enum xlfFloor_precise=547;
	enum xlfAcot=548;
	enum xlfAcoth=549;
	enum xlfCot=550;
	enum xlfCoth=551;
	enum xlfCsc=552;
	enum xlfCsch=553;
	enum xlfSec=554;
	enum xlfSech=555;
	enum xlfImtan=556;
	enum xlfImcot=557;
	enum xlfImcsc=558;
	enum xlfImcsch=559;
	enum xlfImsec=560;
	enum xlfImsech=561;
	enum xlfBitand=562;
	enum xlfBitor=563;
	enum xlfBitxor=564;
	enum xlfBitlshift=565;
	enum xlfBitrshift=566;
	enum xlfPermutationa=567;
	enum xlfCombina=568;
	enum xlfXor=569;
	enum xlfPduration=570;
	enum xlfBase=571;
	enum xlfDecimal=572;
	enum xlfDays=573;
	enum xlfBinom_dist_range=574;
	enum xlfGamma=575;
	enum xlfSkew_p=576;
	enum xlfGauss=577;
	enum xlfPhi=578;
	enum xlfRri=579;
	enum xlfUnichar=580;
	enum xlfUnicode=581;
	enum xlfMunit=582;
	enum xlfArabic=583;
	enum xlfIsoweeknum=584;
	enum xlfNumbervalue=585;
	enum xlfSheet=586;
	enum xlfSheets=587;
	enum xlfFormulatext=588;
	enum xlfIsformula=589;
	enum xlfIfna=590;
	enum xlfCeiling_math=591;
	enum xlfFloor_math=592;
	enum xlfImsinh=593;
	enum xlfImcosh=594;
	enum xlfFilterxml=595;
	enum xlfWebservice=596;
	enum xlfEncodeurl=597;

	/* Excel command numbers */
	enum xlcBeep=(0 | xlCommand);
	enum xlcOpen=(1 | xlCommand);
	enum xlcOpenLinks=(2 | xlCommand);
	enum xlcCloseAll=(3 | xlCommand);
	enum xlcSave=(4 | xlCommand);
	enum xlcSaveAs=(5 | xlCommand);
	enum xlcFileDelete=(6 | xlCommand);
	enum xlcPageSetup=(7 | xlCommand);
	enum xlcPrint=(8 | xlCommand);
	enum xlcPrinterSetup=(9 | xlCommand);
	enum xlcQuit=(10 | xlCommand);
	enum xlcNewWindow=(11 | xlCommand);
	enum xlcArrangeAll=(12 | xlCommand);
	enum xlcWindowSize=(13 | xlCommand);
	enum xlcWindowMove=(14 | xlCommand);
	enum xlcFull=(15 | xlCommand);
	enum xlcClose=(16 | xlCommand);
	enum xlcRun=(17 | xlCommand);
	enum xlcSetPrintArea=(22 | xlCommand);
	enum xlcSetPrintTitles=(23 | xlCommand);
	enum xlcSetPageBreak=(24 | xlCommand);
	enum xlcRemovePageBreak=(25 | xlCommand);
	enum xlcFont=(26 | xlCommand);
	enum xlcDisplay=(27 | xlCommand);
	enum xlcProtectDocument=(28 | xlCommand);
	enum xlcPrecision=(29 | xlCommand);
	enum xlcA1R1c1=(30 | xlCommand);
	enum xlcCalculateNow=(31 | xlCommand);
	enum xlcCalculation=(32 | xlCommand);
	enum xlcDataFind=(34 | xlCommand);
	enum xlcExtract=(35 | xlCommand);
	enum xlcDataDelete=(36 | xlCommand);
	enum xlcSetDatabase=(37 | xlCommand);
	enum xlcSetCriteria=(38 | xlCommand);
	enum xlcSort=(39 | xlCommand);
	enum xlcDataSeries=(40 | xlCommand);
	enum xlcTable=(41 | xlCommand);
	enum xlcFormatNumber=(42 | xlCommand);
	enum xlcAlignment=(43 | xlCommand);
	enum xlcStyle=(44 | xlCommand);
	enum xlcBorder=(45 | xlCommand);
	enum xlcCellProtection=(46 | xlCommand);
	enum xlcColumnWidth=(47 | xlCommand);
	enum xlcUndo=(48 | xlCommand);
	enum xlcCut=(49 | xlCommand);
	enum xlcCopy=(50 | xlCommand);
	enum xlcPaste=(51 | xlCommand);
	enum xlcClear=(52 | xlCommand);
	enum xlcPasteSpecial=(53 | xlCommand);
	enum xlcEditDelete=(54 | xlCommand);
	enum xlcInsert=(55 | xlCommand);
	enum xlcFillRight=(56 | xlCommand);
	enum xlcFillDown=(57 | xlCommand);
	enum xlcDefineName=(61 | xlCommand);
	enum xlcCreateNames=(62 | xlCommand);
	enum xlcFormulaGoto=(63 | xlCommand);
	enum xlcFormulaFind=(64 | xlCommand);
	enum xlcSelectLastCell=(65 | xlCommand);
	enum xlcShowActiveCell=(66 | xlCommand);
	enum xlcGalleryArea=(67 | xlCommand);
	enum xlcGalleryBar=(68 | xlCommand);
	enum xlcGalleryColumn=(69 | xlCommand);
	enum xlcGalleryLine=(70 | xlCommand);
	enum xlcGalleryPie=(71 | xlCommand);
	enum xlcGalleryScatter=(72 | xlCommand);
	enum xlcCombination=(73 | xlCommand);
	enum xlcPreferred=(74 | xlCommand);
	enum xlcAddOverlay=(75 | xlCommand);
	enum xlcGridlines=(76 | xlCommand);
	enum xlcSetPreferred=(77 | xlCommand);
	enum xlcAxes=(78 | xlCommand);
	enum xlcLegend=(79 | xlCommand);
	enum xlcAttachText=(80 | xlCommand);
	enum xlcAddArrow=(81 | xlCommand);
	enum xlcSelectChart=(82 | xlCommand);
	enum xlcSelectPlotArea=(83 | xlCommand);
	enum xlcPatterns=(84 | xlCommand);
	enum xlcMainChart=(85 | xlCommand);
	enum xlcOverlay=(86 | xlCommand);
	enum xlcScale=(87 | xlCommand);
	enum xlcFormatLegend=(88 | xlCommand);
	enum xlcFormatText=(89 | xlCommand);
	enum xlcEditRepeat=(90 | xlCommand);
	enum xlcParse=(91 | xlCommand);
	enum xlcJustify=(92 | xlCommand);
	enum xlcHide=(93 | xlCommand);
	enum xlcUnhide=(94 | xlCommand);
	enum xlcWorkspace=(95 | xlCommand);
	enum xlcFormula=(96 | xlCommand);
	enum xlcFormulaFill=(97 | xlCommand);
	enum xlcFormulaArray=(98 | xlCommand);
	enum xlcDataFindNext=(99 | xlCommand);
	enum xlcDataFindPrev=(100 | xlCommand);
	enum xlcFormulaFindNext=(101 | xlCommand);
	enum xlcFormulaFindPrev=(102 | xlCommand);
	enum xlcActivate=(103 | xlCommand);
	enum xlcActivateNext=(104 | xlCommand);
	enum xlcActivatePrev=(105 | xlCommand);
	enum xlcUnlockedNext=(106 | xlCommand);
	enum xlcUnlockedPrev=(107 | xlCommand);
	enum xlcCopyPicture=(108 | xlCommand);
	enum xlcSelect=(109 | xlCommand);
	enum xlcDeleteName=(110 | xlCommand);
	enum xlcDeleteFormat=(111 | xlCommand);
	enum xlcVline=(112 | xlCommand);
	enum xlcHline=(113 | xlCommand);
	enum xlcVpage=(114 | xlCommand);
	enum xlcHpage=(115 | xlCommand);
	enum xlcVscroll=(116 | xlCommand);
	enum xlcHscroll=(117 | xlCommand);
	enum xlcAlert=(118 | xlCommand);
	enum xlcNew=(119 | xlCommand);
	enum xlcCancelCopy=(120 | xlCommand);
	enum xlcShowClipboard=(121 | xlCommand);
	enum xlcMessage=(122 | xlCommand);
	enum xlcPasteLink=(124 | xlCommand);
	enum xlcAppActivate=(125 | xlCommand);
	enum xlcDeleteArrow=(126 | xlCommand);
	enum xlcRowHeight=(127 | xlCommand);
	enum xlcFormatMove=(128 | xlCommand);
	enum xlcFormatSize=(129 | xlCommand);
	enum xlcFormulaReplace=(130 | xlCommand);
	enum xlcSendKeys=(131 | xlCommand);
	enum xlcSelectSpecial=(132 | xlCommand);
	enum xlcApplyNames=(133 | xlCommand);
	enum xlcReplaceFont=(134 | xlCommand);
	enum xlcFreezePanes=(135 | xlCommand);
	enum xlcShowInfo=(136 | xlCommand);
	enum xlcSplit=(137 | xlCommand);
	enum xlcOnWindow=(138 | xlCommand);
	enum xlcOnData=(139 | xlCommand);
	enum xlcDisableInput=(140 | xlCommand);
	enum xlcEcho=(141 | xlCommand);
	enum xlcOutline=(142 | xlCommand);
	enum xlcListNames=(143 | xlCommand);
	enum xlcFileClose=(144 | xlCommand);
	enum xlcSaveWorkbook=(145 | xlCommand);
	enum xlcDataForm=(146 | xlCommand);
	enum xlcCopyChart=(147 | xlCommand);
	enum xlcOnTime=(148 | xlCommand);
	enum xlcWait=(149 | xlCommand);
	enum xlcFormatFont=(150 | xlCommand);
	enum xlcFillUp=(151 | xlCommand);
	enum xlcFillLeft=(152 | xlCommand);
	enum xlcDeleteOverlay=(153 | xlCommand);
	enum xlcNote=(154 | xlCommand);
	enum xlcShortMenus=(155 | xlCommand);
	enum xlcSetUpdateStatus=(159 | xlCommand);
	enum xlcColorPalette=(161 | xlCommand);
	enum xlcDeleteStyle=(162 | xlCommand);
	enum xlcWindowRestore=(163 | xlCommand);
	enum xlcWindowMaximize=(164 | xlCommand);
	enum xlcError=(165 | xlCommand);
	enum xlcChangeLink=(166 | xlCommand);
	enum xlcCalculateDocument=(167 | xlCommand);
	enum xlcOnKey=(168 | xlCommand);
	enum xlcAppRestore=(169 | xlCommand);
	enum xlcAppMove=(170 | xlCommand);
	enum xlcAppSize=(171 | xlCommand);
	enum xlcAppMinimize=(172 | xlCommand);
	enum xlcAppMaximize=(173 | xlCommand);
	enum xlcBringToFront=(174 | xlCommand);
	enum xlcSendToBack=(175 | xlCommand);
	enum xlcMainChartType=(185 | xlCommand);
	enum xlcOverlayChartType=(186 | xlCommand);
	enum xlcSelectEnd=(187 | xlCommand);
	enum xlcOpenMail=(188 | xlCommand);
	enum xlcSendMail=(189 | xlCommand);
	enum xlcStandardFont=(190 | xlCommand);
	enum xlcConsolidate=(191 | xlCommand);
	enum xlcSortSpecial=(192 | xlCommand);
	enum xlcGallery3dArea=(193 | xlCommand);
	enum xlcGallery3dColumn=(194 | xlCommand);
	enum xlcGallery3dLine=(195 | xlCommand);
	enum xlcGallery3dPie=(196 | xlCommand);
	enum xlcView3d=(197 | xlCommand);
	enum xlcGoalSeek=(198 | xlCommand);
	enum xlcWorkgroup=(199 | xlCommand);
	enum xlcFillGroup=(200 | xlCommand);
	enum xlcUpdateLink=(201 | xlCommand);
	enum xlcPromote=(202 | xlCommand);
	enum xlcDemote=(203 | xlCommand);
	enum xlcShowDetail=(204 | xlCommand);
	enum xlcUngroup=(206 | xlCommand);
	enum xlcObjectProperties=(207 | xlCommand);
	enum xlcSaveNewObject=(208 | xlCommand);
	enum xlcShare=(209 | xlCommand);
	enum xlcShareName=(210 | xlCommand);
	enum xlcDuplicate=(211 | xlCommand);
	enum xlcApplyStyle=(212 | xlCommand);
	enum xlcAssignToObject=(213 | xlCommand);
	enum xlcObjectProtection=(214 | xlCommand);
	enum xlcHideObject=(215 | xlCommand);
	enum xlcSetExtract=(216 | xlCommand);
	enum xlcCreatePublisher=(217 | xlCommand);
	enum xlcSubscribeTo=(218 | xlCommand);
	enum xlcAttributes=(219 | xlCommand);
	enum xlcShowToolbar=(220 | xlCommand);
	enum xlcPrintPreview=(222 | xlCommand);
	enum xlcEditColor=(223 | xlCommand);
	enum xlcShowLevels=(224 | xlCommand);
	enum xlcFormatMain=(225 | xlCommand);
	enum xlcFormatOverlay=(226 | xlCommand);
	enum xlcOnRecalc=(227 | xlCommand);
	enum xlcEditSeries=(228 | xlCommand);
	enum xlcDefineStyle=(229 | xlCommand);
	enum xlcLinePrint=(240 | xlCommand);
	enum xlcEnterData=(243 | xlCommand);
	enum xlcGalleryRadar=(249 | xlCommand);
	enum xlcMergeStyles=(250 | xlCommand);
	enum xlcEditionOptions=(251 | xlCommand);
	enum xlcPastePicture=(252 | xlCommand);
	enum xlcPastePictureLink=(253 | xlCommand);
	enum xlcSpelling=(254 | xlCommand);
	enum xlcZoom=(256 | xlCommand);
	enum xlcResume=(258 | xlCommand);
	enum xlcInsertObject=(259 | xlCommand);
	enum xlcWindowMinimize=(260 | xlCommand);
	enum xlcSize=(261 | xlCommand);
	enum xlcMove=(262 | xlCommand);
	enum xlcSoundNote=(265 | xlCommand);
	enum xlcSoundPlay=(266 | xlCommand);
	enum xlcFormatShape=(267 | xlCommand);
	enum xlcExtendPolygon=(268 | xlCommand);
	enum xlcFormatAuto=(269 | xlCommand);
	enum xlcGallery3dBar=(272 | xlCommand);
	enum xlcGallery3dSurface=(273 | xlCommand);
	enum xlcFillAuto=(274 | xlCommand);
	enum xlcCustomizeToolbar=(276 | xlCommand);
	enum xlcAddTool=(277 | xlCommand);
	enum xlcEditObject=(278 | xlCommand);
	enum xlcOnDoubleclick=(279 | xlCommand);
	enum xlcOnEntry=(280 | xlCommand);
	enum xlcWorkbookAdd=(281 | xlCommand);
	enum xlcWorkbookMove=(282 | xlCommand);
	enum xlcWorkbookCopy=(283 | xlCommand);
	enum xlcWorkbookOptions=(284 | xlCommand);
	enum xlcSaveWorkspace=(285 | xlCommand);
	enum xlcChartWizard=(288 | xlCommand);
	enum xlcDeleteTool=(289 | xlCommand);
	enum xlcMoveTool=(290 | xlCommand);
	enum xlcWorkbookSelect=(291 | xlCommand);
	enum xlcWorkbookActivate=(292 | xlCommand);
	enum xlcAssignToTool=(293 | xlCommand);
	enum xlcCopyTool=(295 | xlCommand);
	enum xlcResetTool=(296 | xlCommand);
	enum xlcConstrainNumeric=(297 | xlCommand);
	enum xlcPasteTool=(298 | xlCommand);
	enum xlcPlacement=(300 | xlCommand);
	enum xlcFillWorkgroup=(301 | xlCommand);
	enum xlcWorkbookNew=(302 | xlCommand);
	enum xlcScenarioCells=(305 | xlCommand);
	enum xlcScenarioDelete=(306 | xlCommand);
	enum xlcScenarioAdd=(307 | xlCommand);
	enum xlcScenarioEdit=(308 | xlCommand);
	enum xlcScenarioShow=(309 | xlCommand);
	enum xlcScenarioShowNext=(310 | xlCommand);
	enum xlcScenarioSummary=(311 | xlCommand);
	enum xlcPivotTableWizard=(312 | xlCommand);
	enum xlcPivotFieldProperties=(313 | xlCommand);
	enum xlcPivotField=(314 | xlCommand);
	enum xlcPivotItem=(315 | xlCommand);
	enum xlcPivotAddFields=(316 | xlCommand);
	enum xlcOptionsCalculation=(318 | xlCommand);
	enum xlcOptionsEdit=(319 | xlCommand);
	enum xlcOptionsView=(320 | xlCommand);
	enum xlcAddinManager=(321 | xlCommand);
	enum xlcMenuEditor=(322 | xlCommand);
	enum xlcAttachToolbars=(323 | xlCommand);
	enum xlcVbaactivate=(324 | xlCommand);
	enum xlcOptionsChart=(325 | xlCommand);
	enum xlcVbaInsertFile=(328 | xlCommand);
	enum xlcVbaProcedureDefinition=(330 | xlCommand);
	enum xlcRoutingSlip=(336 | xlCommand);
	enum xlcRouteDocument=(338 | xlCommand);
	enum xlcMailLogon=(339 | xlCommand);
	enum xlcInsertPicture=(342 | xlCommand);
	enum xlcEditTool=(343 | xlCommand);
	enum xlcGalleryDoughnut=(344 | xlCommand);
	enum xlcChartTrend=(350 | xlCommand);
	enum xlcPivotItemProperties=(352 | xlCommand);
	enum xlcWorkbookInsert=(354 | xlCommand);
	enum xlcOptionsTransition=(355 | xlCommand);
	enum xlcOptionsGeneral=(356 | xlCommand);
	enum xlcFilterAdvanced=(370 | xlCommand);
	enum xlcMailAddMailer=(373 | xlCommand);
	enum xlcMailDeleteMailer=(374 | xlCommand);
	enum xlcMailReply=(375 | xlCommand);
	enum xlcMailReplyAll=(376 | xlCommand);
	enum xlcMailForward=(377 | xlCommand);
	enum xlcMailNextLetter=(378 | xlCommand);
	enum xlcDataLabel=(379 | xlCommand);
	enum xlcInsertTitle=(380 | xlCommand);
	enum xlcFontProperties=(381 | xlCommand);
	enum xlcMacroOptions=(382 | xlCommand);
	enum xlcWorkbookHide=(383 | xlCommand);
	enum xlcWorkbookUnhide=(384 | xlCommand);
	enum xlcWorkbookDelete=(385 | xlCommand);
	enum xlcWorkbookName=(386 | xlCommand);
	enum xlcGalleryCustom=(388 | xlCommand);
	enum xlcAddChartAutoformat=(390 | xlCommand);
	enum xlcDeleteChartAutoformat=(391 | xlCommand);
	enum xlcChartAddData=(392 | xlCommand);
	enum xlcAutoOutline=(393 | xlCommand);
	enum xlcTabOrder=(394 | xlCommand);
	enum xlcShowDialog=(395 | xlCommand);
	enum xlcSelectAll=(396 | xlCommand);
	enum xlcUngroupSheets=(397 | xlCommand);
	enum xlcSubtotalCreate=(398 | xlCommand);
	enum xlcSubtotalRemove=(399 | xlCommand);
	enum xlcRenameObject=(400 | xlCommand);
	enum xlcWorkbookScroll=(412 | xlCommand);
	enum xlcWorkbookNext=(413 | xlCommand);
	enum xlcWorkbookPrev=(414 | xlCommand);
	enum xlcWorkbookTabSplit=(415 | xlCommand);
	enum xlcFullScreen=(416 | xlCommand);
	enum xlcWorkbookProtect=(417 | xlCommand);
	enum xlcScrollbarProperties=(420 | xlCommand);
	enum xlcPivotShowPages=(421 | xlCommand);
	enum xlcTextToColumns=(422 | xlCommand);
	enum xlcFormatCharttype=(423 | xlCommand);
	enum xlcLinkFormat=(424 | xlCommand);
	enum xlcTracerDisplay=(425 | xlCommand);
	enum xlcTracerNavigate=(430 | xlCommand);
	enum xlcTracerClear=(431 | xlCommand);
	enum xlcTracerError=(432 | xlCommand);
	enum xlcPivotFieldGroup=(433 | xlCommand);
	enum xlcPivotFieldUngroup=(434 | xlCommand);
	enum xlcCheckboxProperties=(435 | xlCommand);
	enum xlcLabelProperties=(436 | xlCommand);
	enum xlcListboxProperties=(437 | xlCommand);
	enum xlcEditboxProperties=(438 | xlCommand);
	enum xlcPivotRefresh=(439 | xlCommand);
	enum xlcLinkCombo=(440 | xlCommand);
	enum xlcOpenText=(441 | xlCommand);
	enum xlcHideDialog=(442 | xlCommand);
	enum xlcSetDialogFocus=(443 | xlCommand);
	enum xlcEnableObject=(444 | xlCommand);
	enum xlcPushbuttonProperties=(445 | xlCommand);
	enum xlcSetDialogDefault=(446 | xlCommand);
	enum xlcFilter=(447 | xlCommand);
	enum xlcFilterShowAll=(448 | xlCommand);
	enum xlcClearOutline=(449 | xlCommand);
	enum xlcFunctionWizard=(450 | xlCommand);
	enum xlcAddListItem=(451 | xlCommand);
	enum xlcSetListItem=(452 | xlCommand);
	enum xlcRemoveListItem=(453 | xlCommand);
	enum xlcSelectListItem=(454 | xlCommand);
	enum xlcSetControlValue=(455 | xlCommand);
	enum xlcSaveCopyAs=(456 | xlCommand);
	enum xlcOptionsListsAdd=(458 | xlCommand);
	enum xlcOptionsListsDelete=(459 | xlCommand);
	enum xlcSeriesAxes=(460 | xlCommand);
	enum xlcSeriesX=(461 | xlCommand);
	enum xlcSeriesY=(462 | xlCommand);
	enum xlcErrorbarX=(463 | xlCommand);
	enum xlcErrorbarY=(464 | xlCommand);
	enum xlcFormatChart=(465 | xlCommand);
	enum xlcSeriesOrder=(466 | xlCommand);
	enum xlcMailLogoff=(467 | xlCommand);
	enum xlcClearRoutingSlip=(468 | xlCommand);
	enum xlcAppActivateMicrosoft=(469 | xlCommand);
	enum xlcMailEditMailer=(470 | xlCommand);
	enum xlcOnSheet=(471 | xlCommand);
	enum xlcStandardWidth=(472 | xlCommand);
	enum xlcScenarioMerge=(473 | xlCommand);
	enum xlcSummaryInfo=(474 | xlCommand);
	enum xlcFindFile=(475 | xlCommand);
	enum xlcActiveCellFont=(476 | xlCommand);
	enum xlcEnableTipwizard=(477 | xlCommand);
	enum xlcVbaMakeAddin=(478 | xlCommand);
	enum xlcInsertdatatable=(480 | xlCommand);
	enum xlcWorkgroupOptions=(481 | xlCommand);
	enum xlcMailSendMailer=(482 | xlCommand);
	enum xlcAutocorrect=(485 | xlCommand);
	enum xlcPostDocument=(489 | xlCommand);
	enum xlcPicklist=(491 | xlCommand);
	enum xlcViewShow=(493 | xlCommand);
	enum xlcViewDefine=(494 | xlCommand);
	enum xlcViewDelete=(495 | xlCommand);
	enum xlcSheetBackground=(509 | xlCommand);
	enum xlcInsertMapObject=(510 | xlCommand);
	enum xlcOptionsMenono=(511 | xlCommand);
	enum xlcNormal=(518 | xlCommand);
	enum xlcLayout=(519 | xlCommand);
	enum xlcRmPrintArea=(520 | xlCommand);
	enum xlcClearPrintArea=(521 | xlCommand);
	enum xlcAddPrintArea=(522 | xlCommand);
	enum xlcMoveBrk=(523 | xlCommand);
	enum xlcHidecurrNote=(545 | xlCommand);
	enum xlcHideallNotes=(546 | xlCommand);
	enum xlcDeleteNote=(547 | xlCommand);
	enum xlcTraverseNotes=(548 | xlCommand);
	enum xlcActivateNotes=(549 | xlCommand);
	enum xlcProtectRevisions=(620 | xlCommand);
	enum xlcUnprotectRevisions=(621 | xlCommand);
	enum xlcOptionsMe=(647 | xlCommand);
	enum xlcWebPublish=(653 | xlCommand);
	enum xlcNewwebquery=(667 | xlCommand);
	enum xlcPivotTableChart=(673 | xlCommand);
	enum xlcOptionsSave=(753 | xlCommand);
	enum xlcOptionsSpell=(755 | xlCommand);
	enum xlcHideallInkannots=(808 | xlCommand);

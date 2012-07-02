/*
  Copyright (C) 2005-2012 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration
{
	public class XlCall
	{
		/*
		** Return codes
		**
		** These values can be returned from Excel4() or Excel4v().
		*/
		public enum XlReturn
		{
			XlReturnSuccess   = 0,        /* success */ 
			XlReturnAbort     = 1,        /* macro halted */
			XlReturnInvXlfn   = 2,        /* invalid function number */ 
			XlReturnInvCount  = 4,        /* invalid number of arguments */ 
			XlReturnInvXloper = 8,        /* invalid OPER structure */  
			XlReturnStackOvfl = 16,       /* stack overflow */  
			XlReturnFailed    = 32,       /* command failed */
			XlReturnUncalced  = 64,       /* uncalced cell */
            XlReturnNotThreadSafe = 128,   /* not allowed during multi-threaded calc */
            XlReturnInvAsynchronousContext  = 256,  /* invalid asynchronous function handle */
            XlReturnNotClusterSafe = 512  /* not supported on cluster */
		}

        #region Constants
        /*
		** Function number bits
		*/
		public static readonly int xlCommand = 0x8000;
		public static readonly int xlSpecial = 0x4000;
		public static readonly int xlIntl = 0x2000;
		public static readonly int xlPrompt = 0x1000;

        /*
        ** XLL events
        **
        ** Passed in to an xlEventRegister call to register a corresponding event.
        */

        public static readonly int xleventCalculationEnded = 1;    /* Fires at the end of calculation */
        public static readonly int xleventCalculationCanceled = 2;    /* Fires when calculation is interrupted */

		/*
		** Auxiliary function numbers
		**
		** These functions are available only from the C API,
		** not from the Excel macro language.
		*/
		public static readonly int xlFree =             (0  | xlSpecial);
		public static readonly int xlStack =            (1  | xlSpecial);
		public static readonly int xlCoerce =           (2  | xlSpecial);
		public static readonly int xlSet =              (3  | xlSpecial);
		public static readonly int xlSheetId =          (4  | xlSpecial);
		public static readonly int xlSheetNm =          (5  | xlSpecial);
		public static readonly int xlAbort =            (6  | xlSpecial);
        public static readonly int xlGetInst =          (7  | xlSpecial);  /* Returns application's hinstance as an integer value, supported on 32-bit platform only */
		public static readonly int xlGetHwnd =          (8  | xlSpecial);
		public static readonly int xlGetName =          (9  | xlSpecial);
		public static readonly int xlEnableXLMsgs =     (10 | xlSpecial);
		public static readonly int xlDisableXLMsgs =    (11 | xlSpecial);
		public static readonly int xlDefineBinaryName = (12 | xlSpecial);
		public static readonly int xlGetBinaryName =    (13 | xlSpecial);
        public static readonly int xlAsyncReturn =      (16 | xlSpecial);	/* Set return value from an asynchronous function call */
        public static readonly int xlEventRegister =    (17 | xlSpecial);	/* Register an XLL event */
        public static readonly int xlRunningOnCluster = (18 | xlSpecial);	/* Returns true if running on Compute Cluster */
        public static readonly int xlGetInstPtr =       (19 | xlSpecial);	/* Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms */

        /* edit modes */
		public static readonly int xlModeReady = 0;	// not in edit mode
		public static readonly int xlModeEnter = 1;	// enter mode
		public static readonly int xlModeEdit = 2;	// edit mode
		public static readonly int xlModePoint = 4;	// point mode

        /* document(page) types */
		public static readonly int dtNil = 0x7f;	// window is not a sheet, macro, chart or basic
                                                    // OR window is not the selected window at idle state
		public static readonly int dtSheet = 0;	// sheet
		public static readonly int dtProc = 1;	// XLM macro
		public static readonly int dtChart = 2;	// Chart
		public static readonly int dtBasic = 6;	// VBA 

		/* 
		** User defined function
		**
		** First argument should be a function reference.
		*/
		public static readonly int xlUDF =      255;

		/*
		** Built-in Functions and Command Equivalents
		*/

        // Excel function numbers

        public static readonly int xlfCount = 0;
        public static readonly int xlfIsna = 2;
        public static readonly int xlfIserror = 3;
        public static readonly int xlfSum = 4;
        public static readonly int xlfAverage = 5;
        public static readonly int xlfMin = 6;
        public static readonly int xlfMax = 7;
        public static readonly int xlfRow = 8;
        public static readonly int xlfColumn = 9;
        public static readonly int xlfNa = 10;
        public static readonly int xlfNpv = 11;
        public static readonly int xlfStdev = 12;
        public static readonly int xlfDollar = 13;
        public static readonly int xlfFixed = 14;
        public static readonly int xlfSin = 15;
        public static readonly int xlfCos = 16;
        public static readonly int xlfTan = 17;
        public static readonly int xlfAtan = 18;
        public static readonly int xlfPi = 19;
        public static readonly int xlfSqrt = 20;
        public static readonly int xlfExp = 21;
        public static readonly int xlfLn = 22;
        public static readonly int xlfLog10 = 23;
        public static readonly int xlfAbs = 24;
        public static readonly int xlfInt = 25;
        public static readonly int xlfSign = 26;
        public static readonly int xlfRound = 27;
        public static readonly int xlfLookup = 28;
        public static readonly int xlfIndex = 29;
        public static readonly int xlfRept = 30;
        public static readonly int xlfMid = 31;
        public static readonly int xlfLen = 32;
        public static readonly int xlfValue = 33;
        public static readonly int xlfTrue = 34;
        public static readonly int xlfFalse = 35;
        public static readonly int xlfAnd = 36;
        public static readonly int xlfOr = 37;
        public static readonly int xlfNot = 38;
        public static readonly int xlfMod = 39;
        public static readonly int xlfDcount = 40;
        public static readonly int xlfDsum = 41;
        public static readonly int xlfDaverage = 42;
        public static readonly int xlfDmin = 43;
        public static readonly int xlfDmax = 44;
        public static readonly int xlfDstdev = 45;
        public static readonly int xlfVar = 46;
        public static readonly int xlfDvar = 47;
        public static readonly int xlfText = 48;
        public static readonly int xlfLinest = 49;
        public static readonly int xlfTrend = 50;
        public static readonly int xlfLogest = 51;
        public static readonly int xlfGrowth = 52;
        public static readonly int xlfGoto = 53;
        public static readonly int xlfHalt = 54;
        public static readonly int xlfPv = 56;
        public static readonly int xlfFv = 57;
        public static readonly int xlfNper = 58;
        public static readonly int xlfPmt = 59;
        public static readonly int xlfRate = 60;
        public static readonly int xlfMirr = 61;
        public static readonly int xlfIrr = 62;
        public static readonly int xlfRand = 63;
        public static readonly int xlfMatch = 64;
        public static readonly int xlfDate = 65;
        public static readonly int xlfTime = 66;
        public static readonly int xlfDay = 67;
        public static readonly int xlfMonth = 68;
        public static readonly int xlfYear = 69;
        public static readonly int xlfWeekday = 70;
        public static readonly int xlfHour = 71;
        public static readonly int xlfMinute = 72;
        public static readonly int xlfSecond = 73;
        public static readonly int xlfNow = 74;
        public static readonly int xlfAreas = 75;
        public static readonly int xlfRows = 76;
        public static readonly int xlfColumns = 77;
        public static readonly int xlfOffset = 78;
        public static readonly int xlfAbsref = 79;
        public static readonly int xlfRelref = 80;
        public static readonly int xlfArgument = 81;
        public static readonly int xlfSearch = 82;
        public static readonly int xlfTranspose = 83;
        public static readonly int xlfError = 84;
        public static readonly int xlfStep = 85;
        public static readonly int xlfType = 86;
        public static readonly int xlfEcho = 87;
        public static readonly int xlfSetName = 88;
        public static readonly int xlfCaller = 89;
        public static readonly int xlfDeref = 90;
        public static readonly int xlfWindows = 91;
        public static readonly int xlfSeries = 92;
        public static readonly int xlfDocuments = 93;
        public static readonly int xlfActiveCell = 94;
        public static readonly int xlfSelection = 95;
        public static readonly int xlfResult = 96;
        public static readonly int xlfAtan2 = 97;
        public static readonly int xlfAsin = 98;
        public static readonly int xlfAcos = 99;
        public static readonly int xlfChoose = 100;
        public static readonly int xlfHlookup = 101;
        public static readonly int xlfVlookup = 102;
        public static readonly int xlfLinks = 103;
        public static readonly int xlfInput = 104;
        public static readonly int xlfIsref = 105;
        public static readonly int xlfGetFormula = 106;
        public static readonly int xlfGetName = 107;
        public static readonly int xlfSetValue = 108;
        public static readonly int xlfLog = 109;
        public static readonly int xlfExec = 110;
        public static readonly int xlfChar = 111;
        public static readonly int xlfLower = 112;
        public static readonly int xlfUpper = 113;
        public static readonly int xlfProper = 114;
        public static readonly int xlfLeft = 115;
        public static readonly int xlfRight = 116;
        public static readonly int xlfExact = 117;
        public static readonly int xlfTrim = 118;
        public static readonly int xlfReplace = 119;
        public static readonly int xlfSubstitute = 120;
        public static readonly int xlfCode = 121;
        public static readonly int xlfNames = 122;
        public static readonly int xlfDirectory = 123;
        public static readonly int xlfFind = 124;
        public static readonly int xlfCell = 125;
        public static readonly int xlfIserr = 126;
        public static readonly int xlfIstext = 127;
        public static readonly int xlfIsnumber = 128;
        public static readonly int xlfIsblank = 129;
        public static readonly int xlfT = 130;
        public static readonly int xlfN = 131;
        public static readonly int xlfFopen = 132;
        public static readonly int xlfFclose = 133;
        public static readonly int xlfFsize = 134;
        public static readonly int xlfFreadln = 135;
        public static readonly int xlfFread = 136;
        public static readonly int xlfFwriteln = 137;
        public static readonly int xlfFwrite = 138;
        public static readonly int xlfFpos = 139;
        public static readonly int xlfDatevalue = 140;
        public static readonly int xlfTimevalue = 141;
        public static readonly int xlfSln = 142;
        public static readonly int xlfSyd = 143;
        public static readonly int xlfDdb = 144;
        public static readonly int xlfGetDef = 145;
        public static readonly int xlfReftext = 146;
        public static readonly int xlfTextref = 147;
        public static readonly int xlfIndirect = 148;
        public static readonly int xlfRegister = 149;
        public static readonly int xlfCall = 150;
        public static readonly int xlfAddBar = 151;
        public static readonly int xlfAddMenu = 152;
        public static readonly int xlfAddCommand = 153;
        public static readonly int xlfEnableCommand = 154;
        public static readonly int xlfCheckCommand = 155;
        public static readonly int xlfRenameCommand = 156;
        public static readonly int xlfShowBar = 157;
        public static readonly int xlfDeleteMenu = 158;
        public static readonly int xlfDeleteCommand = 159;
        public static readonly int xlfGetChartItem = 160;
        public static readonly int xlfDialogBox = 161;
        public static readonly int xlfClean = 162;
        public static readonly int xlfMdeterm = 163;
        public static readonly int xlfMinverse = 164;
        public static readonly int xlfMmult = 165;
        public static readonly int xlfFiles = 166;
        public static readonly int xlfIpmt = 167;
        public static readonly int xlfPpmt = 168;
        public static readonly int xlfCounta = 169;
        public static readonly int xlfCancelKey = 170;
        public static readonly int xlfInitiate = 175;
        public static readonly int xlfRequest = 176;
        public static readonly int xlfPoke = 177;
        public static readonly int xlfExecute = 178;
        public static readonly int xlfTerminate = 179;
        public static readonly int xlfRestart = 180;
        public static readonly int xlfHelp = 181;
        public static readonly int xlfGetBar = 182;
        public static readonly int xlfProduct = 183;
        public static readonly int xlfFact = 184;
        public static readonly int xlfGetCell = 185;
        public static readonly int xlfGetWorkspace = 186;
        public static readonly int xlfGetWindow = 187;
        public static readonly int xlfGetDocument = 188;
        public static readonly int xlfDproduct = 189;
        public static readonly int xlfIsnontext = 190;
        public static readonly int xlfGetNote = 191;
        public static readonly int xlfNote = 192;
        public static readonly int xlfStdevp = 193;
        public static readonly int xlfVarp = 194;
        public static readonly int xlfDstdevp = 195;
        public static readonly int xlfDvarp = 196;
        public static readonly int xlfTrunc = 197;
        public static readonly int xlfIslogical = 198;
        public static readonly int xlfDcounta = 199;
        public static readonly int xlfDeleteBar = 200;
        public static readonly int xlfUnregister = 201;
        public static readonly int xlfUsdollar = 204;
        public static readonly int xlfFindb = 205;
        public static readonly int xlfSearchb = 206;
        public static readonly int xlfReplaceb = 207;
        public static readonly int xlfLeftb = 208;
        public static readonly int xlfRightb = 209;
        public static readonly int xlfMidb = 210;
        public static readonly int xlfLenb = 211;
        public static readonly int xlfRoundup = 212;
        public static readonly int xlfRounddown = 213;
        public static readonly int xlfAsc = 214;
        public static readonly int xlfDbcs = 215;
        public static readonly int xlfRank = 216;
        public static readonly int xlfAddress = 219;
        public static readonly int xlfDays360 = 220;
        public static readonly int xlfToday = 221;
        public static readonly int xlfVdb = 222;
        public static readonly int xlfMedian = 227;
        public static readonly int xlfSumproduct = 228;
        public static readonly int xlfSinh = 229;
        public static readonly int xlfCosh = 230;
        public static readonly int xlfTanh = 231;
        public static readonly int xlfAsinh = 232;
        public static readonly int xlfAcosh = 233;
        public static readonly int xlfAtanh = 234;
        public static readonly int xlfDget = 235;
        public static readonly int xlfCreateObject = 236;
        public static readonly int xlfVolatile = 237;
        public static readonly int xlfLastError = 238;
        public static readonly int xlfCustomUndo = 239;
        public static readonly int xlfCustomRepeat = 240;
        public static readonly int xlfFormulaConvert = 241;
        public static readonly int xlfGetLinkInfo = 242;
        public static readonly int xlfTextBox = 243;
        public static readonly int xlfInfo = 244;
        public static readonly int xlfGroup = 245;
        public static readonly int xlfGetObject = 246;
        public static readonly int xlfDb = 247;
        public static readonly int xlfPause = 248;
        public static readonly int xlfResume = 251;
        public static readonly int xlfFrequency = 252;
        public static readonly int xlfAddToolbar = 253;
        public static readonly int xlfDeleteToolbar = 254;
        public static readonly int xlfResetToolbar = 256;
        public static readonly int xlfEvaluate = 257;
        public static readonly int xlfGetToolbar = 258;
        public static readonly int xlfGetTool = 259;
        public static readonly int xlfSpellingCheck = 260;
        public static readonly int xlfErrorType = 261;
        public static readonly int xlfAppTitle = 262;
        public static readonly int xlfWindowTitle = 263;
        public static readonly int xlfSaveToolbar = 264;
        public static readonly int xlfEnableTool = 265;
        public static readonly int xlfPressTool = 266;
        public static readonly int xlfRegisterId = 267;
        public static readonly int xlfGetWorkbook = 268;
        public static readonly int xlfAvedev = 269;
        public static readonly int xlfBetadist = 270;
        public static readonly int xlfGammaln = 271;
        public static readonly int xlfBetainv = 272;
        public static readonly int xlfBinomdist = 273;
        public static readonly int xlfChidist = 274;
        public static readonly int xlfChiinv = 275;
        public static readonly int xlfCombin = 276;
        public static readonly int xlfConfidence = 277;
        public static readonly int xlfCritbinom = 278;
        public static readonly int xlfEven = 279;
        public static readonly int xlfExpondist = 280;
        public static readonly int xlfFdist = 281;
        public static readonly int xlfFinv = 282;
        public static readonly int xlfFisher = 283;
        public static readonly int xlfFisherinv = 284;
        public static readonly int xlfFloor = 285;
        public static readonly int xlfGammadist = 286;
        public static readonly int xlfGammainv = 287;
        public static readonly int xlfCeiling = 288;
        public static readonly int xlfHypgeomdist = 289;
        public static readonly int xlfLognormdist = 290;
        public static readonly int xlfLoginv = 291;
        public static readonly int xlfNegbinomdist = 292;
        public static readonly int xlfNormdist = 293;
        public static readonly int xlfNormsdist = 294;
        public static readonly int xlfNorminv = 295;
        public static readonly int xlfNormsinv = 296;
        public static readonly int xlfStandardize = 297;
        public static readonly int xlfOdd = 298;
        public static readonly int xlfPermut = 299;
        public static readonly int xlfPoisson = 300;
        public static readonly int xlfTdist = 301;
        public static readonly int xlfWeibull = 302;
        public static readonly int xlfSumxmy2 = 303;
        public static readonly int xlfSumx2my2 = 304;
        public static readonly int xlfSumx2py2 = 305;
        public static readonly int xlfChitest = 306;
        public static readonly int xlfCorrel = 307;
        public static readonly int xlfCovar = 308;
        public static readonly int xlfForecast = 309;
        public static readonly int xlfFtest = 310;
        public static readonly int xlfIntercept = 311;
        public static readonly int xlfPearson = 312;
        public static readonly int xlfRsq = 313;
        public static readonly int xlfSteyx = 314;
        public static readonly int xlfSlope = 315;
        public static readonly int xlfTtest = 316;
        public static readonly int xlfProb = 317;
        public static readonly int xlfDevsq = 318;
        public static readonly int xlfGeomean = 319;
        public static readonly int xlfHarmean = 320;
        public static readonly int xlfSumsq = 321;
        public static readonly int xlfKurt = 322;
        public static readonly int xlfSkew = 323;
        public static readonly int xlfZtest = 324;
        public static readonly int xlfLarge = 325;
        public static readonly int xlfSmall = 326;
        public static readonly int xlfQuartile = 327;
        public static readonly int xlfPercentile = 328;
        public static readonly int xlfPercentrank = 329;
        public static readonly int xlfMode = 330;
        public static readonly int xlfTrimmean = 331;
        public static readonly int xlfTinv = 332;
        public static readonly int xlfMovieCommand = 334;
        public static readonly int xlfGetMovie = 335;
        public static readonly int xlfConcatenate = 336;
        public static readonly int xlfPower = 337;
        public static readonly int xlfPivotAddData = 338;
        public static readonly int xlfGetPivotTable = 339;
        public static readonly int xlfGetPivotField = 340;
        public static readonly int xlfGetPivotItem = 341;
        public static readonly int xlfRadians = 342;
        public static readonly int xlfDegrees = 343;
        public static readonly int xlfSubtotal = 344;
        public static readonly int xlfSumif = 345;
        public static readonly int xlfCountif = 346;
        public static readonly int xlfCountblank = 347;
        public static readonly int xlfScenarioGet = 348;
        public static readonly int xlfOptionsListsGet = 349;
        public static readonly int xlfIspmt = 350;
        public static readonly int xlfDatedif = 351;
        public static readonly int xlfDatestring = 352;
        public static readonly int xlfNumberstring = 353;
        public static readonly int xlfRoman = 354;
        public static readonly int xlfOpenDialog = 355;
        public static readonly int xlfSaveDialog = 356;
        public static readonly int xlfViewGet = 357;
        public static readonly int xlfGetpivotdata = 358;
        public static readonly int xlfHyperlink = 359;
        public static readonly int xlfPhonetic = 360;
        public static readonly int xlfAveragea = 361;
        public static readonly int xlfMaxa = 362;
        public static readonly int xlfMina = 363;
        public static readonly int xlfStdevpa = 364;
        public static readonly int xlfVarpa = 365;
        public static readonly int xlfStdeva = 366;
        public static readonly int xlfVara = 367;
        public static readonly int xlfBahttext = 368;
        public static readonly int xlfThaidayofweek = 369;
        public static readonly int xlfThaidigit = 370;
        public static readonly int xlfThaimonthofyear = 371;
        public static readonly int xlfThainumsound = 372;
        public static readonly int xlfThainumstring = 373;
        public static readonly int xlfThaistringlength = 374;
        public static readonly int xlfIsthaidigit = 375;
        public static readonly int xlfRoundbahtdown = 376;
        public static readonly int xlfRoundbahtup = 377;
        public static readonly int xlfThaiyear = 378;
        public static readonly int xlfRtd = 379;
        public static readonly int xlfCubevalue = 380;
        public static readonly int xlfCubemember = 381;
        public static readonly int xlfCubememberproperty = 382;
        public static readonly int xlfCuberankedmember = 383;
        public static readonly int xlfHex2bin = 384;
        public static readonly int xlfHex2dec = 385;
        public static readonly int xlfHex2oct = 386;
        public static readonly int xlfDec2bin = 387;
        public static readonly int xlfDec2hex = 388;
        public static readonly int xlfDec2oct = 389;
        public static readonly int xlfOct2bin = 390;
        public static readonly int xlfOct2hex = 391;
        public static readonly int xlfOct2dec = 392;
        public static readonly int xlfBin2dec = 393;
        public static readonly int xlfBin2oct = 394;
        public static readonly int xlfBin2hex = 395;
        public static readonly int xlfImsub = 396;
        public static readonly int xlfImdiv = 397;
        public static readonly int xlfImpower = 398;
        public static readonly int xlfImabs = 399;
        public static readonly int xlfImsqrt = 400;
        public static readonly int xlfImln = 401;
        public static readonly int xlfImlog2 = 402;
        public static readonly int xlfImlog10 = 403;
        public static readonly int xlfImsin = 404;
        public static readonly int xlfImcos = 405;
        public static readonly int xlfImexp = 406;
        public static readonly int xlfImargument = 407;
        public static readonly int xlfImconjugate = 408;
        public static readonly int xlfImaginary = 409;
        public static readonly int xlfImreal = 410;
        public static readonly int xlfComplex = 411;
        public static readonly int xlfImsum = 412;
        public static readonly int xlfImproduct = 413;
        public static readonly int xlfSeriessum = 414;
        public static readonly int xlfFactdouble = 415;
        public static readonly int xlfSqrtpi = 416;
        public static readonly int xlfQuotient = 417;
        public static readonly int xlfDelta = 418;
        public static readonly int xlfGestep = 419;
        public static readonly int xlfIseven = 420;
        public static readonly int xlfIsodd = 421;
        public static readonly int xlfMround = 422;
        public static readonly int xlfErf = 423;
        public static readonly int xlfErfc = 424;
        public static readonly int xlfBesselj = 425;
        public static readonly int xlfBesselk = 426;
        public static readonly int xlfBessely = 427;
        public static readonly int xlfBesseli = 428;
        public static readonly int xlfXirr = 429;
        public static readonly int xlfXnpv = 430;
        public static readonly int xlfPricemat = 431;
        public static readonly int xlfYieldmat = 432;
        public static readonly int xlfIntrate = 433;
        public static readonly int xlfReceived = 434;
        public static readonly int xlfDisc = 435;
        public static readonly int xlfPricedisc = 436;
        public static readonly int xlfYielddisc = 437;
        public static readonly int xlfTbilleq = 438;
        public static readonly int xlfTbillprice = 439;
        public static readonly int xlfTbillyield = 440;
        public static readonly int xlfPrice = 441;
        public static readonly int xlfYield = 442;
        public static readonly int xlfDollarde = 443;
        public static readonly int xlfDollarfr = 444;
        public static readonly int xlfNominal = 445;
        public static readonly int xlfEffect = 446;
        public static readonly int xlfCumprinc = 447;
        public static readonly int xlfCumipmt = 448;
        public static readonly int xlfEdate = 449;
        public static readonly int xlfEomonth = 450;
        public static readonly int xlfYearfrac = 451;
        public static readonly int xlfCoupdaybs = 452;
        public static readonly int xlfCoupdays = 453;
        public static readonly int xlfCoupdaysnc = 454;
        public static readonly int xlfCoupncd = 455;
        public static readonly int xlfCoupnum = 456;
        public static readonly int xlfCouppcd = 457;
        public static readonly int xlfDuration = 458;
        public static readonly int xlfMduration = 459;
        public static readonly int xlfOddlprice = 460;
        public static readonly int xlfOddlyield = 461;
        public static readonly int xlfOddfprice = 462;
        public static readonly int xlfOddfyield = 463;
        public static readonly int xlfRandbetween = 464;
        public static readonly int xlfWeeknum = 465;
        public static readonly int xlfAmordegrc = 466;
        public static readonly int xlfAmorlinc = 467;
        public static readonly int xlfConvert = 468;
        public static readonly int xlfAccrint = 469;
        public static readonly int xlfAccrintm = 470;
        public static readonly int xlfWorkday = 471;
        public static readonly int xlfNetworkdays = 472;
        public static readonly int xlfGcd = 473;
        public static readonly int xlfMultinomial = 474;
        public static readonly int xlfLcm = 475;
        public static readonly int xlfFvschedule = 476;
        public static readonly int xlfCubekpimember = 477;
        public static readonly int xlfCubeset = 478;
        public static readonly int xlfCubesetcount = 479;
        public static readonly int xlfIferror = 480;
        public static readonly int xlfCountifs = 481;
        public static readonly int xlfSumifs = 482;
        public static readonly int xlfAverageif = 483;
        public static readonly int xlfAverageifs = 484;
        public static readonly int xlfAggregate = 485;
        public static readonly int xlfBinom_dist = 486;
        public static readonly int xlfBinom_inv = 487;
        public static readonly int xlfConfidence_norm = 488;
        public static readonly int xlfConfidence_t = 489;
        public static readonly int xlfChisq_test = 490;
        public static readonly int xlfF_test = 491;
        public static readonly int xlfCovariance_p = 492;
        public static readonly int xlfCovariance_s = 493;
        public static readonly int xlfExpon_dist = 494;
        public static readonly int xlfGamma_dist = 495;
        public static readonly int xlfGamma_inv = 496;
        public static readonly int xlfMode_mult = 497;
        public static readonly int xlfMode_sngl = 498;
        public static readonly int xlfNorm_dist = 499;
        public static readonly int xlfNorm_inv = 500;
        public static readonly int xlfPercentile_exc = 501;
        public static readonly int xlfPercentile_inc = 502;
        public static readonly int xlfPercentrank_exc = 503;
        public static readonly int xlfPercentrank_inc = 504;
        public static readonly int xlfPoisson_dist = 505;
        public static readonly int xlfQuartile_exc = 506;
        public static readonly int xlfQuartile_inc = 507;
        public static readonly int xlfRank_avg = 508;
        public static readonly int xlfRank_eq = 509;
        public static readonly int xlfStdev_s = 510;
        public static readonly int xlfStdev_p = 511;
        public static readonly int xlfT_dist = 512;
        public static readonly int xlfT_dist_2t = 513;
        public static readonly int xlfT_dist_rt = 514;
        public static readonly int xlfT_inv = 515;
        public static readonly int xlfT_inv_2t = 516;
        public static readonly int xlfVar_s = 517;
        public static readonly int xlfVar_p = 518;
        public static readonly int xlfWeibull_dist = 519;
        public static readonly int xlfNetworkdays_intl = 520;
        public static readonly int xlfWorkday_intl = 521;
        public static readonly int xlfEcma_ceiling = 522;
        public static readonly int xlfIso_ceiling = 523;
        public static readonly int xlfBeta_dist = 525;
        public static readonly int xlfBeta_inv = 526;
        public static readonly int xlfChisq_dist = 527;
        public static readonly int xlfChisq_dist_rt = 528;
        public static readonly int xlfChisq_inv = 529;
        public static readonly int xlfChisq_inv_rt = 530;
        public static readonly int xlfF_dist = 531;
        public static readonly int xlfF_dist_rt = 532;
        public static readonly int xlfF_inv = 533;
        public static readonly int xlfF_inv_rt = 534;
        public static readonly int xlfHypgeom_dist = 535;
        public static readonly int xlfLognorm_dist = 536;
        public static readonly int xlfLognorm_inv = 537;
        public static readonly int xlfNegbinom_dist = 538;
        public static readonly int xlfNorm_s_dist = 539;
        public static readonly int xlfNorm_s_inv = 540;
        public static readonly int xlfT_test = 541;
        public static readonly int xlfZ_test = 542;
        public static readonly int xlfErf_precise = 543;
        public static readonly int xlfErfc_precise = 544;
        public static readonly int xlfGammaln_precise = 545;
        public static readonly int xlfCeiling_precise = 546;
        public static readonly int xlfFloor_precise = 547;

        /* Excel command numbers */
        public static readonly int xlcBeep = (0 | xlCommand);
        public static readonly int xlcOpen = (1 | xlCommand);
        public static readonly int xlcOpenLinks = (2 | xlCommand);
        public static readonly int xlcCloseAll = (3 | xlCommand);
        public static readonly int xlcSave = (4 | xlCommand);
        public static readonly int xlcSaveAs = (5 | xlCommand);
        public static readonly int xlcFileDelete = (6 | xlCommand);
        public static readonly int xlcPageSetup = (7 | xlCommand);
        public static readonly int xlcPrint = (8 | xlCommand);
        public static readonly int xlcPrinterSetup = (9 | xlCommand);
        public static readonly int xlcQuit = (10 | xlCommand);
        public static readonly int xlcNewWindow = (11 | xlCommand);
        public static readonly int xlcArrangeAll = (12 | xlCommand);
        public static readonly int xlcWindowSize = (13 | xlCommand);
        public static readonly int xlcWindowMove = (14 | xlCommand);
        public static readonly int xlcFull = (15 | xlCommand);
        public static readonly int xlcClose = (16 | xlCommand);
        public static readonly int xlcRun = (17 | xlCommand);
        public static readonly int xlcSetPrintArea = (22 | xlCommand);
        public static readonly int xlcSetPrintTitles = (23 | xlCommand);
        public static readonly int xlcSetPageBreak = (24 | xlCommand);
        public static readonly int xlcRemovePageBreak = (25 | xlCommand);
        public static readonly int xlcFont = (26 | xlCommand);
        public static readonly int xlcDisplay = (27 | xlCommand);
        public static readonly int xlcProtectDocument = (28 | xlCommand);
        public static readonly int xlcPrecision = (29 | xlCommand);
        public static readonly int xlcA1R1c1 = (30 | xlCommand);
        public static readonly int xlcCalculateNow = (31 | xlCommand);
        public static readonly int xlcCalculation = (32 | xlCommand);
        public static readonly int xlcDataFind = (34 | xlCommand);
        public static readonly int xlcExtract = (35 | xlCommand);
        public static readonly int xlcDataDelete = (36 | xlCommand);
        public static readonly int xlcSetDatabase = (37 | xlCommand);
        public static readonly int xlcSetCriteria = (38 | xlCommand);
        public static readonly int xlcSort = (39 | xlCommand);
        public static readonly int xlcDataSeries = (40 | xlCommand);
        public static readonly int xlcTable = (41 | xlCommand);
        public static readonly int xlcFormatNumber = (42 | xlCommand);
        public static readonly int xlcAlignment = (43 | xlCommand);
        public static readonly int xlcStyle = (44 | xlCommand);
        public static readonly int xlcBorder = (45 | xlCommand);
        public static readonly int xlcCellProtection = (46 | xlCommand);
        public static readonly int xlcColumnWidth = (47 | xlCommand);
        public static readonly int xlcUndo = (48 | xlCommand);
        public static readonly int xlcCut = (49 | xlCommand);
        public static readonly int xlcCopy = (50 | xlCommand);
        public static readonly int xlcPaste = (51 | xlCommand);
        public static readonly int xlcClear = (52 | xlCommand);
        public static readonly int xlcPasteSpecial = (53 | xlCommand);
        public static readonly int xlcEditDelete = (54 | xlCommand);
        public static readonly int xlcInsert = (55 | xlCommand);
        public static readonly int xlcFillRight = (56 | xlCommand);
        public static readonly int xlcFillDown = (57 | xlCommand);
        public static readonly int xlcDefineName = (61 | xlCommand);
        public static readonly int xlcCreateNames = (62 | xlCommand);
        public static readonly int xlcFormulaGoto = (63 | xlCommand);
        public static readonly int xlcFormulaFind = (64 | xlCommand);
        public static readonly int xlcSelectLastCell = (65 | xlCommand);
        public static readonly int xlcShowActiveCell = (66 | xlCommand);
        public static readonly int xlcGalleryArea = (67 | xlCommand);
        public static readonly int xlcGalleryBar = (68 | xlCommand);
        public static readonly int xlcGalleryColumn = (69 | xlCommand);
        public static readonly int xlcGalleryLine = (70 | xlCommand);
        public static readonly int xlcGalleryPie = (71 | xlCommand);
        public static readonly int xlcGalleryScatter = (72 | xlCommand);
        public static readonly int xlcCombination = (73 | xlCommand);
        public static readonly int xlcPreferred = (74 | xlCommand);
        public static readonly int xlcAddOverlay = (75 | xlCommand);
        public static readonly int xlcGridlines = (76 | xlCommand);
        public static readonly int xlcSetPreferred = (77 | xlCommand);
        public static readonly int xlcAxes = (78 | xlCommand);
        public static readonly int xlcLegend = (79 | xlCommand);
        public static readonly int xlcAttachText = (80 | xlCommand);
        public static readonly int xlcAddArrow = (81 | xlCommand);
        public static readonly int xlcSelectChart = (82 | xlCommand);
        public static readonly int xlcSelectPlotArea = (83 | xlCommand);
        public static readonly int xlcPatterns = (84 | xlCommand);
        public static readonly int xlcMainChart = (85 | xlCommand);
        public static readonly int xlcOverlay = (86 | xlCommand);
        public static readonly int xlcScale = (87 | xlCommand);
        public static readonly int xlcFormatLegend = (88 | xlCommand);
        public static readonly int xlcFormatText = (89 | xlCommand);
        public static readonly int xlcEditRepeat = (90 | xlCommand);
        public static readonly int xlcParse = (91 | xlCommand);
        public static readonly int xlcJustify = (92 | xlCommand);
        public static readonly int xlcHide = (93 | xlCommand);
        public static readonly int xlcUnhide = (94 | xlCommand);
        public static readonly int xlcWorkspace = (95 | xlCommand);
        public static readonly int xlcFormula = (96 | xlCommand);
        public static readonly int xlcFormulaFill = (97 | xlCommand);
        public static readonly int xlcFormulaArray = (98 | xlCommand);
        public static readonly int xlcDataFindNext = (99 | xlCommand);
        public static readonly int xlcDataFindPrev = (100 | xlCommand);
        public static readonly int xlcFormulaFindNext = (101 | xlCommand);
        public static readonly int xlcFormulaFindPrev = (102 | xlCommand);
        public static readonly int xlcActivate = (103 | xlCommand);
        public static readonly int xlcActivateNext = (104 | xlCommand);
        public static readonly int xlcActivatePrev = (105 | xlCommand);
        public static readonly int xlcUnlockedNext = (106 | xlCommand);
        public static readonly int xlcUnlockedPrev = (107 | xlCommand);
        public static readonly int xlcCopyPicture = (108 | xlCommand);
        public static readonly int xlcSelect = (109 | xlCommand);
        public static readonly int xlcDeleteName = (110 | xlCommand);
        public static readonly int xlcDeleteFormat = (111 | xlCommand);
        public static readonly int xlcVline = (112 | xlCommand);
        public static readonly int xlcHline = (113 | xlCommand);
        public static readonly int xlcVpage = (114 | xlCommand);
        public static readonly int xlcHpage = (115 | xlCommand);
        public static readonly int xlcVscroll = (116 | xlCommand);
        public static readonly int xlcHscroll = (117 | xlCommand);
        public static readonly int xlcAlert = (118 | xlCommand);
        public static readonly int xlcNew = (119 | xlCommand);
        public static readonly int xlcCancelCopy = (120 | xlCommand);
        public static readonly int xlcShowClipboard = (121 | xlCommand);
        public static readonly int xlcMessage = (122 | xlCommand);
        public static readonly int xlcPasteLink = (124 | xlCommand);
        public static readonly int xlcAppActivate = (125 | xlCommand);
        public static readonly int xlcDeleteArrow = (126 | xlCommand);
        public static readonly int xlcRowHeight = (127 | xlCommand);
        public static readonly int xlcFormatMove = (128 | xlCommand);
        public static readonly int xlcFormatSize = (129 | xlCommand);
        public static readonly int xlcFormulaReplace = (130 | xlCommand);
        public static readonly int xlcSendKeys = (131 | xlCommand);
        public static readonly int xlcSelectSpecial = (132 | xlCommand);
        public static readonly int xlcApplyNames = (133 | xlCommand);
        public static readonly int xlcReplaceFont = (134 | xlCommand);
        public static readonly int xlcFreezePanes = (135 | xlCommand);
        public static readonly int xlcShowInfo = (136 | xlCommand);
        public static readonly int xlcSplit = (137 | xlCommand);
        public static readonly int xlcOnWindow = (138 | xlCommand);
        public static readonly int xlcOnData = (139 | xlCommand);
        public static readonly int xlcDisableInput = (140 | xlCommand);
        public static readonly int xlcEcho = (141 | xlCommand);
        public static readonly int xlcOutline = (142 | xlCommand);
        public static readonly int xlcListNames = (143 | xlCommand);
        public static readonly int xlcFileClose = (144 | xlCommand);
        public static readonly int xlcSaveWorkbook = (145 | xlCommand);
        public static readonly int xlcDataForm = (146 | xlCommand);
        public static readonly int xlcCopyChart = (147 | xlCommand);
        public static readonly int xlcOnTime = (148 | xlCommand);
        public static readonly int xlcWait = (149 | xlCommand);
        public static readonly int xlcFormatFont = (150 | xlCommand);
        public static readonly int xlcFillUp = (151 | xlCommand);
        public static readonly int xlcFillLeft = (152 | xlCommand);
        public static readonly int xlcDeleteOverlay = (153 | xlCommand);
        public static readonly int xlcNote = (154 | xlCommand);
        public static readonly int xlcShortMenus = (155 | xlCommand);
        public static readonly int xlcSetUpdateStatus = (159 | xlCommand);
        public static readonly int xlcColorPalette = (161 | xlCommand);
        public static readonly int xlcDeleteStyle = (162 | xlCommand);
        public static readonly int xlcWindowRestore = (163 | xlCommand);
        public static readonly int xlcWindowMaximize = (164 | xlCommand);
        public static readonly int xlcError = (165 | xlCommand);
        public static readonly int xlcChangeLink = (166 | xlCommand);
        public static readonly int xlcCalculateDocument = (167 | xlCommand);
        public static readonly int xlcOnKey = (168 | xlCommand);
        public static readonly int xlcAppRestore = (169 | xlCommand);
        public static readonly int xlcAppMove = (170 | xlCommand);
        public static readonly int xlcAppSize = (171 | xlCommand);
        public static readonly int xlcAppMinimize = (172 | xlCommand);
        public static readonly int xlcAppMaximize = (173 | xlCommand);
        public static readonly int xlcBringToFront = (174 | xlCommand);
        public static readonly int xlcSendToBack = (175 | xlCommand);
        public static readonly int xlcMainChartType = (185 | xlCommand);
        public static readonly int xlcOverlayChartType = (186 | xlCommand);
        public static readonly int xlcSelectEnd = (187 | xlCommand);
        public static readonly int xlcOpenMail = (188 | xlCommand);
        public static readonly int xlcSendMail = (189 | xlCommand);
        public static readonly int xlcStandardFont = (190 | xlCommand);
        public static readonly int xlcConsolidate = (191 | xlCommand);
        public static readonly int xlcSortSpecial = (192 | xlCommand);
        public static readonly int xlcGallery3dArea = (193 | xlCommand);
        public static readonly int xlcGallery3dColumn = (194 | xlCommand);
        public static readonly int xlcGallery3dLine = (195 | xlCommand);
        public static readonly int xlcGallery3dPie = (196 | xlCommand);
        public static readonly int xlcView3d = (197 | xlCommand);
        public static readonly int xlcGoalSeek = (198 | xlCommand);
        public static readonly int xlcWorkgroup = (199 | xlCommand);
        public static readonly int xlcFillGroup = (200 | xlCommand);
        public static readonly int xlcUpdateLink = (201 | xlCommand);
        public static readonly int xlcPromote = (202 | xlCommand);
        public static readonly int xlcDemote = (203 | xlCommand);
        public static readonly int xlcShowDetail = (204 | xlCommand);
        public static readonly int xlcUngroup = (206 | xlCommand);
        public static readonly int xlcObjectProperties = (207 | xlCommand);
        public static readonly int xlcSaveNewObject = (208 | xlCommand);
        public static readonly int xlcShare = (209 | xlCommand);
        public static readonly int xlcShareName = (210 | xlCommand);
        public static readonly int xlcDuplicate = (211 | xlCommand);
        public static readonly int xlcApplyStyle = (212 | xlCommand);
        public static readonly int xlcAssignToObject = (213 | xlCommand);
        public static readonly int xlcObjectProtection = (214 | xlCommand);
        public static readonly int xlcHideObject = (215 | xlCommand);
        public static readonly int xlcSetExtract = (216 | xlCommand);
        public static readonly int xlcCreatePublisher = (217 | xlCommand);
        public static readonly int xlcSubscribeTo = (218 | xlCommand);
        public static readonly int xlcAttributes = (219 | xlCommand);
        public static readonly int xlcShowToolbar = (220 | xlCommand);
        public static readonly int xlcPrintPreview = (222 | xlCommand);
        public static readonly int xlcEditColor = (223 | xlCommand);
        public static readonly int xlcShowLevels = (224 | xlCommand);
        public static readonly int xlcFormatMain = (225 | xlCommand);
        public static readonly int xlcFormatOverlay = (226 | xlCommand);
        public static readonly int xlcOnRecalc = (227 | xlCommand);
        public static readonly int xlcEditSeries = (228 | xlCommand);
        public static readonly int xlcDefineStyle = (229 | xlCommand);
        public static readonly int xlcLinePrint = (240 | xlCommand);
        public static readonly int xlcEnterData = (243 | xlCommand);
        public static readonly int xlcGalleryRadar = (249 | xlCommand);
        public static readonly int xlcMergeStyles = (250 | xlCommand);
        public static readonly int xlcEditionOptions = (251 | xlCommand);
        public static readonly int xlcPastePicture = (252 | xlCommand);
        public static readonly int xlcPastePictureLink = (253 | xlCommand);
        public static readonly int xlcSpelling = (254 | xlCommand);
        public static readonly int xlcZoom = (256 | xlCommand);
        public static readonly int xlcResume = (258 | xlCommand);
        public static readonly int xlcInsertObject = (259 | xlCommand);
        public static readonly int xlcWindowMinimize = (260 | xlCommand);
        public static readonly int xlcSize = (261 | xlCommand);
        public static readonly int xlcMove = (262 | xlCommand);
        public static readonly int xlcSoundNote = (265 | xlCommand);
        public static readonly int xlcSoundPlay = (266 | xlCommand);
        public static readonly int xlcFormatShape = (267 | xlCommand);
        public static readonly int xlcExtendPolygon = (268 | xlCommand);
        public static readonly int xlcFormatAuto = (269 | xlCommand);
        public static readonly int xlcGallery3dBar = (272 | xlCommand);
        public static readonly int xlcGallery3dSurface = (273 | xlCommand);
        public static readonly int xlcFillAuto = (274 | xlCommand);
        public static readonly int xlcCustomizeToolbar = (276 | xlCommand);
        public static readonly int xlcAddTool = (277 | xlCommand);
        public static readonly int xlcEditObject = (278 | xlCommand);
        public static readonly int xlcOnDoubleclick = (279 | xlCommand);
        public static readonly int xlcOnEntry = (280 | xlCommand);
        public static readonly int xlcWorkbookAdd = (281 | xlCommand);
        public static readonly int xlcWorkbookMove = (282 | xlCommand);
        public static readonly int xlcWorkbookCopy = (283 | xlCommand);
        public static readonly int xlcWorkbookOptions = (284 | xlCommand);
        public static readonly int xlcSaveWorkspace = (285 | xlCommand);
        public static readonly int xlcChartWizard = (288 | xlCommand);
        public static readonly int xlcDeleteTool = (289 | xlCommand);
        public static readonly int xlcMoveTool = (290 | xlCommand);
        public static readonly int xlcWorkbookSelect = (291 | xlCommand);
        public static readonly int xlcWorkbookActivate = (292 | xlCommand);
        public static readonly int xlcAssignToTool = (293 | xlCommand);
        public static readonly int xlcCopyTool = (295 | xlCommand);
        public static readonly int xlcResetTool = (296 | xlCommand);
        public static readonly int xlcConstrainNumeric = (297 | xlCommand);
        public static readonly int xlcPasteTool = (298 | xlCommand);
        public static readonly int xlcPlacement = (300 | xlCommand);
        public static readonly int xlcFillWorkgroup = (301 | xlCommand);
        public static readonly int xlcWorkbookNew = (302 | xlCommand);
        public static readonly int xlcScenarioCells = (305 | xlCommand);
        public static readonly int xlcScenarioDelete = (306 | xlCommand);
        public static readonly int xlcScenarioAdd = (307 | xlCommand);
        public static readonly int xlcScenarioEdit = (308 | xlCommand);
        public static readonly int xlcScenarioShow = (309 | xlCommand);
        public static readonly int xlcScenarioShowNext = (310 | xlCommand);
        public static readonly int xlcScenarioSummary = (311 | xlCommand);
        public static readonly int xlcPivotTableWizard = (312 | xlCommand);
        public static readonly int xlcPivotFieldProperties = (313 | xlCommand);
        public static readonly int xlcPivotField = (314 | xlCommand);
        public static readonly int xlcPivotItem = (315 | xlCommand);
        public static readonly int xlcPivotAddFields = (316 | xlCommand);
        public static readonly int xlcOptionsCalculation = (318 | xlCommand);
        public static readonly int xlcOptionsEdit = (319 | xlCommand);
        public static readonly int xlcOptionsView = (320 | xlCommand);
        public static readonly int xlcAddinManager = (321 | xlCommand);
        public static readonly int xlcMenuEditor = (322 | xlCommand);
        public static readonly int xlcAttachToolbars = (323 | xlCommand);
        public static readonly int xlcVbaactivate = (324 | xlCommand);
        public static readonly int xlcOptionsChart = (325 | xlCommand);
        public static readonly int xlcVbaInsertFile = (328 | xlCommand);
        public static readonly int xlcVbaProcedureDefinition = (330 | xlCommand);
        public static readonly int xlcRoutingSlip = (336 | xlCommand);
        public static readonly int xlcRouteDocument = (338 | xlCommand);
        public static readonly int xlcMailLogon = (339 | xlCommand);
        public static readonly int xlcInsertPicture = (342 | xlCommand);
        public static readonly int xlcEditTool = (343 | xlCommand);
        public static readonly int xlcGalleryDoughnut = (344 | xlCommand);
        public static readonly int xlcChartTrend = (350 | xlCommand);
        public static readonly int xlcPivotItemProperties = (352 | xlCommand);
        public static readonly int xlcWorkbookInsert = (354 | xlCommand);
        public static readonly int xlcOptionsTransition = (355 | xlCommand);
        public static readonly int xlcOptionsGeneral = (356 | xlCommand);
        public static readonly int xlcFilterAdvanced = (370 | xlCommand);
        public static readonly int xlcMailAddMailer = (373 | xlCommand);
        public static readonly int xlcMailDeleteMailer = (374 | xlCommand);
        public static readonly int xlcMailReply = (375 | xlCommand);
        public static readonly int xlcMailReplyAll = (376 | xlCommand);
        public static readonly int xlcMailForward = (377 | xlCommand);
        public static readonly int xlcMailNextLetter = (378 | xlCommand);
        public static readonly int xlcDataLabel = (379 | xlCommand);
        public static readonly int xlcInsertTitle = (380 | xlCommand);
        public static readonly int xlcFontProperties = (381 | xlCommand);
        public static readonly int xlcMacroOptions = (382 | xlCommand);
        public static readonly int xlcWorkbookHide = (383 | xlCommand);
        public static readonly int xlcWorkbookUnhide = (384 | xlCommand);
        public static readonly int xlcWorkbookDelete = (385 | xlCommand);
        public static readonly int xlcWorkbookName = (386 | xlCommand);
        public static readonly int xlcGalleryCustom = (388 | xlCommand);
        public static readonly int xlcAddChartAutoformat = (390 | xlCommand);
        public static readonly int xlcDeleteChartAutoformat = (391 | xlCommand);
        public static readonly int xlcChartAddData = (392 | xlCommand);
        public static readonly int xlcAutoOutline = (393 | xlCommand);
        public static readonly int xlcTabOrder = (394 | xlCommand);
        public static readonly int xlcShowDialog = (395 | xlCommand);
        public static readonly int xlcSelectAll = (396 | xlCommand);
        public static readonly int xlcUngroupSheets = (397 | xlCommand);
        public static readonly int xlcSubtotalCreate = (398 | xlCommand);
        public static readonly int xlcSubtotalRemove = (399 | xlCommand);
        public static readonly int xlcRenameObject = (400 | xlCommand);
        public static readonly int xlcWorkbookScroll = (412 | xlCommand);
        public static readonly int xlcWorkbookNext = (413 | xlCommand);
        public static readonly int xlcWorkbookPrev = (414 | xlCommand);
        public static readonly int xlcWorkbookTabSplit = (415 | xlCommand);
        public static readonly int xlcFullScreen = (416 | xlCommand);
        public static readonly int xlcWorkbookProtect = (417 | xlCommand);
        public static readonly int xlcScrollbarProperties = (420 | xlCommand);
        public static readonly int xlcPivotShowPages = (421 | xlCommand);
        public static readonly int xlcTextToColumns = (422 | xlCommand);
        public static readonly int xlcFormatCharttype = (423 | xlCommand);
        public static readonly int xlcLinkFormat = (424 | xlCommand);
        public static readonly int xlcTracerDisplay = (425 | xlCommand);
        public static readonly int xlcTracerNavigate = (430 | xlCommand);
        public static readonly int xlcTracerClear = (431 | xlCommand);
        public static readonly int xlcTracerError = (432 | xlCommand);
        public static readonly int xlcPivotFieldGroup = (433 | xlCommand);
        public static readonly int xlcPivotFieldUngroup = (434 | xlCommand);
        public static readonly int xlcCheckboxProperties = (435 | xlCommand);
        public static readonly int xlcLabelProperties = (436 | xlCommand);
        public static readonly int xlcListboxProperties = (437 | xlCommand);
        public static readonly int xlcEditboxProperties = (438 | xlCommand);
        public static readonly int xlcPivotRefresh = (439 | xlCommand);
        public static readonly int xlcLinkCombo = (440 | xlCommand);
        public static readonly int xlcOpenText = (441 | xlCommand);
        public static readonly int xlcHideDialog = (442 | xlCommand);
        public static readonly int xlcSetDialogFocus = (443 | xlCommand);
        public static readonly int xlcEnableObject = (444 | xlCommand);
        public static readonly int xlcPushbuttonProperties = (445 | xlCommand);
        public static readonly int xlcSetDialogDefault = (446 | xlCommand);
        public static readonly int xlcFilter = (447 | xlCommand);
        public static readonly int xlcFilterShowAll = (448 | xlCommand);
        public static readonly int xlcClearOutline = (449 | xlCommand);
        public static readonly int xlcFunctionWizard = (450 | xlCommand);
        public static readonly int xlcAddListItem = (451 | xlCommand);
        public static readonly int xlcSetListItem = (452 | xlCommand);
        public static readonly int xlcRemoveListItem = (453 | xlCommand);
        public static readonly int xlcSelectListItem = (454 | xlCommand);
        public static readonly int xlcSetControlValue = (455 | xlCommand);
        public static readonly int xlcSaveCopyAs = (456 | xlCommand);
        public static readonly int xlcOptionsListsAdd = (458 | xlCommand);
        public static readonly int xlcOptionsListsDelete = (459 | xlCommand);
        public static readonly int xlcSeriesAxes = (460 | xlCommand);
        public static readonly int xlcSeriesX = (461 | xlCommand);
        public static readonly int xlcSeriesY = (462 | xlCommand);
        public static readonly int xlcErrorbarX = (463 | xlCommand);
        public static readonly int xlcErrorbarY = (464 | xlCommand);
        public static readonly int xlcFormatChart = (465 | xlCommand);
        public static readonly int xlcSeriesOrder = (466 | xlCommand);
        public static readonly int xlcMailLogoff = (467 | xlCommand);
        public static readonly int xlcClearRoutingSlip = (468 | xlCommand);
        public static readonly int xlcAppActivateMicrosoft = (469 | xlCommand);
        public static readonly int xlcMailEditMailer = (470 | xlCommand);
        public static readonly int xlcOnSheet = (471 | xlCommand);
        public static readonly int xlcStandardWidth = (472 | xlCommand);
        public static readonly int xlcScenarioMerge = (473 | xlCommand);
        public static readonly int xlcSummaryInfo = (474 | xlCommand);
        public static readonly int xlcFindFile = (475 | xlCommand);
        public static readonly int xlcActiveCellFont = (476 | xlCommand);
        public static readonly int xlcEnableTipwizard = (477 | xlCommand);
        public static readonly int xlcVbaMakeAddin = (478 | xlCommand);
        public static readonly int xlcInsertdatatable = (480 | xlCommand);
        public static readonly int xlcWorkgroupOptions = (481 | xlCommand);
        public static readonly int xlcMailSendMailer = (482 | xlCommand);
        public static readonly int xlcAutocorrect = (485 | xlCommand);
        public static readonly int xlcPostDocument = (489 | xlCommand);
        public static readonly int xlcPicklist = (491 | xlCommand);
        public static readonly int xlcViewShow = (493 | xlCommand);
        public static readonly int xlcViewDefine = (494 | xlCommand);
        public static readonly int xlcViewDelete = (495 | xlCommand);
        public static readonly int xlcSheetBackground = (509 | xlCommand);
        public static readonly int xlcInsertMapObject = (510 | xlCommand);
        public static readonly int xlcOptionsMenono = (511 | xlCommand);
        public static readonly int xlcNormal = (518 | xlCommand);
        public static readonly int xlcLayout = (519 | xlCommand);
        public static readonly int xlcRmPrintArea = (520 | xlCommand);
        public static readonly int xlcClearPrintArea = (521 | xlCommand);
        public static readonly int xlcAddPrintArea = (522 | xlCommand);
        public static readonly int xlcMoveBrk = (523 | xlCommand);
        public static readonly int xlcHidecurrNote = (545 | xlCommand);
        public static readonly int xlcHideallNotes = (546 | xlCommand);
        public static readonly int xlcDeleteNote = (547 | xlCommand);
        public static readonly int xlcTraverseNotes = (548 | xlCommand);
        public static readonly int xlcActivateNotes = (549 | xlCommand);
        public static readonly int xlcProtectRevisions = (620 | xlCommand);
        public static readonly int xlcUnprotectRevisions = (621 | xlCommand);
        public static readonly int xlcOptionsMe = (647 | xlCommand);
        public static readonly int xlcWebPublish = (653 | xlCommand);
        public static readonly int xlcNewwebquery = (667 | xlCommand);
        public static readonly int xlcPivotTableChart = (673 | xlCommand);
        public static readonly int xlcOptionsSave = (753 | xlCommand);
        public static readonly int xlcOptionsSpell = (755 | xlCommand);
        public static readonly int xlcHideallInkannots = (808 | xlCommand);

        #endregion


		// THROWS: XlCallException if anything goes wrong.
		public static object Excel(int xlFunction, params object[] parameters )
		{
			object result;
			XlReturn xlReturn = TryExcel(xlFunction, out result, parameters);

			if (xlReturn == XlReturn.XlReturnSuccess)
			{
					return result;
			}
			else
			{
					throw new XlCallException(xlReturn);
			}
		}

		public static XlReturn TryExcel(int xlFunction, out object result, params object[] parameters)
		{
            if (_suspended)
            {
                result = null;
                return XlReturn.XlReturnFailed;
            }
            return ExcelIntegration.TryExcelImpl(xlFunction, out result, parameters);
        }

        /// <summary>
        /// Supports the registration-free RTD service.
        /// </summary>
        /// <param name="progId">The ProgId or type name of the RTD server.</param>
        /// <param name="server">not used</param>
        /// <param name="topics">strings passed to the ConnectData call.</param>
        /// <returns></returns>
        public static object RTD(string progId, string server, params string[] topics)
        {
            return Rtd.RtdRegistration.RTD(progId, server, topics);
        }

        // Support for suspending calls to the C API
        // Used in the RTD Server wrapper - otherwise C API calls from the RTD methods can crash Excel.
        static bool _suspended = false;

        internal static IDisposable Suspend()
        {
            return new XlCallSuspended();
        }

        class XlCallSuspended : IDisposable
        {
            public XlCallSuspended()
            {
                _suspended = true;
            }

            public void Dispose()
            {
                _suspended = false;
            }
        }

    }

	public class XlCallException : Exception
	{
		public XlCall.XlReturn xlReturn;

		public XlCallException(XlCall.XlReturn xlReturn)
		{
			this.xlReturn = xlReturn;
		}
	}
}

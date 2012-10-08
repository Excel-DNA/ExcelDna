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

#include "stdafx.h"
#include "ExcelDna.h"
#include "ExcelDnaLoader.h"

// Minimal parts of XLOPER types, 
// used only for xlAddInManagerInfo(12). Really.
struct XLOPER
{
	union
	{
		double num;					/* xltypeNum */
		void* str;					/* xltypeStr */
		WORD err;					/* xltypeErr */
	} val;
	WORD xltype; // Should be at offset 8 bytes
};

struct XLOPER12
{
	union
	{
		double num;					/* xltypeNum */
		void* str;					/* xltypeStr */
		int err;					/* xltypeErr */
		struct
		{
			double unused1;
			double unused2;
			double unused3;
		} unused;
	} val;
	DWORD xltype; // Should be at offset 24 bytes
};
#define xltypeNum  1;
#define xltypeStr  2;
#define xltypeErr  16;
#define xlerrValue 15

// The one and only ExportInfo
XlAddInExportInfo* pExportInfo = NULL;
// Flag to coordinate load/unload close and remove.
HMODULE lockModule;
bool locked = false;
bool removed = false;    // Used to check whether AutoRemove is called before AutoClose.
bool autoOpened = false; // Not set when loaded for COM server only. Used for re-open check.

// The actual thunk table 
extern "C" 
{
	PFN thunks[EXPORT_COUNT];
}

XlAddInExportInfo* CreateExportInfo()
{
	pExportInfo = new XlAddInExportInfo();
	pExportInfo->ExportInfoVersion = 6;
	pExportInfo->AppDomainId = -1;
	pExportInfo->pXlAutoOpen = NULL;
	pExportInfo->pXlAutoClose = NULL;
	pExportInfo->pXlAutoRemove = NULL;
	pExportInfo->pXlAutoFree = NULL;
	pExportInfo->pXlAutoFree12 = NULL;
	pExportInfo->pSetExcel12EntryPt = NULL;
	pExportInfo->pDllRegisterServer = NULL;
	pExportInfo->pDllUnregisterServer = NULL;
	pExportInfo->pDllGetClassObject = NULL;
	pExportInfo->pDllCanUnloadNow = NULL;
	pExportInfo->pSyncMacro = NULL;
	pExportInfo->ThunkTableLength = EXPORT_COUNT;
	pExportInfo->ThunkTable = (PFN*)thunks;
	return pExportInfo;
}

// Safe to be called repeatedly, but not from multiple threads
short EnsureInitialized()
{
	short result = 0;
	if (pExportInfo != NULL)
	{
		result = 1;
	}
	else
	{
		XlAddInExportInfo* pExportInfoTemp = CreateExportInfo();
		result = XlLibraryInitialize(pExportInfoTemp);
		if (result)
		{
			pExportInfo	= pExportInfoTemp;
		}
	}
	return result;
}

// Called only when AutoClose is called after AutoRemove.
void Uninitialize()
{
	delete pExportInfo;
	pExportInfo = NULL;
	for (int i = 0; i < EXPORT_COUNT; i++)
	{
		thunks[i] = NULL;
	}
}

// Ensure that the library stays loaded.
// May be called many times, but should keep opened only until unlocked once.
void LockModule()
{
	if (!locked)
	{
		CPath xllPath(GetAddInFullPath());
		xllPath.StripPath();
		lockModule = LoadLibrary(xllPath);
		locked = true;
	}
}

// Allow the library to be unloaded.
void UnlockModule()
{
	if (locked)
	{
		FreeLibrary(lockModule);
		locked = false;
	}
}

// Standard DLL entry point.
BOOL __stdcall DllMain( HMODULE hModule,
						DWORD  ul_reason_for_call,
						LPVOID lpReserved
						)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		LoaderInitialize(hModule);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		LoaderUnload();
		break;
	}
	return TRUE;
}

extern "C"
{
	// Forward declares, since these are now called by AutoOpen.
	short __stdcall xlAutoClose();
	short __stdcall xlAutoRemove();

	// Excel Add-In standard exports
	short __stdcall xlAutoOpen()
	{
		short result = 0;

		// If we are loaded as an add-in already, then ensure re-load = AddInRemove + AutoClose + AutoOpen,
		// which mains a clean AppDomain for each load.
		if (autoOpened)
		{
			xlAutoRemove();
			xlAutoClose();
		}

		if (EnsureInitialized() && 
			pExportInfo->pXlAutoOpen != NULL)
		{
			result = pExportInfo->pXlAutoOpen();
			LockModule();
			// Set the 'removed' flag to false, which prevents AutoClose from actually unloading (or calling through to the add-in),
			// unless AutoRemove is called first (from the add-in manager, a host or the re-open sequence above).
			removed = false;
			// Keep track that we are loaded as an add-in, not just a COM or RTD server.
			// This allows us to re-open in a clean AppDomain, yet load COM server first then add-in without damage.
			autoOpened = true;
		}
		return result;
	}

	short __stdcall xlAutoClose()
	{
		short result = 0;
		if (EnsureInitialized() && 
			pExportInfo->pXlAutoClose != NULL)
		{
			result = pExportInfo->pXlAutoClose();
			if (removed)
			{
				// TODO: Consider how and when to unload
				//       Unloading the AppDomain could be a bit too dramatic if we are serving as a COM Server or RTD Server directly.
				// DOCUMENT: What the current implementation is.
				// No more managed functions should be called.
				Uninitialize();

				// Complete the clean-up by unloading AppDomain
				XlLibraryUnload();
				// ... recording that we are no longer open as an add-in.
				autoOpened = false;
				// ...and allowing the .xll itself to be unloaded
				UnlockModule();
			}
		}
		return result;
	}
	
	// Since v0.29 loading is much more expensive, so I want to reduce the number of times we load.
	// We've never used or exposed xlAutoAdd to Excel-DNA addins, so no harm in disabling for now.
	// To add back, also uncomment in the ExcelDna.def file.
	//short __stdcall xlAutoAdd()
	//{
	//	short result = 0;
	//	if (EnsureInitialized() && 
	//		pExportInfo->pXlAutoAdd != NULL)
	//	{
	//		result = pExportInfo->pXlAutoAdd();
	//	}
	//	return result;
	//}

	short __stdcall xlAutoRemove()
	{
		short result = 0;
		if (EnsureInitialized() && 
			pExportInfo->pXlAutoRemove != NULL)
		{
			result = pExportInfo->pXlAutoRemove();
			// Set the 'removed' flag which will allow the AutoClose to actually unload (and call through to the add-in).
			removed = true;
		}
		return result;
	}

	void __stdcall xlAutoFree(void* pXloper)
	{
		if (pExportInfo != NULL && pExportInfo->pXlAutoFree != NULL)
		{
			pExportInfo->pXlAutoFree(pXloper);
		}
	}

	void __stdcall xlAutoFree12(void* pXloper12)
	{
		if (pExportInfo != NULL && pExportInfo->pXlAutoFree12 != NULL)
		{
			pExportInfo->pXlAutoFree12(pXloper12);
		}
	}

	XLOPER* __stdcall xlAddInManagerInfo(XLOPER* pXloper)
	{
		static XLOPER result;
		static char name[256];

		// Return error by default
		result.xltype = xltypeErr;
		result.val.err = xlerrValue;

		if (pXloper->xltype == 1 && pXloper->val.num == 1.0)
		{
			CString addInNameW;
			HRESULT hr = GetAddInName(addInNameW);
			if (!FAILED(hr))
			{
				CStringA addInName(addInNameW);
				byte length = (byte)min(addInName.GetLength(), 255);
				name[0] = (char)length;
				const char* pAddInName = addInName;
				char* pName = (char*)name + 1;
				CStringA::CopyChars(pName, 255, pAddInName, length);

				result.xltype = xltypeStr;
				result.val.str = name;
			}
		}

		return &result;
	}

	XLOPER12* __stdcall xlAddInManagerInfo12(XLOPER12* pXloper)
	{
		static XLOPER12 result;
		static wchar_t name[256];

		// Return error by default
		result.xltype = xltypeErr;
		result.val.err = xlerrValue;

		if (pXloper->xltype == 1 && pXloper->val.num == 1.0)
		{
			CString addInName;
			HRESULT hr = GetAddInName(addInName);
			if (!FAILED(hr))
			{
				// We could probably use CString as is (maybe with truncation)!?
				int length = (int)min(addInName.GetLength(), 255);
				name[0] = (wchar_t)length;
				const wchar_t* pAddInName = addInName;
				wchar_t* pName = (wchar_t*)name + 1;
				CString::CopyChars(pName, 255, pAddInName, length);
				result.xltype = xltypeStr;
				result.val.str = name;
			}
		}

		return &result;
	}

	// Support for Excel 2010 SDK - used when loading under HPC XLL Host
	void __stdcall SetExcel12EntryPt(void* pexcel12New)
	{
		if (EnsureInitialized() && 
			pExportInfo->pSetExcel12EntryPt != NULL)
		{
			pExportInfo->pSetExcel12EntryPt(pexcel12New);
		}
	}

	// We are also a COM Server, to support the =RTD(...) worksheet function and VBA ComServer integration.
	HRESULT __stdcall DllRegisterServer()
	{
		HRESULT result = E_UNEXPECTED;
		if (EnsureInitialized() && 
			pExportInfo->pDllRegisterServer != NULL)
		{
			result = pExportInfo->pDllRegisterServer();
		}
		return result;
	}

	HRESULT __stdcall DllUnregisterServer()
	{
		HRESULT result = E_UNEXPECTED;
		if (EnsureInitialized() && 
			pExportInfo->pDllUnregisterServer != NULL)
		{
			result = pExportInfo->pDllUnregisterServer();
		}
		return result;
	}
	
	HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID iid, void** ppv)
	{
		HRESULT result = E_UNEXPECTED;
		GUID cls = clsid;
		GUID i = iid;
		if (EnsureInitialized() && 
			pExportInfo->pDllGetClassObject != NULL)
		{

			result = pExportInfo->pDllGetClassObject(cls, i, ppv);
		}
		return result;
	}

	HRESULT __stdcall DllCanUnloadNow()
	{
		HRESULT result = S_OK;
		if (EnsureInitialized() && 
			pExportInfo->pDllCanUnloadNow != NULL)
		{
			result = pExportInfo->pDllCanUnloadNow();
		}
		return result;
	}

	void __stdcall SyncMacro(double param)
	{
		if (EnsureInitialized() && 
			pExportInfo->pSyncMacro != NULL)
		{
			pExportInfo->pSyncMacro(param);
		}
	}
}


#ifndef _M_X64
// The dll export implementation that jmps to thunk in the thunktable
// For x64 this is implemented in JmpExports64.asm

// Use extern so that functions are not decorated when exported.
// naked ensures no prologue or epilogue generated by the compiler 
// - jump directly to unmanaged thunk at offset i
#define expf(i) extern "C" __declspec(dllexport,naked) void f##i(void){	__asm jmp thunks + i * 4 /* sizeof(PFN) (only used on 32-bit) */ }

// Declare the functions -- NOTE: list here must go from 0 to EXPORT_COUNT-1
expf(0)
expf(1)
expf(2)
expf(3)
expf(4)
expf(5)
expf(6)
expf(7)
expf(8)
expf(9)
expf(10)
expf(11)
expf(12)
expf(13)
expf(14)
expf(15)
expf(16)
expf(17)
expf(18)
expf(19)
expf(20)
expf(21)
expf(22)
expf(23)
expf(24)
expf(25)
expf(26)
expf(27)
expf(28)
expf(29)
expf(30)
expf(31)
expf(32)
expf(33)
expf(34)
expf(35)
expf(36)
expf(37)
expf(38)
expf(39)
expf(40)
expf(41)
expf(42)
expf(43)
expf(44)
expf(45)
expf(46)
expf(47)
expf(48)
expf(49)
expf(50)
expf(51)
expf(52)
expf(53)
expf(54)
expf(55)
expf(56)
expf(57)
expf(58)
expf(59)
expf(60)
expf(61)
expf(62)
expf(63)
expf(64)
expf(65)
expf(66)
expf(67)
expf(68)
expf(69)
expf(70)
expf(71)
expf(72)
expf(73)
expf(74)
expf(75)
expf(76)
expf(77)
expf(78)
expf(79)
expf(80)
expf(81)
expf(82)
expf(83)
expf(84)
expf(85)
expf(86)
expf(87)
expf(88)
expf(89)
expf(90)
expf(91)
expf(92)
expf(93)
expf(94)
expf(95)
expf(96)
expf(97)
expf(98)
expf(99)
expf(100)
expf(101)
expf(102)
expf(103)
expf(104)
expf(105)
expf(106)
expf(107)
expf(108)
expf(109)
expf(110)
expf(111)
expf(112)
expf(113)
expf(114)
expf(115)
expf(116)
expf(117)
expf(118)
expf(119)
expf(120)
expf(121)
expf(122)
expf(123)
expf(124)
expf(125)
expf(126)
expf(127)
expf(128)
expf(129)
expf(130)
expf(131)
expf(132)
expf(133)
expf(134)
expf(135)
expf(136)
expf(137)
expf(138)
expf(139)
expf(140)
expf(141)
expf(142)
expf(143)
expf(144)
expf(145)
expf(146)
expf(147)
expf(148)
expf(149)
expf(150)
expf(151)
expf(152)
expf(153)
expf(154)
expf(155)
expf(156)
expf(157)
expf(158)
expf(159)
expf(160)
expf(161)
expf(162)
expf(163)
expf(164)
expf(165)
expf(166)
expf(167)
expf(168)
expf(169)
expf(170)
expf(171)
expf(172)
expf(173)
expf(174)
expf(175)
expf(176)
expf(177)
expf(178)
expf(179)
expf(180)
expf(181)
expf(182)
expf(183)
expf(184)
expf(185)
expf(186)
expf(187)
expf(188)
expf(189)
expf(190)
expf(191)
expf(192)
expf(193)
expf(194)
expf(195)
expf(196)
expf(197)
expf(198)
expf(199)
expf(200)
expf(201)
expf(202)
expf(203)
expf(204)
expf(205)
expf(206)
expf(207)
expf(208)
expf(209)
expf(210)
expf(211)
expf(212)
expf(213)
expf(214)
expf(215)
expf(216)
expf(217)
expf(218)
expf(219)
expf(220)
expf(221)
expf(222)
expf(223)
expf(224)
expf(225)
expf(226)
expf(227)
expf(228)
expf(229)
expf(230)
expf(231)
expf(232)
expf(233)
expf(234)
expf(235)
expf(236)
expf(237)
expf(238)
expf(239)
expf(240)
expf(241)
expf(242)
expf(243)
expf(244)
expf(245)
expf(246)
expf(247)
expf(248)
expf(249)
expf(250)
expf(251)
expf(252)
expf(253)
expf(254)
expf(255)
expf(256)
expf(257)
expf(258)
expf(259)
expf(260)
expf(261)
expf(262)
expf(263)
expf(264)
expf(265)
expf(266)
expf(267)
expf(268)
expf(269)
expf(270)
expf(271)
expf(272)
expf(273)
expf(274)
expf(275)
expf(276)
expf(277)
expf(278)
expf(279)
expf(280)
expf(281)
expf(282)
expf(283)
expf(284)
expf(285)
expf(286)
expf(287)
expf(288)
expf(289)
expf(290)
expf(291)
expf(292)
expf(293)
expf(294)
expf(295)
expf(296)
expf(297)
expf(298)
expf(299)
expf(300)
expf(301)
expf(302)
expf(303)
expf(304)
expf(305)
expf(306)
expf(307)
expf(308)
expf(309)
expf(310)
expf(311)
expf(312)
expf(313)
expf(314)
expf(315)
expf(316)
expf(317)
expf(318)
expf(319)
expf(320)
expf(321)
expf(322)
expf(323)
expf(324)
expf(325)
expf(326)
expf(327)
expf(328)
expf(329)
expf(330)
expf(331)
expf(332)
expf(333)
expf(334)
expf(335)
expf(336)
expf(337)
expf(338)
expf(339)
expf(340)
expf(341)
expf(342)
expf(343)
expf(344)
expf(345)
expf(346)
expf(347)
expf(348)
expf(349)
expf(350)
expf(351)
expf(352)
expf(353)
expf(354)
expf(355)
expf(356)
expf(357)
expf(358)
expf(359)
expf(360)
expf(361)
expf(362)
expf(363)
expf(364)
expf(365)
expf(366)
expf(367)
expf(368)
expf(369)
expf(370)
expf(371)
expf(372)
expf(373)
expf(374)
expf(375)
expf(376)
expf(377)
expf(378)
expf(379)
expf(380)
expf(381)
expf(382)
expf(383)
expf(384)
expf(385)
expf(386)
expf(387)
expf(388)
expf(389)
expf(390)
expf(391)
expf(392)
expf(393)
expf(394)
expf(395)
expf(396)
expf(397)
expf(398)
expf(399)
expf(400)
expf(401)
expf(402)
expf(403)
expf(404)
expf(405)
expf(406)
expf(407)
expf(408)
expf(409)
expf(410)
expf(411)
expf(412)
expf(413)
expf(414)
expf(415)
expf(416)
expf(417)
expf(418)
expf(419)
expf(420)
expf(421)
expf(422)
expf(423)
expf(424)
expf(425)
expf(426)
expf(427)
expf(428)
expf(429)
expf(430)
expf(431)
expf(432)
expf(433)
expf(434)
expf(435)
expf(436)
expf(437)
expf(438)
expf(439)
expf(440)
expf(441)
expf(442)
expf(443)
expf(444)
expf(445)
expf(446)
expf(447)
expf(448)
expf(449)
expf(450)
expf(451)
expf(452)
expf(453)
expf(454)
expf(455)
expf(456)
expf(457)
expf(458)
expf(459)
expf(460)
expf(461)
expf(462)
expf(463)
expf(464)
expf(465)
expf(466)
expf(467)
expf(468)
expf(469)
expf(470)
expf(471)
expf(472)
expf(473)
expf(474)
expf(475)
expf(476)
expf(477)
expf(478)
expf(479)
expf(480)
expf(481)
expf(482)
expf(483)
expf(484)
expf(485)
expf(486)
expf(487)
expf(488)
expf(489)
expf(490)
expf(491)
expf(492)
expf(493)
expf(494)
expf(495)
expf(496)
expf(497)
expf(498)
expf(499)
expf(500)
expf(501)
expf(502)
expf(503)
expf(504)
expf(505)
expf(506)
expf(507)
expf(508)
expf(509)
expf(510)
expf(511)
expf(512)
expf(513)
expf(514)
expf(515)
expf(516)
expf(517)
expf(518)
expf(519)
expf(520)
expf(521)
expf(522)
expf(523)
expf(524)
expf(525)
expf(526)
expf(527)
expf(528)
expf(529)
expf(530)
expf(531)
expf(532)
expf(533)
expf(534)
expf(535)
expf(536)
expf(537)
expf(538)
expf(539)
expf(540)
expf(541)
expf(542)
expf(543)
expf(544)
expf(545)
expf(546)
expf(547)
expf(548)
expf(549)
expf(550)
expf(551)
expf(552)
expf(553)
expf(554)
expf(555)
expf(556)
expf(557)
expf(558)
expf(559)
expf(560)
expf(561)
expf(562)
expf(563)
expf(564)
expf(565)
expf(566)
expf(567)
expf(568)
expf(569)
expf(570)
expf(571)
expf(572)
expf(573)
expf(574)
expf(575)
expf(576)
expf(577)
expf(578)
expf(579)
expf(580)
expf(581)
expf(582)
expf(583)
expf(584)
expf(585)
expf(586)
expf(587)
expf(588)
expf(589)
expf(590)
expf(591)
expf(592)
expf(593)
expf(594)
expf(595)
expf(596)
expf(597)
expf(598)
expf(599)
expf(600)
expf(601)
expf(602)
expf(603)
expf(604)
expf(605)
expf(606)
expf(607)
expf(608)
expf(609)
expf(610)
expf(611)
expf(612)
expf(613)
expf(614)
expf(615)
expf(616)
expf(617)
expf(618)
expf(619)
expf(620)
expf(621)
expf(622)
expf(623)
expf(624)
expf(625)
expf(626)
expf(627)
expf(628)
expf(629)
expf(630)
expf(631)
expf(632)
expf(633)
expf(634)
expf(635)
expf(636)
expf(637)
expf(638)
expf(639)
expf(640)
expf(641)
expf(642)
expf(643)
expf(644)
expf(645)
expf(646)
expf(647)
expf(648)
expf(649)
expf(650)
expf(651)
expf(652)
expf(653)
expf(654)
expf(655)
expf(656)
expf(657)
expf(658)
expf(659)
expf(660)
expf(661)
expf(662)
expf(663)
expf(664)
expf(665)
expf(666)
expf(667)
expf(668)
expf(669)
expf(670)
expf(671)
expf(672)
expf(673)
expf(674)
expf(675)
expf(676)
expf(677)
expf(678)
expf(679)
expf(680)
expf(681)
expf(682)
expf(683)
expf(684)
expf(685)
expf(686)
expf(687)
expf(688)
expf(689)
expf(690)
expf(691)
expf(692)
expf(693)
expf(694)
expf(695)
expf(696)
expf(697)
expf(698)
expf(699)
expf(700)
expf(701)
expf(702)
expf(703)
expf(704)
expf(705)
expf(706)
expf(707)
expf(708)
expf(709)
expf(710)
expf(711)
expf(712)
expf(713)
expf(714)
expf(715)
expf(716)
expf(717)
expf(718)
expf(719)
expf(720)
expf(721)
expf(722)
expf(723)
expf(724)
expf(725)
expf(726)
expf(727)
expf(728)
expf(729)
expf(730)
expf(731)
expf(732)
expf(733)
expf(734)
expf(735)
expf(736)
expf(737)
expf(738)
expf(739)
expf(740)
expf(741)
expf(742)
expf(743)
expf(744)
expf(745)
expf(746)
expf(747)
expf(748)
expf(749)
expf(750)
expf(751)
expf(752)
expf(753)
expf(754)
expf(755)
expf(756)
expf(757)
expf(758)
expf(759)
expf(760)
expf(761)
expf(762)
expf(763)
expf(764)
expf(765)
expf(766)
expf(767)
expf(768)
expf(769)
expf(770)
expf(771)
expf(772)
expf(773)
expf(774)
expf(775)
expf(776)
expf(777)
expf(778)
expf(779)
expf(780)
expf(781)
expf(782)
expf(783)
expf(784)
expf(785)
expf(786)
expf(787)
expf(788)
expf(789)
expf(790)
expf(791)
expf(792)
expf(793)
expf(794)
expf(795)
expf(796)
expf(797)
expf(798)
expf(799)
expf(800)
expf(801)
expf(802)
expf(803)
expf(804)
expf(805)
expf(806)
expf(807)
expf(808)
expf(809)
expf(810)
expf(811)
expf(812)
expf(813)
expf(814)
expf(815)
expf(816)
expf(817)
expf(818)
expf(819)
expf(820)
expf(821)
expf(822)
expf(823)
expf(824)
expf(825)
expf(826)
expf(827)
expf(828)
expf(829)
expf(830)
expf(831)
expf(832)
expf(833)
expf(834)
expf(835)
expf(836)
expf(837)
expf(838)
expf(839)
expf(840)
expf(841)
expf(842)
expf(843)
expf(844)
expf(845)
expf(846)
expf(847)
expf(848)
expf(849)
expf(850)
expf(851)
expf(852)
expf(853)
expf(854)
expf(855)
expf(856)
expf(857)
expf(858)
expf(859)
expf(860)
expf(861)
expf(862)
expf(863)
expf(864)
expf(865)
expf(866)
expf(867)
expf(868)
expf(869)
expf(870)
expf(871)
expf(872)
expf(873)
expf(874)
expf(875)
expf(876)
expf(877)
expf(878)
expf(879)
expf(880)
expf(881)
expf(882)
expf(883)
expf(884)
expf(885)
expf(886)
expf(887)
expf(888)
expf(889)
expf(890)
expf(891)
expf(892)
expf(893)
expf(894)
expf(895)
expf(896)
expf(897)
expf(898)
expf(899)
expf(900)
expf(901)
expf(902)
expf(903)
expf(904)
expf(905)
expf(906)
expf(907)
expf(908)
expf(909)
expf(910)
expf(911)
expf(912)
expf(913)
expf(914)
expf(915)
expf(916)
expf(917)
expf(918)
expf(919)
expf(920)
expf(921)
expf(922)
expf(923)
expf(924)
expf(925)
expf(926)
expf(927)
expf(928)
expf(929)
expf(930)
expf(931)
expf(932)
expf(933)
expf(934)
expf(935)
expf(936)
expf(937)
expf(938)
expf(939)
expf(940)
expf(941)
expf(942)
expf(943)
expf(944)
expf(945)
expf(946)
expf(947)
expf(948)
expf(949)
expf(950)
expf(951)
expf(952)
expf(953)
expf(954)
expf(955)
expf(956)
expf(957)
expf(958)
expf(959)
expf(960)
expf(961)
expf(962)
expf(963)
expf(964)
expf(965)
expf(966)
expf(967)
expf(968)
expf(969)
expf(970)
expf(971)
expf(972)
expf(973)
expf(974)
expf(975)
expf(976)
expf(977)
expf(978)
expf(979)
expf(980)
expf(981)
expf(982)
expf(983)
expf(984)
expf(985)
expf(986)
expf(987)
expf(988)
expf(989)
expf(990)
expf(991)
expf(992)
expf(993)
expf(994)
expf(995)
expf(996)
expf(997)
expf(998)
expf(999)
#endif

// DispIDs.h: Defines a DISPID for each COM property and method we're providing.

// ISysLink
// properties
#define DISPID_SL_APPEARANCE																	1
#define DISPID_SL_AUTOMATICALLYMARKLINKSASVISITED							2
#define DISPID_SL_BACKCOLOR																		3
#define DISPID_SL_BACKSTYLE																		4
#define DISPID_SL_BORDERSTYLE																	5
#define DISPID_SL_CARETLINK																		6
#define DISPID_SL_DETECTDOUBLECLICKS													7
#define DISPID_SL_DISABLEDEVENTS															8
#define DISPID_SL_DONTREDRAW																	9
#define DISPID_SL_ENABLED																			DISPID_ENABLED
#define DISPID_SL_FONT																				10
#define DISPID_SL_FORECOLOR																		11
#define DISPID_SL_HALIGNMENT																	12
#define DISPID_SL_HOTTRACKING																	13
#define DISPID_SL_HOVERTIME																		14
#define DISPID_SL_HWND																				15
#define DISPID_SL_HWNDTOOLTIP																	16
#define DISPID_SL_LINKS																				17
#define DISPID_SL_MOUSEICON																		18
#define DISPID_SL_MOUSEPOINTER																19
#define DISPID_SL_PROCESSCONTEXTMENUKEYS											20
#define DISPID_SL_PROCESSRETURNKEY														21
#define DISPID_SL_REGISTERFOROLEDRAGDROP											22
#define DISPID_SL_RIGHTTOLEFT																	23
#define DISPID_SL_SHOWTOOLTIPS																24
#define DISPID_SL_SUPPORTOLEDRAGIMAGES												25
#define DISPID_SL_TEXT																				DISPID_VALUE
#define DISPID_SL_USEMNEMONIC																	26
#define DISPID_SL_USESYSTEMFONT																27
#define DISPID_SL_USEVISUALSTYLE															28
#define DISPID_SL_VERSION																			29
// fingerprint
#define DISPID_SL_APPID																				500
#define DISPID_SL_APPNAME																			501
#define DISPID_SL_APPSHORTNAME																502
#define DISPID_SL_BUILD																				503
#define DISPID_SL_CHARSET																			504
#define DISPID_SL_ISRELEASE																		505
#define DISPID_SL_PROGRAMMER																	506
#define DISPID_SL_TESTER																			507
// methods
#define DISPID_SL_ABOUT																				DISPID_ABOUTBOX
#define DISPID_SL_GETIDEALSIZE																30
#define DISPID_SL_HITTEST																			31
#define DISPID_SL_LOADSETTINGSFROMFILE												32
#define DISPID_SL_REFRESH																			DISPID_REFRESH
#define DISPID_SL_SAVESETTINGSTOFILE													33
#define DISPID_SL_FINISHOLEDRAGDROP														34


// _ISysLinkEvents
// methods
#define DISPID_SLE_CARETCHANGED																1
#define DISPID_SLE_CLICK																			2
#define DISPID_SLE_CONTEXTMENU																3
#define DISPID_SLE_CUSTOMDRAW																	4
#define DISPID_SLE_CUSTOMIZETEXTDRAWING												5
#define DISPID_SLE_DBLCLICK																		6
#define DISPID_SLE_DESTROYEDCONTROLWINDOW											7
#define DISPID_SLE_KEYDOWN																		8
#define DISPID_SLE_KEYPRESS																		9
#define DISPID_SLE_KEYUP																			10
#define DISPID_SLE_LINKGETINFOTIPTEXT													11
#define DISPID_SLE_LINKMOUSEENTER															12
#define DISPID_SLE_LINKMOUSELEAVE															13
#define DISPID_SLE_MCLICK																			14
#define DISPID_SLE_MDBLCLICK																	15
#define DISPID_SLE_MOUSEDOWN																	16
#define DISPID_SLE_MOUSEENTER																	17
#define DISPID_SLE_MOUSEHOVER																	18
#define DISPID_SLE_MOUSELEAVE																	19
#define DISPID_SLE_MOUSEMOVE																	20
#define DISPID_SLE_MOUSEUP																		21
#define DISPID_SLE_OLEDRAGDROP																22
#define DISPID_SLE_OLEDRAGENTER																23
#define DISPID_SLE_OLEDRAGLEAVE																24
#define DISPID_SLE_OLEDRAGMOUSEMOVE														25
#define DISPID_SLE_OPENLINK																		DISPID_VALUE
#define DISPID_SLE_RCLICK																			26
#define DISPID_SLE_RDBLCLICK																	27
#define DISPID_SLE_RECREATEDCONTROLWINDOW											28
#define DISPID_SLE_RESIZEDCONTROLWINDOW												29
#define DISPID_SLE_TEXTCHANGED																30
#define DISPID_SLE_XCLICK																			31
#define DISPID_SLE_XDBLCLICK																	32


// ILink
// properties
#define DISPID_L_CARET																				1
#define DISPID_L_DRAWASNORMALTEXT															2
#define DISPID_L_ENABLED																			3
#define DISPID_L_HOT																					4
#define DISPID_L_ID																						5
#define DISPID_L_INDEX																				6
#define DISPID_L_TEXT																					7
#define DISPID_L_URL																					DISPID_VALUE
#define DISPID_L_VISITED																			8
// methods


// ILinks
// properties
#define DISPID_LS_CASESENSITIVEFILTERS												1
#define DISPID_LS_COMPARISONFUNCTION													2
#define DISPID_LS_FILTER																			3
#define DISPID_LS_FILTERTYPE																	4
#define DISPID_LS_ITEM																				DISPID_VALUE
#define DISPID_LS__NEWENUM																		DISPID_NEWENUM
// methods
#define DISPID_LS_CONTAINS																		5
#define DISPID_LS_COUNT																				6


// IWindowedLabel
// properties
#define DISPID_WLBL_APPEARANCE																1
#define DISPID_WLBL_AUTOSIZE																	2
#define DISPID_WLBL_BACKCOLOR																	3
#define DISPID_WLBL_BACKSTYLE																	4
#define DISPID_WLBL_BORDERSTYLE																5
#define DISPID_WLBL_CLIPLASTLINE															6
#define DISPID_WLBL_DISABLEDEVENTS														7
#define DISPID_WLBL_DONTREDRAW																8
#define DISPID_WLBL_ENABLED																		DISPID_ENABLED
#define DISPID_WLBL_FONT																			9
#define DISPID_WLBL_FORECOLOR																	10
#define DISPID_WLBL_HALIGNMENT																11
#define DISPID_WLBL_HOVERTIME																	12
#define DISPID_WLBL_HWND																			13
#define DISPID_WLBL_MOUSEICON																	14
#define DISPID_WLBL_MOUSEPOINTER															15
#define DISPID_WLBL_OWNERDRAWN																16
#define DISPID_WLBL_REGISTERFOROLEDRAGDROP										17
#define DISPID_WLBL_RIGHTTOLEFT																18
#define DISPID_WLBL_SUPPORTOLEDRAGIMAGES											19
#define DISPID_WLBL_TEXT																			DISPID_VALUE
#define DISPID_WLBL_TEXTTRUNCATIONSTYLE												20
#define DISPID_WLBL_USEMNEMONIC																21
#define DISPID_WLBL_USESYSTEMFONT															22
#define DISPID_WLBL_VERSION																		23
// fingerprint
#define DISPID_WLBL_APPID																			500
#define DISPID_WLBL_APPNAME																		501
#define DISPID_WLBL_APPSHORTNAME															502
#define DISPID_WLBL_BUILD																			503
#define DISPID_WLBL_CHARSET																		504
#define DISPID_WLBL_ISRELEASE																	505
#define DISPID_WLBL_PROGRAMMER																506
#define DISPID_WLBL_TESTER																		507
// methods
#define DISPID_WLBL_ABOUT																			DISPID_ABOUTBOX
#define DISPID_WLBL_LOADSETTINGSFROMFILE											24
#define DISPID_WLBL_REFRESH																		DISPID_REFRESH
#define DISPID_WLBL_SAVESETTINGSTOFILE												25
#define DISPID_WLBL_FINISHOLEDRAGDROP													26


// _IWindowedLabelEvents
// methods
#define DISPID_WLBLE_CLICK																		1
#define DISPID_WLBLE_CONTEXTMENU															2
#define DISPID_WLBLE_DBLCLICK																	3
#define DISPID_WLBLE_DESTROYEDCONTROLWINDOW										4
#define DISPID_WLBLE_MCLICK																		5
#define DISPID_WLBLE_MDBLCLICK																6
#define DISPID_WLBLE_MOUSEDOWN																7
#define DISPID_WLBLE_MOUSEENTER																8
#define DISPID_WLBLE_MOUSEHOVER																9
#define DISPID_WLBLE_MOUSELEAVE																10
#define DISPID_WLBLE_MOUSEMOVE																11
#define DISPID_WLBLE_MOUSEUP																	12
#define DISPID_WLBLE_OLEDRAGDROP															13
#define DISPID_WLBLE_OLEDRAGENTER															14
#define DISPID_WLBLE_OLEDRAGLEAVE															15
#define DISPID_WLBLE_OLEDRAGMOUSEMOVE													16
#define DISPID_WLBLE_OWNERDRAW																17
#define DISPID_WLBLE_RCLICK																		18
#define DISPID_WLBLE_RDBLCLICK																19
#define DISPID_WLBLE_RECREATEDCONTROLWINDOW										20
#define DISPID_WLBLE_RESIZEDCONTROLWINDOW											21
#define DISPID_WLBLE_TEXTCHANGED															DISPID_VALUE
#define DISPID_WLBLE_XCLICK																		22
#define DISPID_WLBLE_XDBLCLICK																23


// IWindowlessLabel
// properties
#define DISPID_WLLBL_APPEARANCE																1
#define DISPID_WLLBL_AUTOSIZE																	2
#define DISPID_WLLBL_BACKCOLOR																3
#define DISPID_WLLBL_BACKSTYLE																4
#define DISPID_WLLBL_BORDERSTYLE															5
#define DISPID_WLLBL_CLIPLASTLINE															6
#define DISPID_WLLBL_DISABLEDEVENTS														7
#define DISPID_WLLBL_DONTREDRAW																8
#define DISPID_WLLBL_ENABLED																	DISPID_ENABLED
#define DISPID_WLLBL_FONT																			9
#define DISPID_WLLBL_FORECOLOR																10
#define DISPID_WLLBL_HALIGNMENT																11
#define DISPID_WLLBL_HOVERTIME																12
#define DISPID_WLLBL_MOUSEICON																13
#define DISPID_WLLBL_MOUSEPOINTER															14
#define DISPID_WLLBL_REGISTERFOROLEDRAGDROP										15
#define DISPID_WLLBL_RIGHTTOLEFT															16
#define DISPID_WLLBL_SUPPORTOLEDRAGIMAGES											17
#define DISPID_WLLBL_TEXT																			DISPID_VALUE
#define DISPID_WLLBL_TEXTTRUNCATIONSTYLE											18
#define DISPID_WLLBL_USEMNEMONIC															19
#define DISPID_WLLBL_USESYSTEMFONT														20
#define DISPID_WLLBL_VALIGNMENT																21
#define DISPID_WLLBL_VERSION																	22
#define DISPID_WLLBL_WORDWRAPPING															23
// fingerprint
#define DISPID_WLLBL_APPID																		500
#define DISPID_WLLBL_APPNAME																	501
#define DISPID_WLLBL_APPSHORTNAME															502
#define DISPID_WLLBL_BUILD																		503
#define DISPID_WLLBL_CHARSET																	504
#define DISPID_WLLBL_ISRELEASE																505
#define DISPID_WLLBL_PROGRAMMER																506
#define DISPID_WLLBL_TESTER																		507
// methods
#define DISPID_WLLBL_ABOUT																		DISPID_ABOUTBOX
#define DISPID_WLLBL_LOADSETTINGSFROMFILE											24
#define DISPID_WLLBL_REFRESH																	DISPID_REFRESH
#define DISPID_WLLBL_SAVESETTINGSTOFILE												25
#define DISPID_WLLBL_FINISHOLEDRAGDROP												26


// _IWindowlessLabelEvents
// methods
#define DISPID_WLLBLE_CLICK																		1
#define DISPID_WLLBLE_CONTEXTMENU															2
#define DISPID_WLLBLE_DBLCLICK																3
#define DISPID_WLLBLE_MCLICK																	4
#define DISPID_WLLBLE_MDBLCLICK																5
#define DISPID_WLLBLE_MOUSEDOWN																6
#define DISPID_WLLBLE_MOUSEENTER															7
#define DISPID_WLLBLE_MOUSEHOVER															8
#define DISPID_WLLBLE_MOUSELEAVE															9
#define DISPID_WLLBLE_MOUSEMOVE																10
#define DISPID_WLLBLE_MOUSEUP																	11
#define DISPID_WLLBLE_OLEDRAGDROP															12
#define DISPID_WLLBLE_OLEDRAGENTER														13
#define DISPID_WLLBLE_OLEDRAGLEAVE														14
#define DISPID_WLLBLE_OLEDRAGMOUSEMOVE												15
#define DISPID_WLLBLE_RCLICK																	16
#define DISPID_WLLBLE_RDBLCLICK																17
#define DISPID_WLLBLE_RESIZEDCONTROLWINDOW										18
#define DISPID_WLLBLE_TEXTCHANGED															DISPID_VALUE
#define DISPID_WLLBLE_XCLICK																	19
#define DISPID_WLLBLE_XDBLCLICK																20


// IOLEDataObject
// properties
// methods
#define DISPID_ODO_CLEAR																			1
#define DISPID_ODO_GETCANONICALFORMAT													2
#define DISPID_ODO_GETDATA																		3
#define DISPID_ODO_GETFORMAT																	4
#define DISPID_ODO_SETDATA																		5
#define DISPID_ODO_GETDROPDESCRIPTION													6
#define DISPID_ODO_SETDROPDESCRIPTION													7

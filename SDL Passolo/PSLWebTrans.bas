''This macro is a online translations macro program for strings
''in Passolo translation list.
''It has the following features:
''- Use the online translation engine automatically translate strings
''  in the Passolo translation list
''- Integrated some of the well-known online translation engines, and
''  you can customize other online translation engines
''- You can choose the string type, skiping some of string, and processing
''  the strings before and after translation
''- Integrated shortcuts, terminators, Accelerator check macro, and you can
''  check and correct errors in translations after the strings has be translated
''Idea and implemented by wanfu 2010.05.12 (Last modified on 2017.06.16)

Public Type STRING_INFO
	PreSpace			As String	'字符串前置空格
	EndSpace			As String	'字符串后置空格
	Spaces				As String	'字符串快捷键前空格
	AccKey				As String	'字符串快捷键
	AccKeyIFR			As String	'字符串快捷键标志符
	AccKeyKey			As String	'字符串快捷键字符
	EndString			As String	'字符串终止符
	Shortcut			As String	'字符串加速器
	PreString			As String	'字符串快捷键前的字符（不含快捷键前的空格和终止符）
	ExpString			As String	'字符串快捷键至加速器前的字符
	AccKeyPos			As Integer	'字符串快捷键位置
	AccKeyNum			As Integer	'字符串快捷键数
	Length				As Integer	'字符串长度
	LineNum				As Integer	'字符串行数
End Type

Public Type CHECK_STRING_VALUE
	AscRange			As String
	Range				As String
End Type

'编辑工具定义
Public Type TOOLS_PROPERTIE
	sName				As String	'工具名称
	FilePath			As String	'工具文件路径(含文件名)
	Argument			As String	'运行参数
End Type

Public Type UI_FILE
	FilePath			As String	'语言文件完全路径
	AppName				As String	'程序名称
	Version				As String	'程序版本
	LangName			As String	'语言名称
	LangID				As String	'适用语言ID
	Encoding			As String	'字符编码
End Type

Public Type INIFILE_DATA
	Title				As String	'主题
	Item()				As String	'项目
	Value()				As String	'字串值
End Type

Public Type CODEPAGE_DATA
	sName				As String	'代码名称
	CharSet				As String	'字符编码
End Type

'光标坐标定义
Private Type POINTAPI
	x As Long
	y As Long
End Type

'用于文本框查找定位常数
Private Enum SendMsgValue
	'用于文本框查找定位常数
	EM_GETSEL = &HB0				'0,变量			获取光标位置（以字符数表示）
	EM_SETSEL = &HB1				'起点,终点		设置编辑控件中文本选定内容范围（或设置光标位置）起点和终点均为字符值
									'				当指定的起点等于0和终点等于-1时，文本全部被选中，此法常用在清空编辑控件
									'				当指定的起点等于-2和终点等于-1时，全文均不选，光标移至文本未端
	EM_GETLINECOUNT = &HBA			'0,0			获取编辑控件的总行数
	EM_LINEINDEX = &HBB				'行号,0			获取指定行(或:-1,0 表示光标所在行)首字符在文本中的位置（以字符数表示）
	EM_LINELENGTH = &HC1			'偏移值,0		获取指定位置所在行(或:-1,0 表示光标所在行）的文本长度（以字符数表示）
	EM_LINEFROMCHAR = &HC9			'偏移值,0		获取指定位置(或:-1,0 表示光标位置)所在的行号
	EM_GETLINE = &HC4				'行号,ByVal变量	获取编辑控件某一行的内容，变量须预先赋空格
	EM_SCROLLCARET = &HB7			'0,0 			把可见范围移至光标处
	EM_UNDO = &HC7					'0,0 			撤消前一次编辑操作，当重复发送本消息，控件将在撤消和恢复中来回切换
	EM_REPLACESEL = &HC2			'1(0),字符串	用指定字符串替换编辑控件中的当前选定内容
									'				如果第三个参数wParam为1，则本次操作允许撤消，0禁止撤消。字符串可用传值方式，也可用传址方式
									'				（例：SendMessage Text1.hwnd, EM_REPLACESEL, 0, Text2.Text '这是传值方式）
	EM_GETMODIFY = &HB8				'0,0			判断编辑控件的内容是否已发生变化，返回TRUE(1)则控件文本已被修改，返回FALSE(0)则未变
	EN_CHANGE = &H300 				'				编辑控件的内容发生改变。与EN_UPDATE不同，该消息是在编辑框显示的正文被刷新后才发出的
	EN_UPDATE = &H400				'				控件准备显示改变了的正文时发送该消息。它与EN_CHANGE通知消息相似，只是它发生于更新文本显示出来之前
	EM_GETLIMITTEXT = &H425			'0,0			获取一个编辑控件中文本的最大长度
	EM_LIMITTEXT = &HC5				'最大值,0		设置编辑控件中的最大文本长度
	EM_GETFIRSTVISIBLEINE = &HCE	'0,0			获得文本控件中处于可见位置的最顶部的文本所在的行号
	EM_GETHANDLE = &HBD				'0,0			取得文本缓冲区

	'用于对话框字体
	WM_SETFONT = &H30				'字体句柄,True	绘制文本时程序发送此消息获取控件要用的字体
	WM_GETFONT = &H31				'0,0 			获取当前控件绘制文本的字体句柄
	WM_FONTCHANGE = &H1D 			'0,0			当系统的字体资源库变化时发送此消息给所有顶级窗口
	WM_SETREDRAW = &HB 				'Boolean,0		设置窗口是否能重画，False 禁止重画，True 允许重画
	WM_SETFOCUS = &H7				'控件句柄,0,0	设置焦点
End Enum

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
	ByRef Destination As Any, _
	ByRef Source As Any, _
	ByVal Length As Long)

'用于文本框查找定位函数
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByRef lParam As Any) As Long
Private Declare Function SendMessageLNG Lib "user32.dll" Alias "SendMessageA" ( _
	ByVal hwnd As Long, _
	ByVal wMsg As Long, _
	ByVal wParam As Long, _
	ByVal lParam As Long) As Long
Public Declare Function GetFocus Lib "USER32.dll" () As Long	'用于返回焦点控件的句柄

'用于返回控件ID的句柄
Private Declare Function GetDlgItem Lib "user32.dll" ( _
	ByVal hDlg As Long, _
	ByVal nIDDlgItem As Long) As Long

'ChooseFont 类型的 flags 参数定义
Private Enum CF_VALUE
	CF_APPLY = &H200
	CF_ANSIONLY = &H400
	CF_TTONLY = &H40000
	CF_ENABLEHOOK = &H8
	CF_ENABLETEMPLATE = &H10
	CF_ENABLETEMPLATEHANDLE = &H20
	CF_FIXEDPITCHONLY = &H4000
	CF_NOVECTORFONTS = &H800
	CF_NOOEMFONTS = CF_NOVECTORFONTS
	CF_NOFACESEL = &H80000
	CF_NOSCRIPTSEL = CF_NOFACESEL
	CF_NOSTYLESEL = &H100000
	CF_NOSIZESEL = &H200000
	CF_NOSIMULATIONS = &H1000
	CF_NOVERTFONTS = &H1000000
	CF_SCALABLEONLY = &H20000
	CF_SCRIPTSONLY = CF_ANSIONLY
	CF_SELECTSCRIPT = &H400000
	CF_SHOWHELP = &H4
	CF_USESTYLE = &H80
	CF_WYSIWYG = &H8000			'must also have CF_SCREENFONTS CF_PRINTERFONTS
	CF_FORCEFONTEXIST = &H10000
	CF_INACTIVEFONTS = &H2000000
	CF_INITTOLOGFONTSTRUCT = &H40&
	CF_SCREENFONTS = &H1		'显示屏幕字体
	CF_PRINTERFONTS = &H2		'显示打印机字体
	CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)	'两者都显示
	CF_EFFECTS = &H100&			'添加字体效果
	CF_LIMITSIZE = &H2000&		'设置字体大小限制
End Enum

'字体类型
Private Type LOG_FONT
	lfHeight As Long			'字体大小
	lfWidth As Long				'字体宽度
	lfEscapement As Long		'字体显示角度
	lfOrientation As Long		'字体角度
	lfWeight As Long			'是否粗体
	lfItalic As Byte			'是否斜体
	lfUnderline As Byte			'是否下划线
	lfStrikeOut As Byte			'是否删除线
	lfCharSet As Byte			'字符集
	lfOutPrecision As Byte		'输出精度
	lfClipPrecision As Byte		'裁减精度
	lfQuality As Byte			'逻辑字体与输出设备实际字体之间的精度
	lfPitchAndFamily As Byte	'字体间距和字体集
	'lfFaceName As String * LF_FACESIZE	'字体名称(不能这样定义，创建字体时会出错)
	lfFaceName(31) As Byte		'字体名称
	lfColor As Long				'字体颜色
End Type

'字体对话框类型
Private Type CHOOSE_FONT
	lStructSize As Long			' size of CHOOSEFONT structure in byte
	hwndOwner As Long			' caller's window handle
	hDC As Long					' printer DC/IC or NULL
	lpLogFont As Long			' LogFont 结构地址
	iPointSize As Long			' 10 * size in points of selected font
	flags As CF_VALUE			' enum type flags
	rgbColors As Long			' returned text color
	lCustData As Long			' data passed to hook fn
	lpfnHook As Long			' ptr. to hook function
	lpTemplateName As String	' custom template name
	hInstance As Long			' instance handle of.EXE that contains cust. dlg. template
	lpszStyle As String			' return the style field here must be LF_FACESIZE or bigger
	nFontType As Integer		' same value reported to the EnumFonts call back with the extra FONTTYPE_ bits added
	MISSING_ALIGNMENT As Integer
	nSizeMin As Long			' minimum pt size allowed
	nSizeMax As Long			' max pt size allowed if CF_LIMITSIZE is used
End Type

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSE_FONT) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOG_FONT) As Long
Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" ( _
	ByVal hObject As Long, _
	ByVal nCount As Long, _
	ByVal lpObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" ( _
	ByVal hDC As Long, _
	ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

'RedrawWindow 函数的 fuRedraw 参数定义
Private Enum RDW
	RDW_INVALIDATE = &H1		'禁用（屏蔽）重画区域
	RDW_INTERNALPAINT = &H2		'即使窗口并非无效，也向其投递一条WM_PAINT消息
	RDW_ERASE = &H4				'重画前，先清除重画区域的背景。也必须指定RDW_INVALIDATE
	RDW_VALIDATE = &H8			'检验重画区域
	RDW_NOINTERNALPAINT = &H10	'禁止内部生成或由这个函数生成的任何待决WM_PAINT消息。针对无效区域，仍会生成WM_PAINT消息
	RDW_NOERASE = &H20			'禁止删除重画区域的背景
	RDW_NOCHILDREN = &H40		'重画操作排除子窗口（前提是它们存在于重画区域）
	RDW_ALLCHILDREN = &H80		'重画操作包括子窗口（前提是它们存在于重画区域）
	RDW_UPDATENOW = &H100		'立即更新指定的重画区域
	RDW_ERASENOW = &H200		'立即删除指定的重画区域
	RDW_FRAME = &H400			'如非客户区包含在重画区域中，则对非客户区进行更新。也必须指定RDW_INVALIDATE
	RDW_NOFRAME = &H800			'禁止非客户区域重画（如果它是重画区域的一部分）。也必须指定RDW_VALIDATE
End Enum

'重画对话框函数
Private Declare Function RedrawWindow Lib "user32.dll" ( _
	ByVal hwnd As Long, _
	ByVal lprcUpdate As Long, _
	ByVal hrgnUpdate As Long, _
	ByVal fuRedraw As Long) As Long

'代码页常数
Private Enum KnownCodePage
	CP_UNICODELITTLE = 1200
	CP_UNICODEBIG = 1201
	CP_UTF7 = 65000
	CP_UTF8 = 65001
End Enum

'转换宽字符为多字节
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
	ByVal CodePage As Long, _
	ByVal dwFlags As Long, _
	ByVal lpWideCharStr As Long, _
	ByVal cchWideChar As Long, _
	ByRef lpMultiByteStr As Any, _
	ByVal cchMultiByte As Long, _
	ByVal lpDefaultChar As Long, _
	ByVal lpUsedDefaultChar As Long) As Long
'获取本机默认代码页
Private Declare Function GetACP Lib "kernel32.dll" () As Long

Public trn As PslTransList,TransString As PslTransString,OSLanguage As String
Public PslLangDataList() As String,LangPair As String
Public UIFileList() As UI_FILE,UIDataList() As INIFILE_DATA,LangFile As String
Public LFList() As LOG_FONT,LFListBak() As LOG_FONT
Public CheckHexStr() As CHECK_STRING_VALUE,CheckSkipStr() As String

Public StringSrc As STRING_INFO,StringTrn As STRING_INFO,MoveAcckey As String
Public ModifiedCount As Integer,AddedCount As Integer,MovedCount As Integer,DeledCount As Integer,ReplacedCount As Integer
Public AllCont As Long,AccKey As Long,EndChar As Long,Acceler As Long

Public DefaultCheckList() As String,DefaultProjectList() As String,AppRepStr As String,PreRepStr As String
Public cWriteLoc As String,cSelected() As String,cUpdateSet() As String,CheckVersion As String
Public CheckList() As String,CheckListBak() As String,CheckDataList() As String,CheckDataListBak() As String
Public ProjectList() As String,ProjectListBak() As String,ProjectDataList() As String,ProjectDataListBak() As String

Public DefaultEngineList() As String,WaitTimes As Long
Public tWriteLoc As String,tSelected() As String,tSelectedBak() As String,tUpdateSet() As String,tUpdateSetBak() As String
Public EngineList() As String,EngineListBak() As String,EngineDataList() As String,EngineDataListBak() As String
Public DelLngNameList() As String,DelSrcLngList() As String,DelTranLngList() As String

Public AllStrList() As String,UseStrList() As String
Public Tools() As TOOLS_PROPERTIE,ToolsBak() As TOOLS_PROPERTIE
Public FileDataList() As String,CodeList() As CODEPAGE_DATA,RegExp As Object

Public Const Version = "2017.04.22"
Public Const Build = "170616"
Public Const ToUpdateEngineVersion = "2017.04.22"
Public Const ToUpdateCheckVersion = "2016.09.13"
Public Const EngineRegKey = "HKCU\Software\VB and VBA Program Settings\WebTranslate\"
Public Const EngineFilePath = MacroDir & "\Data\PSLWebTrans.dat"
Public Const CheckRegKey = "HKCU\Software\VB and VBA Program Settings\AccessKey\"
Public Const CheckFilePath = MacroDir & "\Data\PSLCheckAccessKeys.dat"
Public Const JoinStr = vbFormFeed  'vbBack
Public Const SubJoinStr = vbVerticalTab  'Chr$(1)
Public Const LngJoinStr = "|"
Public Const SubLngJoinStr = Chr$(1)
Public Const NullValue = "Null"
Public Const ConvertStrHexRange = "\x30-\x39,\x41-\x46;\x30-\x39,\x41-\x46;\x30-\x37"
			'0=用于十六进制转义符,1=用于 Unicode 转义符,2=用于八进制转义符
Public Const CheckStrHexRange = "\x00-\x07,\x0E-\x1F;\x00-\x40,\x5B-\x60,\x7B-\xBF;" & _
								"\x00-\x60,\x7B-\xBF;\x00-\x40,\x5B-\xBF;\x41-\x5A,\x61-\x7A;" & _
								"\x30-\x39,\x41-\x5A,\x61-\x7A"
			'0=控制字符,1=全为数字和符号,2=全为大写英文,3=全为小写英文,4=大小写混合英文,5=快捷键字符范围

Public Const DefaultObject = "Microsoft.XMLHTTP;Msxml2.XMLHTTP"
Public Const AppName = "PSLWebTrans"
Public Const updateMainFile = "PSLWebTrans.bas"
Public Const updateINIFile = "PSLMacrosUpdates.rar"
Public Const updateINIMainUrl = "http://jp.wanfutrade.com/download/PSLMacrosUpdates.rar"
Public Const updateINIMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLMacrosUpdates.rar"
Public Const updateMainUrl = "http://jp.wanfutrade.com/download/PSLWebTrans.rar"
Public Const updateMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLWebTrans.rar"
Public Const MacroLoc = MacroDir
Public Const DefaultWaitTimes = 2    '2秒


'翻译引擎默认设置
Function EngineSettings(ByVal DataName As String) As String
	Dim StesArray(19) As String
	Select Case DataName
	Case DefaultEngineList(0)
		StesArray(0) = DefaultObject
		StesArray(1) = "fefed727-bbc1-4421-828d-fc828b24d59b"
		StesArray(2) = "http://api.microsofttranslator.com/V2/Http.svc/Translate?"
		StesArray(3) = "{Url}&appId={appId}&text={text}&from={from}&to={to}"
		StesArray(4) = "GET"
		StesArray(5) = "False"
		StesArray(6) = ""
		StesArray(7) = ""
		StesArray(8) = ""
		StesArray(9) = "Content-Type,application/xml; charset=utf-8"
		StesArray(10) = "responseBody"
		StesArray(11) = "Serialization/"">"
		StesArray(12) = "</string>"
		StesArray(13) = "Serialization/"">"
		StesArray(14) = "</string>"
		StesArray(15) = "Serialization/"">"
		StesArray(16) = "</string>"
		StesArray(17) = "string"
		StesArray(18) = "string"
		StesArray(19) = "1"
	Case DefaultEngineList(1)
		StesArray(0) = DefaultObject
		StesArray(1) = ""
		StesArray(2) = "https://translate.google.cn/?"
		StesArray(3) = "{Url}&text={text}&langpair={from}|{to}"
		StesArray(4) = "GET"
		StesArray(5) = "False"
		StesArray(6) = ""
		StesArray(7) = ""
		StesArray(8) = ""
		StesArray(9) = "Content-Type,text/html; charset=utf-8"
		StesArray(10) = "responseBody"
		StesArray(11) = "onmouseout=""this.style.backgroundColor='#fff'"">"
		StesArray(12) = "</span>"
		StesArray(13) = "onmouseout=""this.style.backgroundColor='#fff'"">"
		StesArray(14) = "</span>"
		StesArray(15) = "onmouseout=""this.style.backgroundColor='#fff'"">"
		StesArray(16) = "</span>"
		StesArray(17) = ""
		StesArray(18) = ""
		StesArray(19) = "1"
	Case DefaultEngineList(2)
		StesArray(0) = DefaultObject
		StesArray(1) = ""
		StesArray(2) = "http://fanyi.yahoo.com.cn/translate_txt?"
		StesArray(3) = "{Url}&ei=UTF-8&fr=&lp={from}_{to}&trtext={Text}"
		StesArray(4) = "POST"
		StesArray(5) = "False"
		StesArray(6) = ""
		StesArray(7) = ""
		StesArray(8) = ""
		StesArray(9) = "Content-Type,text/html; charset=utf-8"
		StesArray(10) = "responseBody"
		StesArray(11) = "<div id=""pd"" class=""pd"">"
		StesArray(12) = "</div>"
		StesArray(13) = "<div id=""pd"" class=""pd"">"
		StesArray(14) = "</div>"
		StesArray(15) = "<div id=""pd"" class=""pd"">"
		StesArray(16) = "</div>"
		StesArray(17) = ""
		StesArray(18) = ""
		StesArray(19) = "0"
	End Select
	EngineSettings = Join(StesArray,SubJoinStr)
End Function


'字串处理默认设置
Function CheckSettings(ByVal DataName As String,ByVal DataType As Long) As String
	Dim i As Long,j As Long,CheckName As String,File As String
	Dim TempList() As String,DataList() As INIFILE_DATA

	If DataType = 0 Then
		ReDim TempList(17) As String
	Else
		ReDim TempList(20) As String
	End If

	If DataName <> "" Then
		If DataType = 0 Then
			Select Case DataName
			Case DefaultCheckList(0)
				CheckName = "en2zh"
			Case DefaultCheckList(1)
				CheckName = "zh2en"
			End Select
		ElseIf DataType = 1 Then
			Select Case DataName
			Case DefaultProjectList(0)
				CheckName = "CheckOnly"
			Case DefaultProjectList(1)
				CheckName = "CheckAndCorrect"
			Case DefaultProjectList(2)
				CheckName = "DelAccessKey"
			Case DefaultProjectList(3)
				CheckName = "DelAccelerator"
			Case DefaultProjectList(4)
				CheckName = "DelAccessKeyAndAccelerator"
			End Select
		End If
	End If
	If CheckName = "" Then GoTo ExitFunction

	File = MacroDir & "\Data\PSLCheckAccessKeys.ini"
	If getINIFile(DataList,File,"unicodeFFFE",1) = False Then GoTo NotReadFile

	On Error GoTo ErrMassage
	For i = 0 To UBound(DataList)
		With DataList(i)
			If .Title = "Option" Then
				For j = 0 To UBound(.Item)
					If .Item(j) = "Version" Then
						If StrComp(ToUpdateCheckVersion,.Value(j)) = 1 Then
							CheckName = ""
							Err.Raise(1,"NotVersion",File & JoinStr & .Value(j) & _
										JoinStr & ToUpdateCheckVersion)
							Exit For
						End If
					End If
				Next j
			ElseIf DataType = 0 And .Title = CheckName Then
				For j = 0 To UBound(.Item)
					Select Case .Item(j)
					Case "ExcludeChar"
						TempList(0) = .Value(j)
					Case "LineSplitChar"
						TempList(1) = .Value(j)
					Case "CheckBracket"
						TempList(2) = .Value(j)
					Case "KeepCharPair"
						TempList(3) = .Value(j)
					Case "ShowAsiaKey"
						TempList(4) = .Value(j)
					Case "CheckEndChar"
						TempList(5) = .Value(j)
					Case "NoTrnEndChar"
						TempList(6) = .Value(j)
					Case "AutoTrnEndChar"
						TempList(7) = .Value(j)
					Case "CheckShortChar"
						TempList(8) = .Value(j)
					Case "CheckShortKey"
						TempList(9) = .Value(j)
					Case "KeepShortKey"
						TempList(10) = .Value(j)
					Case "PreRepString"
						TempList(11) = .Value(j)
					Case "AutoRepString"
						TempList(12) = .Value(j)
					Case "AccessKeyChar"
						TempList(13) = .Value(j)
					Case "AddAccessKeyWithFirstChar"
						TempList(14) = .Value(j)
					Case "LineSplitMode"
						TempList(15) = .Value(j)
					Case "AppInsertSplitChar"
						TempList(16) = .Value(j)
					Case "ReplaceSplitChar"
						TempList(17) = .Value(j)
					End Select
				Next j
				CheckSettings = Join(TempList,SubJoinStr)
			ElseIf DataType = 1 And .Title = "Projects" Then
				For j = 0 To UBound(.Item)
					If .Item(j) = CheckName Then TempList = ReSplit(.Value(j),LngJoinStr)
				Next j
				CheckSettings = Join(TempList,LngJoinStr)
			End If
		End With
	Next i

	If CheckSettings = "" And CheckName <> "" Then
		If DataType = 0 Then
			Temp = "NotSection"
		ElseIf DataType = 1 Then
			Temp = "NotValue"
		End If
		Err.Raise(1,Temp,File & JoinStr & CheckName)
	End If
	Exit Function

	NotReadFile:
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & File

	ErrMassage:
	Call sysErrorMassage(Err,1)

	ExitFunction:
	If CheckSettings = "" Then
		If DataType = 0 Then
			CheckSettings = Join(TempList,SubJoinStr)
		ElseIf DataType = 1 Then
			CheckSettings = Join(TempList,LngJoinStr)
		End If
	End If
End Function


' 主程序
Sub Main
	Dim i As Long,j As Long,n As Long,srcString As String,trnString As String
	Dim xmlHttp As Object,Obj As Object,TrnList As PslTransList
	Dim CheckID As Long,EngineID As Long,StringCount As Long
	Dim TranedCount As Integer,SkipedCount As Integer,NotChangeCount As Integer,NotTranCount As Integer
	Dim TransListOpen As Boolean
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer
	Dim TransID As Long,TransIDBak As Long,MsgList() As String,Stemp As Boolean
	Dim trnsList As PslTransDisplay
	Dim trnTitleList() As String,SrcLangList() As String,TempList() As String,TempArray() As String
	Dim mCheckSrc As Long,ProjectIDSrc As Long,mCheckTrn As Long,ProjectIDTrn As Long
	Dim ShowOriginalTran As Long,ApplyCheckResult As Long,Temp As String

	'检测系统语言
	On Error Resume Next
	Set Obj = CreateObject("WScript.Shell")
	If Obj Is Nothing Then
		MsgBox(Err.Description & " - " & "WScript.Shell",vbInformation)
		Exit Sub
	End If
	Temp = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\Default"
	OSLanguage = Obj.RegRead(Temp)
	If OSLanguage = "" Then
		Temp = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\InstallLanguage"
		OSLanguage = Obj.RegRead(Temp)
		If Err.Source = "WshShell.RegRead" Then
			MsgBox(Err.Description,vbInformation)
			Exit Sub
		End If
	End If
	Set Obj = Nothing

	'检测 Adodb.Stream 是否存在
	Set Obj = CreateObject("Adodb.Stream")
	If Obj Is Nothing Then
		MsgBox(Err.Description & " - " & "Adodb.Stream",vbInformation)
		Exit Sub
	End If
	Set Obj = Nothing

	'检测 VBScript.RegExp 是否存在
	Set RegExp = CreateObject("VBScript.RegExp")
	If RegExp Is Nothing Then
		MsgBox(Err.Description & " - " & "VBScript.RegExp",vbInformation)
		Exit Sub
	End If
	On Error GoTo SysErrorMsg

	'初始化数组
	ReDim Tools(3) As TOOLS_PROPERTIE,FileDataList(0) As String,trnTitleList(0) As String
	ReDim UIFileList(0) As UI_FILE,UIDataList(0) As INIFILE_DATA,PslLangDataList(0) As String
	ReDim tSelected(27) As String,tUpdateSet(5) As String,LFList(0) As LOG_FONT
	ReDim DefaultEngineList(2) As String,EngineList(0) As String,EngineDataList(0) As String
	ReDim cSelected(26) As String,cUpdateSet(5) As String
	ReDim DefaultCheckList(1) As String,CheckList(0) As String,CheckDataList(0) As String
	ReDim DefaultProjectList(4) As String,ProjectList(0) As String,ProjectDataList(0) As String
	UIFileList(0).LangID = "0"

	'转换 HEX 字符值为 Long 数组和正则表达式模板，用于转换文本的转义符
	MsgList = ReSplit(ConvertStrHexRange,";")
	ReDim CheckHexStr(UBound(MsgList)) As CHECK_STRING_VALUE
	For i = 0 To UBound(MsgList)
		CheckHexStr(i).Range = Replace$(MsgList(i),",","")
		If i = 0 Then
			CheckHexStr(i).Range = "\\x[" & CheckHexStr(i).Range & "]{2}"
		ElseIf i = 1 Then
			CheckHexStr(i).Range = "\\u[" & CheckHexStr(i).Range & "]{4}"
		ElseIf i = 2 Then
			CheckHexStr(i).Range = "\\[" & CheckHexStr(i).Range & "]+"
		End If
		CheckHexStr(i).AscRange = "[" & Convert(Replace$(MsgList(i),",","")) & "]"
	Next i

	'转换跳过字符值为正则表达式模板，用于文本的正则表达式比较
	CheckSkipStr = ReSplit("[" & Replace$(Replace$(CheckStrHexRange,";","];["),",","]|[") & "]",";")

	'读取翻译引擎设置
	DefaultEngineList(0) = "Microsoft"
	DefaultEngineList(1) = "Google"
	DefaultEngineList(2) = "Yahoo"
	j = GetEngineSet("","")

	'读取界面语言字串
	If GetUI(MacroDir & "\Data\",tSelected(0),OSLanguage,UIDataList,UIFileList,LangFile) = False Then Exit Sub
	If getMsgList(UIDataList,MsgList,"Main",0) = False Then Exit Sub

	'检测 PSL 版本
	If PSL.Version < 600 Then
		MsgBox MsgList(43),vbOkOnly+vbInformation,MsgList(41)
		Exit Sub
	End If

	'获取更新数据并检查新版本
	If CheckUpdate(tUpdateSet) = True Then Exit Sub

	'检测翻译列表是否被选择
	Set trnsList = PSL.ActiveTransDisplay
	If trnsList Is Nothing Then
		Set trn = PSL.ActiveTransList
		If trn Is Nothing Then
			MsgBox MsgList(44),vbOkOnly+vbInformation,MsgList(41)
			Exit Sub
		End If
		trnTitleList(0) = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
	ElseIf trnsList.StringCount > 0 Then
		n = 0
		ReDim trnTitleList(PSL.ActiveProject.TransLists.Count)
		For i = 1 To PSL.ActiveProject.TransLists.Count
			Set trn = PSL.ActiveProject.TransLists(i)
			If trn.IsOpen = True Then
				trnTitleList(n) = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
				n = n + 1
			End If
		Next i
		If n > 0 Then ReDim Preserve trnTitleList(n - 1) Else ReDim trnTitleList(0)
		Set trn = trnsList.String(1,pslDisplay).TransList
	Else
		Set trn = PSL.ActiveTransList
		If trn Is Nothing Then
			MsgBox MsgList(44),vbOkOnly+vbInformation,MsgList(41)
			Exit Sub
		End If
		trnTitleList(0) = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
	End If

	'初始化编辑工具菜单名称
	Tools(0).sName = MsgList(96)
	Tools(1).sName = MsgList(97)
	Tools(2).sName = MsgList(98)
	Tools(3).sName = MsgList(99)
	Tools(1).FilePath = "notepad.exe"

	'获取字符编码列表
	CodeList = getCodePageList(0,49)

	'检查翻译引擎设置
	If j < 4 Or UBound(EngineList) < 2 Then
		For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
			Stemp = False
			For j = LBound(EngineList) To UBound(EngineList)
				If EngineList(j) = DefaultEngineList(i) Then
					Stemp = True
					Exit For
				End If
			Next j
			If Stemp = False Then
				Temp = DefaultEngineList(i) & JoinStr & EngineSettings(DefaultEngineList(i)) & JoinStr & _
						Join(LangCodeList(DefaultEngineList(i),0,-1),SubLngJoinStr)
				CreateArray(DefaultEngineList(i),Temp,EngineList,EngineDataList)
			End If
		Next i
	End If

	'获取来源语言列表
	ReDim SrcLangList(0)
	SrcLangList(0) = PSL.GetLangCode(trn.SourceList.LangID,pslCodeText)
	n = 1
	ReDim Preserve SrcLangList(trn.Project.Languages.Count)
	For i = 1 To trn.Project.Languages.Count
		Set TrnList = trn.Project.TransLists(i)
		If TrnList.ListID <> trn.ListID Then
			SrcLangList(n) = PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
			n = n + 1
		End If
	Next i
	If n > 1 Then ReDim Preserve SrcLangList(n - 1) Else ReDim Preserve trnTitleList(0)

	'读取字串处理设置
	DefaultCheckList(0) = MsgList(70)
	DefaultCheckList(1) = MsgList(71)
	DefaultProjectList(0) = MsgList(89)
	DefaultProjectList(1) = MsgList(90)
	DefaultProjectList(2) = MsgList(91)
	DefaultProjectList(3) = MsgList(92)
	DefaultProjectList(4) = MsgList(93)
	j = GetCheckSet("","")
	If j < 4 Or UBound(CheckList) < 1  Then
		For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
			Stemp = False
			For j = LBound(CheckList) To UBound(CheckList)
				If CheckList(j) = DefaultCheckList(i) Then
					Stemp = True
					Exit For
				End If
			Next j
			If Stemp = False Then
				Temp = CheckSettings(DefaultCheckList(i),0)
				If Trim(Replace(Temp,SubJoinStr,"")) <> "" Then
					CreateArray(DefaultCheckList(i),DefaultCheckList(i) & JoinStr & Temp & JoinStr & _
							Join(LangCodeList(DefaultCheckList(i),1,-1),SubLngJoinStr),CheckList,CheckDataList)
				End If
			End If
		Next i
		If CheckArray(CheckList) = False Then
			If MsgBox(MsgList(78),vbYesNo+vbInformation,MsgList(42)) = vbNo Then Exit Sub
			For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
				Temp = CheckSettings("",0)
				CreateArray(DefaultCheckList(i),DefaultCheckList(i) & JoinStr & Temp & JoinStr & _
						Join(LangCodeList(DefaultCheckList(i),1,-1),SubLngJoinStr),CheckList,CheckDataList)
			Next i
			EngineListBak = EngineList
			EngineDataListBak = EngineDataList
			CheckListBak = CheckList
			CheckDataListBak = CheckDataList
			tUpdateSetBak = tUpdateSet
			Temp = tSelected(0)
			Call Settings(0,0,1)
			If Temp <> tSelected(0) Then
				getMsgList(UIDataList,MsgList,"Main",1)
			End If
		End If
	End If
	If CheckArray(ProjectList) = False Then
		For i = LBound(DefaultProjectList) To UBound(DefaultProjectList)
			Temp = CheckSettings(DefaultProjectList(i),1)
			If Trim(Replace(Temp,LngJoinStr,"")) <> "" Then
				CreateArray(DefaultProjectList(i),DefaultProjectList(i) & JoinStr & Temp,ProjectList,ProjectDataList)
			End If
		Next i
		If CheckArray(ProjectList) = False Then
			If MsgBox(MsgList(87),vbYesNo+vbInformation,MsgList(42)) = vbNo Then Exit Sub
			For i = LBound(DefaultProjectList) To UBound(DefaultProjectList)
				Temp = CheckSettings("",1)
				CreateArray(DefaultProjectList(i),DefaultProjectList(i) & JoinStr & Temp,ProjectList,ProjectDataList)
			Next i
			ProjectListBak = ProjectList
			ProjectDataListBak = ProjectDataList
			Call Projects(0)
		End If
	End If

	'更改字串检查配置名称
	If tSelected(2) = "en2zh" Then
		tSelected(2) = DefaultCheckList(0)
	ElseIf tSelected(2) = "zh2en" Then
		tSelected(2) = DefaultCheckList(1)
	End If

	'备份配置
	tSelectedBak = tSelected()
	tUpdateSetBak = tUpdateSet
	LFListBak = LFList

	'对话框
	StartDlg:
	Temp = tSelected(0)
	Begin Dialog UserDialog 660,518,MsgList(1),.MainDlgFunc ' %GRID:10,7,1,1
		TextBox 0,0,0,21,.SuppValueBox
		Text 20,7,620,14,Replace$(MsgList(0),"%s",Version),.Text1,2
		Text 20,28,620,32,MsgList(2),.Text2
		Text 20,63,620,14,MsgList(3) & Join$(trnTitleList,"; "),.Text3,2

		GroupBox 20,84,300,56,MsgList(4),.Configuration
		DropListBox 40,105,260,21,EngineList(),.EngineList

		GroupBox 340,84,300,56,MsgList(5),.SrcLang
		DropListBox 360,105,260,21,SrcLangList(),.SrcLangList

		GroupBox 20,147,620,63,MsgList(6),.StrTypeSelection
		CheckBox 40,164,140,14,MsgList(7),.AllType
		CheckBox 190,164,140,14,MsgList(8),.Menu
		CheckBox 340,164,140,14,MsgList(9),.Dialog
		CheckBox 490,164,140,14,MsgList(10),.Strings
		CheckBox 40,185,140,14,MsgList(11),.AccTable
		CheckBox 190,185,140,14,MsgList(12),.Versions
		CheckBox 340,185,140,14,MsgList(13),.Other
		CheckBox 490,185,140,14,MsgList(14),.Seleted

		GroupBox 20,217,620,63,MsgList(15),.SkipSelection
		CheckBox 40,234,190,14,MsgList(16),.ForReview
		CheckBox 240,234,190,14,MsgList(17),.Validated
		CheckBox 440,234,190,14,MsgList(18),.NotTran
		CheckBox 40,255,190,14,MsgList(19),.NumAndSymbol
		CheckBox 240,255,190,14,MsgList(20),.AllUCase
		CheckBox 440,255,190,14,MsgList(21),.AllLCase

		GroupBox 20,287,620,168,MsgList(22),.PreProcessing
		Text 40,311,100,14,MsgList(23),.CheckConfigText
		Text 40,350,100,14,MsgList(24),.PreTrnText
		Text 40,406,100,14,MsgList(25),.AppTrnText
		Text 40,332,580,8,String$(200,MsgList(26)),.LineText1,2
		Text 40,388,580,8,String$(200,MsgList(26)),.LineText2,2
		DropListBox 150,308,270,21,CheckList(),.CheckList
		CheckBox 480,311,150,14,MsgList(27),.AutoSelection
		DropListBox 150,347,270,21,ProjectList(),.SrcProjectList
		PushButton 420,347,30,21,MsgList(28),.SrcEditButton
		CheckBox 480,350,150,14,MsgList(29),.CheckSrc
		CheckBox 150,371,240,14,MsgList(30),.PreStrRep
		CheckBox 400,371,230,14,MsgList(31),.SplitTran
		DropListBox 150,403,270,21,ProjectList(),.TrnProjectList
		PushButton 420,403,30,21,MsgList(28),.TrnEditButton
		CheckBox 480,406,150,14,MsgList(32),.CheckTrn
		CheckBox 150,427,240,14,MsgList(33),.AppStrRep
		CheckBox 400,427,230,14,MsgList(34),.TranComment

		CheckBox 30,462,370,14,MsgList(35),.KeepSet
		CheckBox 400,462,240,14,MsgList(36),.ShowMsg
		PushButton 20,490,90,21,MsgList(37),.HelpButton
		PushButton 120,490,90,21,MsgList(38),.SetButton
		PushButton 220,490,110,21,MsgList(39),.SaveButton
		OKButton 450,490,90,21,.OKButton
		CancelButton 550,490,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub
	AllCont = 1
	AccKey = 0
	EndChar = 0
	Acceler = 0
	mCheckSrc = dlg.CheckSrc
	mCheckTrn = dlg.CheckTrn
	EngineID = dlg.EngineList
	CheckID = dlg.CheckList
	ProjectIDSrc = dlg.SrcProjectList
	ProjectIDTrn = dlg.TrnProjectList
	EngineName = EngineList(EngineID)
	srcLang = SrcLangList(dlg.SrcLangList)

	'根据是否更改语言重置字串
	If Temp <> tSelected(0) Then
		getMsgList(UIDataList,MsgList,"Main",1)
	End If

	'获取字串类型组合
	If dlg.Menu = 1 Then StrTypes = "|Menu|"
	If dlg.Dialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If dlg.Strings = 1 Then StrTypes = StrTypes & "|StringTable|"
	If dlg.AccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If dlg.Versions = 1 Then StrTypes = StrTypes & "|Version|"

	'提示打开关闭的翻译列表，以便可以在线翻译
	TransListOpen = False
	If trn.IsOpen = False Then
		If MsgBox(MsgList(3) & trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText) & _
				vbCrLf & vbCrLf & MsgList(48),vbYesNo,MsgList(40)) = vbNo Then
			EngineDataList = EngineDataListBak
			GoTo StartDlg
		End If
		PSL.Output MsgList(49)
		If trn.Open = False Then
			MsgBox MsgList(50),vbOkOnly+vbInformation,MsgList(42)
			Exit Sub
		Else
			TransListOpen = True
		End If
	End If

	'创建翻译引擎对象
	TempArray = ReSplit(EngineDataList(EngineID),JoinStr)
	TempDataList = ReSplit(TempArray(1),SubJoinStr)
	TempList = ReSplit(TempDataList(0),IIf(InStr(TempDataList(0),";"),";",","))
	For i = 0 To UBound(TempList)
		Temp = Trim(TempList(i))
		If Temp <> "" Then
			Set xmlHttp = CreateObject(Temp)
			If Not xmlHttp Is Nothing Then Exit For
		End If
	Next i

	'排序检查配置值并转换配置中的转义符
	If CheckArray(CheckDataList) = True Then
		CheckDataListBak = CheckDataList
		TempArray = ReSplit(CheckDataList(CheckID),JoinStr)
		TempDataList = ReSplit(TempArray(1),SubJoinStr)
		For i = 0 To UBound(TempDataList)
			If i <> 4 And i <> 14 And i <> 15 And i < 18 Then
				If TempDataList(i) <> "" Then
					If i = 1 Or i = 5 Or i = 13 Or i = 16 Or i = 17 Then
						If i = 5 Or i = 7 Then Temp = " " Else Temp = ","
						TempList = ReSplit(TempDataList(i),Temp,-1)
						Call SortArrayByLength(TempList,0,UBound(TempList),True)
						TempDataList(i) = Convert(Join(TempList,Temp))
					Else
						TempDataList(i) = Convert(TempDataList(i))
					End If
				End If
			End If
		Next i
		TempArray(1) = Join(TempDataList,SubJoinStr)
		CheckDataList(CheckID) = Join(TempArray,JoinStr)
		CheckName = TempArray(0)
	Else
		CheckName = MsgList(88)
	End If

	'获取检查方案设置
	If mCheckSrc = 1 Then
		TempArray = ReSplit(ProjectDataList(ProjectIDSrc),JoinStr)
		SrcProjectName = TempArray(0)
	Else
		SrcProjectName = MsgList(86)
	End If
	If mCheckTrn = 1 Then
		TempArray = ReSplit(ProjectDataList(ProjectIDTrn),JoinStr)
		TempDataList = ReSplit(TempArray(1),LngJoinStr)
		ShowOriginalTran = StrToLong(TempDataList(19))
		ApplyCheckResult = StrToLong(TempDataList(20))
		TrnProjectName = TempArray(0)
	Else
		TrnProjectName = MsgList(86)
	End If

	'设置行标志符数组
	LineSplitCharArr = ReSplit(Convert("\r\n,\r,\n"),",",-1)

	'释放不再使用的动态数组所使用的内存
	Erase SrcLangList,TempArray,trnTitleList
	Erase CheckList,CheckListBak,TempList,TempDataList
	Erase EngineList,EngineListBak,EngineDataListBak
	Erase DelLngNameList,DelSrcLngList,DelTranLngList
	Erase Tools,ToolsBak,FileDataList
	Erase UIFileList,PslLangDataList

	'根据是否选择 "仅选定字串" 项设置要翻译的字串数
	If trnsList Is Nothing Then
		If dlg.Seleted = 0 Then
			StringCount = trn.StringCount(pslDisplay)
		Else
			StringCount = trn.StringCount(pslSelection)
		End If
	Else
		If dlg.Seleted = 0 Then
			StringCount = trnsList.StringCount(pslDisplay)
		Else
			StringCount = trnsList.StringCount(pslSelection)
		End If
	End If

	'开始处理每条字串
	n = 0
	PSL.OutputWnd.Clear
	'PSL.Output MsgList(54) & vbCrLf & Replace$(Replace$(Replace$(Replace$(MsgList(85), _
	'		"%s",EngineName),"%d",CheckName),"%p",SrcProjectName),"%n",TrnProjectName)
	StartTime = Timer
	For j = 1 To StringCount
		'根据是否选择 "仅选定字串" 项设置要翻译的字串
		If trnsList Is Nothing Then
			If dlg.Seleted = 0 Then
				Set TransString = trn.String(j,pslDisplay)
			Else
				Set TransString = trn.String(j,pslSelection)
			End If
		Else
			If dlg.Seleted = 0 Then
				Set TransString = trnsList.String(j,pslDisplay)
			Else
				Set TransString = trnsList.String(j,pslSelection)
			End If
			Set trn = TransString.TransList
		End If

		'获取翻译列表 ID，以便对翻译列表的某些操作只进行一次
		TransID = trn.ListID
		If TransIDBak <> TransID Then
			TransIDBak = TransID
			'提示保存打开的翻译列表，以免处理后数据不可恢复
			'If trn.IsOpen = True Then
				'i = MsgBox(MsgList(3) & trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText) & _
				'		vbCrLf & vbCrLf & MsgList(51),vbYesNoCancel,MsgList(40))
				'If i = vbYes Then
				'	trn.Save
				'ElseIf i = vbCancel Then
				'	GoTo Skip
				'End If
			'End If
			'如果翻译列表的更改时间晚于原始列表，自动更新
			If trn.SourceList.LastChange > trn.LastChange Then
				PSL.Output MsgList(45)
				If trn.Update = False Then
					MsgBox(MsgList(3) & trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText) & _
						vbCrLf & vbCrLf & MsgList(46),vbOkOnly+vbInformation,MsgList(42))
					GoTo Skip
				End If
			End If
			'设置检查宏专用的用户定义属性
			If trn.Property(19980) <> "CheckAccessKeys" Then trn.Property(19980) = "CheckAccessKeys"
		End If

		'消息和字串初始化并获取翻译列表的现有来源和翻译字串
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		StringSrc.LineNum = 0
		StringTrn.LineNum = 0
		StringSrc.AccKeyNum = 0
		StringTrn.AccKeyNum = 0

		'字串类型处理
		If dlg.AllType = 0 And dlg.Seleted = 0 Then
			If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
				If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
			Else
				If dlg.Other = 0 Then GoTo Skip
			End If
		End If

		'跳过已锁定的字串
		If TransString.State(pslStateLocked) = True Then
			If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(59))
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过只读的字串
		If TransString.State(pslStateReadOnly) = True Then
			If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(60))
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过已翻译供复审的字串
		If dlg.ForReview = 1 Then
			If TransString.State(pslStateTranslated) = True Then
				If TransString.State(pslStateReview) = True Then
					If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(61))
					SkipedCount = SkipedCount + 1
					GoTo Skip
				End If
			End If
		End If
		'跳过已翻译并验证的字串
		If dlg.Validated = 1 Then
			If TransString.State(pslStateTranslated) = True Then
				If TransString.State(pslStateReview) = False Then
					If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(62))
					SkipedCount = SkipedCount + 1
					GoTo Skip
				End If
			End If
		End If
		'跳过未翻译的字串
		If dlg.NotTran = 1 Then
			If TransString.State(pslStateTranslated) = False Then
				If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(63))
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'跳过为空或全为空格的字串
		If Trim(TransString.SourceText) = "" Then
			If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(64))
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过全为数字和符号的字串
		If dlg.NumAndSymbol = 1 Then
			If LCase(TransString.SourceText) = UCase(TransString.SourceText) Then
				If CheckStrRegExp(TransString.SourceText,CheckSkipStr(1),0,1) = True Then
					If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(65))
					SkipedCount = SkipedCount + 1
					GoTo Skip
				End If
			End If
		End If
		'跳过全为大写英文的字串
		If dlg.AllUCase = 1 Then
			If UCase(TransString.SourceText) = TransString.SourceText Then
				If Trim$(TransString.SourceText) <> "OK" Then
					If CheckStrRegExp(TransString.SourceText,CheckSkipStr(2),0,1) = True Then
						If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(66))
						SkipedCount = SkipedCount + 1
						GoTo Skip
					End If
				End If
			End If
		End If
		'跳过全为小写英文的字串
		If dlg.AllLCase = 1 Then
			If LCase(TransString.SourceText) = TransString.SourceText Then
				If CheckStrRegExp(TransString.SourceText,CheckSkipStr(3),0,1) = True Then
					If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(55) & MsgList(72) & MsgList(67))
					SkipedCount = SkipedCount + 1
					GoTo Skip
				End If
			End If
		End If

		'获取翻译源文字串
		Set TrnList = trn
		If dlg.SrcLangList = 0 Then
			OldSrcString = TransString.SourceText
		Else
			For i = 1 To trn.Project.TransLists.Count
				Set TrnList = trn.Project.TransLists(i)
				If TrnList.SourceList.ListID = trn.SourceList.ListID Then
					If PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) = srcLang Then
						Exit For
					End If
				End If
			Next i
			If dlg.Seleted = 0 Then
				OldSrcString = TrnList.String(j,pslDisplay).Text
			Else
				OldSrcString = TrnList.String(TransString.Number,pslNumber).Text
			End If
		End If

		'转换转义符
		Stemp  = False
		srcString = Convert(OldSrcString)
		If srcString <> OldSrcString Then
			OldSrcString = srcString
			Stemp = True
		End If

		'开始预处理字串并翻译字串
		If dlg.PreStrRep = 1 Then srcString = ReplaceStr(CheckID,srcString,0,0)
		If dlg.SplitTran = 0 Then
			If mCheckSrc = 1 Then srcString = CheckHanding(CheckID,srcString,srcString,ProjectIDSrc)
			If InStr(srcString,"&") Then srcString = Replace(srcString,"&","")
			trnString = getTranslate(xmlHttp,EngineDataList,EngineID,srcString,LangPair,0)
		Else
			trnString = SplitTran(xmlHttp,EngineDataList,srcString,LangPair,EngineID,CheckID,ProjectIDSrc,mCheckSrc,n,0)
		End If

		'开始后处理字串并替换原有翻译
		If Trim(trnString) <> "" Then
			If trnString <> OldSrcString Or trnString <> TransString.Text Then
				If mCheckTrn = 1 Then
					If ApplyCheckResult = 1 Then
						trnString = CheckHanding(CheckID,OldSrcString,trnString,ProjectIDTrn)
					Else
						Call CheckHanding(CheckID,OldSrcString,trnString,ProjectIDTrn)
					End If
				End If
				If dlg.AppStrRep = 1 Then trnString = ReplaceStr(CheckID,trnString,2,1)
			End If
			If trnString <> OldSrcString And trnString <> TransString.Text Then
				If Stemp = True Then trnString = ReConvert(trnString)
				TransString.Text = trnString
				TransString.State(pslStateReview) = True
				If dlg.TranComment = 1 Then
					TransString.TransComment = Replace(MsgList(47),"%s",EngineName)
				Else
					TransString.TransComment = ""
				End If
				TranedCount = TranedCount + 1
				'组织消息并输出
				If dlg.ShowMsg = 1 Then
					'计算行数
					For i = 0 To UBound(LineSplitCharArr)
						FindStr = Trim(LineSplitCharArr(i))
						If InStr(TransString.SourceText,FindStr) Then
							StringSrc.LineNum = StringSrc.LineNum + UBound(ReSplit(TransString.SourceText,FindStr,-1))
						End If
						If InStr(trnString,FindStr) Then
							StringTrn.LineNum = StringTrn.LineNum + UBound(ReSplit(trnString,FindStr,-1))
						End If
					Next i
					'计算消息
					If StringSrc.LineNum <> StringTrn.LineNum Then
						LineMsg = LineErrMassage(StringSrc.LineNum,StringTrn.LineNum,LineNumErrCount)
					End If
					i = StringSrc.AccKeyNum - StringTrn.AccKeyNum
					If (i > 1 Or i < 1) Then
						AcckeyMsg = AccKeyErrMassage(StringSrc.AccKeyNum,StringTrn.AccKeyNum,accKeyNumErrCount)
					End If
					If mCheckTrn = 1 Or dlg.PreStrRep = 1 Or dlg.AppStrRep = 1 Then
						ChangeMsg = ReplaceMassage(CheckID,ProjectIDTrn)
					End If
					If AcckeyMsg & LineMsg & ChangeMsg <> "" Then
						TransString.OutputError(MsgList(56) & MsgList(72) & MsgList(73) & ChangeMsg & AcckeyMsg & LineMsg)
					Else
						TransString.OutputError(MsgList(56) & MsgList(74))
					End If
				End If
			Else
				NotChangeCount = NotChangeCount + 1
				'组织消息并输出
				If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(57))
			End If
		Else
			NotTranCount = NotTranCount + 1
			'组织消息并输出
			If dlg.ShowMsg = 1 Then TransString.OutputError(MsgList(58))
		End If
		'n = 1
		Skip:
	Next j
	Set xmlHttp = Nothing

	'翻译计数及消息输出
	ErrorCount = ModifiedCount + AddedCount + MovedCount + DeledCount + _
				ReplacedCount + LineNumErrCount + accKeyNumErrCount
	PSL.Output TranMassage(TranedCount,SkipedCount,NotChangeCount,NotTranCount,ErrorCount)
	'If ErrorCount = 0 And TransListOpen = True Then trn.Close
	PSL.Output MsgList(68) & Format(DateAdd("s",Timer - StartTime,0),MsgList(69))
	Call ExitMacro(0)
	Exit Sub

	'显示程序错误消息
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
	Call ExitMacro(0)
End Sub


'安全退出程序
Sub ExitMacro(ByVal Mode As Long)
	On Error Resume Next
	If Dir(trn.Project.Location & "\~temp.txt") <> "" Then Kill trn.Project.Location & "\~temp.txt"
	'取消检查宏专用的用户定义属性的设置
	For i = 1 To PSL.ActiveProject.TransLists.Count
		Set trn = PSL.ActiveProject.TransLists(i)
		If trn.IsOpen = True Then
			If trn.Property(19980) = "CheckAccessKeys" Then trn.Property(19980) = ""
		End If
	Next i
	If Mode > 0 Then Exit All
End Sub


'主对话框函数
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,n As Long
	Dim srcLng As String,trnLng As String,Temp As String
	Dim TrnList As PslTransList,Stemp As Boolean,xmlHttp As Object
	Dim MsgList() As String,TempList() As String,TempArray() As String
	Select Case Action%
	Case 1
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		If CheckArray(tSelected) = True Then
			DlgText "EngineList",tSelected(1)
			DlgText "CheckList",tSelected(2)
			DlgValue "AllType",StrToLong(tSelected(3))
			DlgValue "Menu",StrToLong(tSelected(4))
			DlgValue "Dialog",StrToLong(tSelected(5))
			DlgValue "Strings",StrToLong(tSelected(6))
			DlgValue "AccTable",StrToLong(tSelected(7))
			DlgValue "Versions",StrToLong(tSelected(8))
			DlgValue "Other",StrToLong(tSelected(9))
			DlgValue "Seleted",StrToLong(tSelected(10))
			DlgValue "ForReview",StrToLong(tSelected(11))
			DlgValue "Validated",StrToLong(tSelected(12))
			DlgValue "NotTran",StrToLong(tSelected(13))
			DlgValue "NumAndSymbol",StrToLong(tSelected(14))
			DlgValue "AllUCase",StrToLong(tSelected(15))
			DlgValue "AllLCase",StrToLong(tSelected(16))
			DlgValue "AutoSelection",StrToLong(tSelected(17))
			DlgValue "SrcProjectList",StrToLong(tSelected(18))
			DlgValue "CheckSrc",StrToLong(tSelected(19))
			DlgValue "PreStrRep",StrToLong(tSelected(20))
			DlgValue "SplitTran",StrToLong(tSelected(21))
			DlgValue "TrnProjectList",StrToLong(tSelected(22))
			DlgValue "CheckTrn",StrToLong(tSelected(23))
			DlgValue "AppStrRep",StrToLong(tSelected(24))
			DlgValue "KeepSet",StrToLong(tSelected(25))
			DlgValue "ShowMsg",StrToLong(tSelected(26))
			DlgValue "TranComment",StrToLong(tSelected(27))
		End If
		If DlgText("EngineList") = "" Then DlgValue "EngineList",0
		If DlgText("CheckList") = "" Then DlgValue "CheckList",0
		If DlgValue("AllType") + DlgValue("Menu") + DlgValue("Dialog") + DlgValue("Strings") + _
			DlgValue("AccTable") + DlgValue("Versions") + DlgValue("Other") + DlgValue("Seleted") = 0 Then
			DlgValue "Seleted",1
		End If
		If trn.IsOpen = False Then
			DlgEnable "Seleted",False
			DlgValue "Seleted",0
		End If
		If DlgValue("AutoSelection") = 1 Then
			'获取目标语言
			trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
			Temp = IIf(trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko","Asia","")
			If trnLng = "" Or trnLng = "zh" Then
				trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
				If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then
					trnLng = "zh-CN"
				ElseIf trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then
					trnLng = "zh-TW"
				End If
			Else
				trnLng = trnLng & LngJoinStr & PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
			End If
			DlgValue "CheckList",getCheckID(CheckDataList,trnLng,Temp)
			DlgEnable "CheckList",False
		End If
		Stemp = CheckNullData(DlgText("EngineList"),EngineDataList,"1,6-9,15-19",1)
		If Stemp = False Then Stemp = CheckTargetValue(EngineDataList,DlgValue("EngineList"))
		If Stemp = True Then DlgEnable "OKButton",False

		Stemp = CheckNullData(DlgText("CheckList"),CheckDataList,"1,4,14-17",2)
		If Stemp = False Then If CheckArray(ProjectDataList) = False Then Stemp = True
		If Stemp = True Then
			DlgValue "CheckSrc",0
			DlgValue "PreStrRep",0
			DlgValue "CheckTrn",0
			DlgValue "AppStrRep",0
			DlgEnable "CheckSrc",False
			DlgEnable "PreStrRep",False
			DlgEnable "CheckTrn",False
			DlgEnable "AppStrRep",False
		End If
		DlgEnable "SaveButton",False
		DlgEnable "LineText1",False
		DlgEnable "LineText2",False
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		MainDlgFunc = True ' 防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			MainDlgFunc = False
			Exit Function
		Case "HelpButton"
			If getMsgList(UIDataList,MsgList,"MainDlgFunc",0) = False Then Exit Function
			ReDim TempList(2) As String
			TempList(0) = MsgList(11)
			TempList(1) = MsgList(12)
			TempList(2) = MsgList(13)
			Select Case ShowPopupMenu(TempList,vbPopupUseRightButton)
			Case 0
				Call Help("MainHelp")
			Case 1
				i = Download(tUpdateSet,tUpdateSet(1),3)
				If i = 0 Then Exit Function
				If tUpdateSet(5) < Format(Date,"yyyy-MM-dd") Then
					Stemp = True
				ElseIf i = 3 And ArrayComp(tUpdateSet,tUpdateSetBak) = False Then
					Stemp = True
				End If
				If Stemp = True Then
					tUpdateSet(5) = Format(Date,"yyyy-MM-dd")
					If WriteEngineSet(tWriteLoc,"Update") = False Then Exit Function
					tUpdateSetBak(5) = tUpdateSet(5)
				End If
				If i = 3 Then Call ExitMacro(1)
			Case 2
				Call Help("About")
			End Select
			Exit Function
		Case "OKButton"
			If getMsgList(UIDataList,MsgList,"MainDlgFunc",0) = False Then Exit Function
			'检测字串类型和内容选择是否为空
			If DlgValue("AllType") + DlgValue("Menu") + DlgValue("Dialog") + DlgValue("Strings") + _
				DlgValue("AccTable") + DlgValue("Versions") + DlgValue("Other") + DlgValue("Seleted") = 0 Then
				MsgBox(MsgList(1),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'检测 Microsoft.XMLHTTP 是否存在
			On Error Resume Next
			TempArray = ReSplit(EngineDataList(DlgValue("EngineList")),JoinStr)
			TempDataList = ReSplit(TempArray(1),SubJoinStr)
			TempList = ReSplit(TempDataList(0),IIf(InStr(TempDataList(0),";"),";",","))
			For i = 0 To UBound(TempList)
				Temp = Trim(TempList(i))
				If Temp <> "" Then
					Set xmlHttp = CreateObject(Temp)
					If Not xmlHttp Is Nothing Then Exit For
				End If
			Next i
			On Error GoTo 0
			If xmlHttp Is Nothing Then
				Err.Source = Join(TempList,"; ")
				Call sysErrorMassage(Err,1)
				Exit Function
			End If
			'选择翻译来源列表
			Set TrnList = trn
			If DlgValue("SrcLangList") <> 0 Then
				Temp = DlgText("SrcLangList")
				For i = 1 To trn.Project.TransLists.Count
					Set TrnList = trn.Project.TransLists(i)
					If TrnList.SourceList.ListID = trn.SourceList.ListID Then
						If PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) = Temp Then
							Exit For
						End If
					End If
				Next i
			End If
			'获取PSL的来源语言代码
			If DlgValue("SrcLangList") = 0 Then
				srcLng = PSL.GetLangCode(trn.SourceList.LangID,pslCode639_1)
			Else
				srcLng = PSL.GetLangCode(TrnList.Language.LangID,pslCode639_1)
			End If
			If srcLng = "" Or srcLng = "zh" Then
				If DlgValue("SrcLangList") = 0 Then
					srcLng = PSL.GetLangCode(trn.SourceList.LangID,pslCodeLangRgn)
				Else
					srcLng = PSL.GetLangCode(TrnList.Language.LangID,pslCodeLangRgn)
				End If
				If srcLng = "zh-CHS" Or srcLng = "zh-SG" Then
					srcLng = "zh-CN"
				ElseIf srcLng = "zh-CHT" Or srcLng = "zh-HK" Or srcLng = "zh-MO" Then
					srcLng = "zh-TW"
				End If
			Else
				srcLng = srcLng & LngJoinStr & PSL.GetLangCode(trn.SourceList.LangID,pslCodeLangRgn)
			End If
			'获取目标语言
			trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
			If trnLng = "" Or trnLng = "zh" Then
				trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
				If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then
					trnLng = "zh-CN"
				ElseIf trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then
					trnLng = "zh-TW"
				End If
			Else
				trnLng = trnLng & LngJoinStr & PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
			End If
			'查找翻译引擎中对应的语言代码
			'TempArray = ReSplit(EngineDataList(DlgValue("EngineList")),JoinStr)
			LangPair = getEngineLngPair(ReSplit(TempArray(2),SubLngJoinStr),srcLng,trnLng)
			If LangPair = "" Then
				MsgBox Replace$(MsgList(10),"%s",TempArray(0)),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			'转换翻译引擎的配置中的转义符
			EngineDataListBak = EngineDataList
			'TempArray = ReSplit(EngineDataList(DlgValue("EngineList")),JoinStr)
			'TempDataList = ReSplit(TempArray(1),SubJoinStr)
			For i = 0 To UBound(TempDataList)
				If i > 10 And i < 19 Then
					If TempDataList(i) <> "" Then
						TempDataList(i) = Convert(TempDataList(i))
					End If
				End If
			Next i
			TempArray(1) = Join(TempDataList,SubJoinStr)
			EngineDataListBak(DlgValue("EngineList")) = Join(TempArray,JoinStr)
			'获取测试翻译
			trnString = getTranslate(xmlHttp,EngineDataListBak,DlgValue("EngineList"),"Testing at " & Time,LangPair,3)
			Set xmlHttp = Nothing
			'测试 Internet 连接
			If trnString = "NotConnected" Then
				MsgBox(MsgList(7),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'测试引擎网址是否为空
			If trnString = "NullUrl" Then
				MsgBox(MsgList(8),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'测试引擎引擎是否超时
			If trnString = "Timeout" Then
				i = MsgBox(MsgList(4),vbYesNoCancel+vbInformation,MsgList(3))
				If i = vbYes Then
					Temp = InputBox(MsgList(5),MsgList(6),"5")
					If Temp <> "" Then
						WaitTimes = CLng(Temp)
					Else
						Exit Function
					End If
				ElseIf i = vbCancel Then
					Exit Function
				End If
			End If
			'测试引擎结果是否为空
			If Trim(trnString) = "" Then
				MsgBox(MsgList(9),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'保存选择
			If DlgValue("KeepSet") = 1 Then
				If ArrayComp(tSelected,tSelectedBak) = True Then
			 		If WriteEngineSet(tWriteLoc,"Main") = False Then
						MsgBox Replace$(MsgList(2),"%s",tWriteLoc),vbOkOnly+vbInformation,MsgList(0)
						Exit Function
					Else
						tSelectedBak = tSelected
					End If
				End If
			End If
			EngineDataList = EngineDataListBak
			MainDlgFunc = False
			Exit Function
		Case "SaveButton"
			If ArrayComp(tSelected,tSelectedBak) = True Then
				If WriteEngineSet(tWriteLoc,"Main") = False Then
					If getMsgList(UIDataList,MsgList,"MainDlgFunc",0) = False Then Exit Function
					MsgBox Replace$(MsgList(2),"%s",tWriteLoc),vbOkOnly+vbInformation,MsgList(0)
				Else
					tSelectedBak = tSelected
					DlgEnable "SaveButton",False
				End If
			End If
			Exit Function
		Case "Menu", "Dialog", "Strings", "AccTable", "Versions", "Other"
			If DlgValue("Menu") + DlgValue("Dialog") + DlgValue("Strings") + DlgValue("AccTable") + _
				DlgValue("Versions") + DlgValue("Other") = 0 Then
				If DlgValue("Seleted") = 0 Then DlgValue "AllType",1
			Else
				If DlgValue("Seleted") <> 0 Then DlgValue "AllType",0
				DlgValue "Seleted",0
			End If
		Case "AllType"
			If DlgValue("AllType") = 1 Then
				DlgValue "Menu",0
				DlgValue "Dialog",0
				DlgValue "Strings",0
				DlgValue "AccTable",0
				DlgValue "Versions",0
				DlgValue "Other",0
				DlgValue "Seleted",0
			End If
		Case "Seleted"
			If DlgValue("Seleted") = 1 Then
				DlgValue "AllType",0
				DlgValue "Menu",0
				DlgValue "Dialog",0
				DlgValue "Strings",0
				DlgValue "AccTable",0
				DlgValue "Versions",0
				DlgValue "Other",0
			Else
				DlgValue "AllType",1
			End If
		Case "AutoSelection"
			If DlgValue("AutoSelection") = 1 Then
				trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
				Temp = IIf(trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko","Asia","")
				If trnLng = "" Or trnLng = "zh" Then
					trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
					If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then
						trnLng = "zh-CN"
					ElseIf trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then
						trnLng = "zh-TW"
					End If
				Else
					trnLng = trnLng & LngJoinStr & PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
				End If
				DlgValue "CheckList",getCheckID(CheckDataList,trnLng,Temp)
				DlgEnable "CheckList",False
			Else
				DlgText "CheckList",IIf(tSelected(2) = "",CheckList(0),tSelected(2))
				DlgEnable "CheckList",True
			End If
		Case "SrcEditButton", "TrnEditButton"
			ProjectListBak = ProjectList
			ProjectDataListBak = ProjectDataList
			If DlgItem$ = "SrcEditButton" Then
				Call Projects(DlgValue("SrcProjectList"))
			Else
				Call Projects(DlgValue("TrnProjectList"))
			End If
			DlgListBoxArray "SrcProjectList",ProjectList()
			DlgListBoxArray "TrnProjectList",ProjectList()
			DlgValue "SrcProjectList",StrToLong(tSelected(18))
			DlgValue "TrnProjectList",StrToLong(tSelected(22))
			If DlgText("SrcProjectList") = "" Then DlgValue "SrcProjectList",0
			If DlgText("TrnProjectList") = "" Then DlgValue "TrnProjectList",0
		Case "SetButton"
			EngineListBak = EngineList
			EngineDataListBak = EngineDataList
			CheckListBak = CheckList
			CheckDataListBak = CheckDataList
			tUpdateSetBak = tUpdateSet
			TempList = tSelected
			ReDim tmpLFList(0) As LOG_FONT
			tmpLFList(0) = LFList(0)
			Call Settings(DlgValue("EngineList"),DlgValue("CheckList"),0)
			DlgListBoxArray "EngineList",EngineList()
			DlgListBoxArray "CheckList",CheckList()
			DlgText "EngineList",tSelected(1)
			If DlgValue("AutoSelection") = 1 Then
				'获取目标语言
				trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
				Temp = IIf(trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko","Asia","")
				If trnLng = "" Or trnLng = "zh" Then
					trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
					If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then
						trnLng = "zh-CN"
					ElseIf trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then
						trnLng = "zh-TW"
					End If
				Else
					trnLng = trnLng & LngJoinStr & PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
				End If
				DlgValue "CheckList",getCheckID(CheckDataList,trnLng,Temp)
			Else
				DlgText "CheckList",tSelected(2)
			End If
			If DlgText("EngineList") = "" Then DlgValue "EngineList",0
			If DlgText("CheckList") = "" Then DlgValue "CheckList",0
			If TempList(0) <> tSelected(0) Then
				If getMsgList(UIDataList,MsgList,"Main",1) = True Then
					n = 0
					ReDim TempList(PSL.ActiveProject.TransLists.Count) As String
					For i = 1 To PSL.ActiveProject.TransLists.Count
						Set TrnList = PSL.ActiveProject.TransLists(i)
						If TrnList.IsOpen = True Then
							TempList(n) = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
							n = n + 1
						End If
					Next i
					If n > 0 Then
						ReDim Preserve TempList(n - 1) As String
					Else
						ReDim TempList(0) As String
						TempList(0) = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
					End If
					DlgText -1,MsgList(1)
					DlgText "Text1",Replace$(MsgList(0),"%s",Version)
					DlgText "Text2",MsgList(2)
					DlgText "Text3",MsgList(3) & Join(TempList,"; ")

					DlgText "Configuration",MsgList(4)
					DlgText "SrcLang",MsgList(5)

					DlgText "StrTypeSelection",MsgList(6)
					DlgText "AllType",MsgList(7)
					DlgText "Menu",MsgList(8)
					DlgText "Dialog",MsgList(9)
					DlgText "Strings",MsgList(10)
					DlgText "AccTable",MsgList(11)
					DlgText "Versions",MsgList(12)
					DlgText "Other",MsgList(13)
					DlgText "Seleted",MsgList(14)

					DlgText "SkipSelection",MsgList(15)
					DlgText "ForReview",MsgList(16)
					DlgText "Validated",MsgList(17)
					DlgText "NotTran",MsgList(18)
					DlgText "NumAndSymbol",MsgList(19)
					DlgText "AllUCase",MsgList(20)
					DlgText "AllLCase",MsgList(21)

					DlgText "PreProcessing",MsgList(22)
					DlgText "CheckConfigText",MsgList(23)
					DlgText "PreTrnText",MsgList(24)
					DlgText "AppTrnText",MsgList(25)
					DlgText "LineText1",String$(200,MsgList(26))
					DlgText "LineText2",String$(200,MsgList(26))
					DlgText "AutoSelection",MsgList(27)
					DlgText "SrcEditButton",MsgList(28)
					DlgText "TrnEditButton",MsgList(28)
					DlgText "CheckSrc",MsgList(29)
					DlgText "PreStrRep",MsgList(30)
					DlgText "SplitTran",MsgList(31)
					DlgText "CheckTrn",MsgList(32)
					DlgText "AppStrRep",MsgList(33)
					DlgText "TranComment",MsgList(34)
					DlgListBoxArray "SrcProjectList",ProjectList()
					DlgListBoxArray "TrnProjectList",ProjectList()
					DlgValue "SrcProjectList",StrToLong(tSelected(18))
					DlgValue "TrnProjectList",StrToLong(tSelected(22))

					DlgText "KeepSet",MsgList(35)
					DlgText "ShowMsg",MsgList(36)
					DlgText "HelpButton",MsgList(37)
					DlgText "SetButton",MsgList(38)
					DlgText "SaveButton",MsgList(39)
				End If
			End If
			'判断对话框字体是否已被改变
			If FontComp(LFList(0),tmpLFList(0)) = True Then
				n = CLng(DlgText("SuppValueBox"))
				j = CreateFont(n,LFList(0))
				If j = 0 Then Exit Function
				For i = 0 To DlgCount() - 1
					SendMessageLNG(GetDlgItem(n,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
				Next i
				DrawWindow(n,j)
			End If
			Stemp = CheckNullData(DlgText("EngineList"),EngineDataList,"1,6-9,15-19",1)
			If Stemp = False Then Stemp = CheckTargetValue(EngineDataList,DlgValue("EngineList"))
			DlgEnable "OKButton",IIf(Stemp = True,False,True)
			If CheckNullData(DlgText("CheckList"),CheckDataList,"1,4,14-17",2) = True Then
				DlgValue "CheckSrc",0
				DlgValue "PreStrRep",0
				DlgValue "CheckTrn",0
				DlgValue "AppStrRep",0
				DlgEnable "CheckSrc",False
				DlgEnable "PreStrRep",False
				DlgEnable "CheckTrn",False
				DlgEnable "AppStrRep",False
			Else
				DlgEnable "CheckSrc",True
				DlgEnable "PreStrRep",True
				DlgEnable "CheckTrn",True
				DlgEnable "AppStrRep",True
			End If
		End Select
		tSelected(1) = DlgText("EngineList")
		tSelected(2) = DlgText("CheckList")
		tSelected(3) = DlgValue("AllType")
		tSelected(4) = DlgValue("Menu")
		tSelected(5) = DlgValue("Dialog")
		tSelected(6) =  DlgValue("Strings")
		tSelected(7) = DlgValue("AccTable")
		tSelected(8) = DlgValue("Versions")
		tSelected(9) = DlgValue("Other")
		tSelected(10) = DlgValue("Seleted")
		tSelected(11) = DlgValue("ForReview")
		tSelected(12) = DlgValue("Validated")
		tSelected(13) = DlgValue("NotTran")
		tSelected(14) = DlgValue("NumAndSymbol")
		tSelected(15) = DlgValue("AllUCase")
		tSelected(16) = DlgValue("AllLCase")
		tSelected(17) = DlgValue("AutoSelection")
		tSelected(18) = DlgValue("SrcProjectList")
		tSelected(19) = DlgValue("CheckSrc")
		tSelected(20) = DlgValue("PreStrRep")
		tSelected(21) = DlgValue("SplitTran")
		tSelected(22) = DlgValue("TrnProjectList")
		tSelected(23) = DlgValue("CheckTrn")
		tSelected(24) = DlgValue("AppStrRep")
		tSelected(25) = DlgValue("KeepSet")
		tSelected(26) = DlgValue("ShowMsg")
		tSelected(27) = DlgValue("TranComment")
		DlgEnable "SaveButton",ArrayComp(tSelected,tSelectedBak,"1-" & CStr$(UBound(tSelected)))
	End Select
End Function


'检测并下载新版本
Public Function CheckUpdate(UpdateSet() As String) As Boolean
	Dim i As Long,j As Long,n As Long
	'获取更新数据并检查新版本
	If CheckArray(UpdateSet) = True Then
		If UpdateSet(0) = "" Then UpdateSet(0) = "1"
		If UpdateSet(1) = "" Then UpdateSet(1) = updateMainUrl & vbCrLf & updateMinorUrl
		If UpdateSet(2) = "" Or (UpdateSet(2) <> "" And Dir$(UpdateSet(2)) = "") Then
			getCMDPath(".rar",UpdateSet(2),UpdateSet(3))
		End If
	Else
		UpdateSet = ReSplit("1" & JoinStr & updateMainUrl & vbCrLf & updateMinorUrl & JoinStr & _
					getCMDPath(".rar","","") & JoinStr & "7" & JoinStr,JoinStr)
	End If
	If UpdateSet(0) <> "" And UpdateSet(0) <> "2" Then
		If UpdateSet(5) <> "" Then
			i = CLng(DateDiff("d",CDate(UpdateSet(5)),Date))
			j = StrComp(Format(Date,"yyyy-MM-dd"),UpdateSet(5))
			If UpdateSet(4) <> "" Then n = i - CLng(UpdateSet(4))
		End If
		If UpdateSet(5) = "" Or (j = 1 And n >= 0) Then
			i = Download(UpdateSet,UpdateSet(1),StrToLong(UpdateSet(0)))
			If i > 0 Then
				If UpdateSet(5) < Format(Date,"yyyy-MM-dd") Then
					UpdateSet(5) = Format(Date,"yyyy-MM-dd")
					WriteEngineSet(tWriteLoc,"Update")
				End If
				If i = 3 Then CheckUpdate = True
			End If
		End If
	End If
End Function


'检测并下载新版本
'Mode: 0 = 自动下载并安装, 1 = 用户决定, 2 = 关闭, 3 = 手动检查 4 = 测试
'返回值 = 0 失败, 1 = 检查更新信息成功, 2 = 下载文件成功, 3 = 更新成功
Public Function Download(UpdateSet() As String,ByVal Url As String,ByVal Mode As Long) As Long
	Dim i As Long,j As Long,m As Long,n As Long
	Dim xmlHttp As Object,Body() As Byte,BodyBak() As Byte
	Dim UrlList() As String,TempList() As String,TempArray() As String,MsgList() As String
	Dim ExePath As String,Argument As String,Temp As String,UpdateData As INIFILE_DATA
	Dim WebVersion As String,TempPath As String

	If Mode = 2 Then Exit Function
	If getMsgList(UIDataList,MsgList,"Download",1) = False Then Exit Function

	'获取解压程序和参数
	If CheckArray(UpdateSet) = True Then
		If Url = "" Then Url = UpdateSet(1)
		ExePath = Trim$(UpdateSet(2))
		Argument = UpdateSet(3)
	End If

	'检查解压程序和参数
	PSL.OutputWnd.Clear
	If Mode = 4 Then PSL.Output MsgList(22) Else PSL.Output MsgList(23)
	If ExePath = "" Then
		MsgBox(IIf(Mode <> 4,MsgList(1),MsgList(2)) & vbCrLf & MsgList(3),vbOkOnly+vbInformation,MsgList(0))
		Exit Function
	End If
	If Url = "" Or Argument = "" Then
		i = 1
	ElseIf InStr(Argument,"%1") = 0 Then
		i = 1
	ElseIf InStr(Argument,"%2") = 0 Then
		i = 1
	ElseIf InStr(Argument,"%3") = 0 Then
		i = 1
	End If
	If i = 1 Then
		MsgBox(IIf(Mode <> 4,MsgList(1),MsgList(2)) & vbCrLf & MsgList(4),vbOkOnly+vbInformation,MsgList(0))
		Exit Function
	End If

	'检测下载服务是否存在
	On Error Resume Next
	TempList = ReSplit(DefaultObject,";")
	For i = 0 To UBound(TempList)
		Set xmlHttp = CreateObject(TempList(i))
		If Not xmlHttp Is Nothing Then Exit For
	Next i
	If xmlHttp Is Nothing Then
		Err.Source = Join(TempList,"; ")
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
	On Error GoTo 0

	'获取更新配置信息
	If Mode <> 4 Then
		'合并更新配置文件的默认和自定义下载网址
		UrlList = ReSplit(Url,vbCrLf)
		For i = 0 To UBound(UrlList)
			Temp = Trim$(UrlList(i))
			n = InStrRev(Temp,"/")
			If n > 0 Then
				UrlList(i) = Left$(Temp,n) & updateINIFile
			End If
		Next i
		UrlList = ReSplit(updateINIMainUrl & vbCrLf & updateINIMinorUrl & vbCrLf & Join$(UrlList,vbCrLf),vbCrLf)

		'下载并检查更新配置文件
		For i = 0 To UBound(UrlList)
			'返回值，1 = 成功，0 = 失败，-1 = 文件不存在，-2 = 错误
			Select Case DownloadFile(Body,xmlHttp,UrlList(i))
			Case 1
				Temp = BytesToBstr(Body,"utf-8")
				If Temp <> "" Then
					If InStr(LCase$(Temp),LCase$(AppName)) Then
						'检查更新配置文件
						If CheckUpdateINIFile(UpdateData,Temp) = True Then
							Exit For
						End If
					End If
				End If
			Case -2
				Set xmlHttp = Nothing
				MsgBox(MsgList(35),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End Select
		Next i

		'显示更新信息
		If UpdateData.Title <> "" Then
			Download = 1
			Select Case StrComp(UpdateData.Title,Version)
			Case Is > 0
				If Mode = 1 Or Mode = 3 Then
					Temp = Replace$(MsgList(15),"%s",UpdateData.Title) & vbCrLf & vbCrLf & MsgList(20)
					If MsgBox(Temp & vbCrLf & Join(UpdateData.Value,vbCrLf),vbYesNo+vbInformation,MsgList(17)) = vbNo Then
						Set xmlHttp = Nothing
						Exit Function
					End If
				End If
			Case 0
				If Mode = 3 Then
					Temp = Replace$(MsgList(14),"%s",UpdateData.Title) & vbCrLf & MsgList(21)
					If MsgBox(Temp,vbYesNo+vbInformation,MsgList(17)) = vbNo Then
						Set xmlHttp = Nothing
						Exit Function
					End If
				Else
					Set xmlHttp = Nothing
					Exit Function
				End If
			Case Is < 0
				If Mode = 3 Then
					MsgBox(Replace$(MsgList(14),"%s",UpdateData.Title),vbOkOnly+vbInformation,MsgList(16))
				End If
				Set xmlHttp = Nothing
				Exit Function
			End Select
		End If
	End If

	'下载程序文件
	If Mode <> 4 Then PSL.Output MsgList(24)
	m = 0: n = 0: j = 0
	If UpdateData.Title <> "" Then
		UrlList = ClearArray(ReSplit(Url & vbCrLf & Join(UpdateData.Item,vbCrLf),vbCrLf),1)
	Else
		UrlList = ReSplit(Url,vbCrLf)
	End If
	For i = 0 To UBound(UrlList)
		If UrlList(i) <> "" Then
			'返回值，1 = 成功，0 = 失败，-1 = 文件不存在，-2 = 错误
			Select Case DownloadFile(Body,xmlHttp,UrlList(i))
			Case 1
				If Mode <> 4 Then Exit For
				If LenB(BodyBak) = 0 Then BodyBak = Body
			Case 0
				If Mode = 4 Then
					ReDim Preserve TempList(m)
					TempList(m) = UrlList(i)
				End If
				m = m + 1
			Case -1
				If Mode = 4 Then
					ReDim Preserve TempArray(n)
					TempArray(n) = UrlList(i)
				End If
				n = n + 1
			Case -2
				MsgBox(MsgList(35),vbOkOnly+vbInformation,MsgList(0))
				Set xmlHttp = Nothing
				Exit Function
			End Select
			j = j + 1
		End If
	Next i
	If m + n <> 0 Then
		If Mode <> 4 Then
			If m = j Or n = j Then
				MsgBox(MsgList(1) & vbCrLf & IIf(n = j,MsgList(5),MsgList(6)),vbOkOnly+vbInformation,MsgList(0))
				Set xmlHttp = Nothing
				Exit Function
			End If
		Else
			If m <> 0 And n <> 0 Then
				Temp = MsgList(2) & vbCrLf & MsgList(33) & vbCrLf & Join(TempArray,vbCrLf) & _
						vbCrLf & vbCrLf & MsgList(34) & vbCrLf & Join(TempList,vbCrLf)
			ElseIf m <> 0 Then
				Temp = MsgList(2) & vbCrLf & MsgList(34) & vbCrLf & Join(TempList,vbCrLf)
			ElseIf n <> 0 Then
				Temp = MsgList(2) & vbCrLf & MsgList(33) & vbCrLf & Join(TempArray,vbCrLf)
			End If
			MsgBox(Temp,vbOkOnly+vbInformation,MsgList(12))
			Set xmlHttp = Nothing
			Exit Function
		End If
	End If
	Set xmlHttp = Nothing

	'保存下载的程序文件
	If Mode = 4 Then Body = BodyBak
	TempPath = MacroLoc & "\temp\"
	Temp = TempPath & "temp.rar"
	On Error Resume Next
	If Dir$(TempPath & "*.*") = "" Then MkDir TempPath
	If BytesToFile(Body,Temp) = False Then
		i = FreeFile
		Open Temp For Binary Access Write As #i
		Put #i,,Body
		Close #i
	End If
	On Error GoTo 0

	'解压文件
	i = 0
	If Dir$(Temp) <> "" Then
		i = ExtractFile(Temp,TempPath,ExePath,Argument)
		If i = 1 Then
			Temp = TempPath & updateMainFile
		ElseIf i = -2 Then
			Temp = Mid$(Left$(ExePath,InStrRev(ExePath,".") - 1),InStrRev(ExePath,"\") + 1)
			Temp = Replace$(MsgList(8),"%s",Temp) & vbCrLf & _
					Replace$(MsgList(9),"%s",ExePath) & vbCrLf & _
					Replace$(MsgList(10),"%s",Argument) & vbCrLf & vbCrLf & MsgList(11)
			MsgBox(Temp,vbOkOnly+vbInformation,MsgList(0))
		ElseIf i = -3 Then
			MsgBox(IIf(Mode <> 4,MsgList(1),MsgList(2)) & vbCrLf & MsgList(7),vbOkOnly+vbInformation,MsgList(0))
		End If
	End If
	If i <> 1 Then
		DelDir(TempPath)
		Exit Function
	End If

	'获取下载的程序版本号
	If Mode <> 4 Then PSL.Output MsgList(26)
	WebVersion = GetWebVersion(Temp,"Const Version = ")
	If WebVersion = "" Then
		MsgBox(MsgList(19),vbOkOnly+vbInformation,MsgList(0))
		DelDir(TempPath)
		Exit Function
	End If

	'比较版本，显示更新信息
	If Url <> Join$(UrlList,vbCrLf) Then UpdateSet(1) = Join$(UrlList,vbCrLf)
	If Mode = 4 Then
		MsgBox(MsgList(13) & vbCrLf & Replace$(MsgList(14),"%s",WebVersion),vbOkOnly+vbInformation,MsgList(12))
		Download = 2
		DelDir(TempPath)
		Exit Function
	End If
	n = StrComp(WebVersion,Version)
	If n = 1 Or (n = 0 And Mode = 3) Then
		If UpdateData.Title = "" Then
			If Mode = 1 Then
				Temp = Replace$(MsgList(15),"%s",WebVersion)
			ElseIf Mode = 3 Then
				If n = 0 Then
					Temp = Replace$(MsgList(14),"%s",WebVersion) & vbCrLf & MsgList(21)
				Else
					Temp = Replace$(MsgList(15),"%s",WebVersion)
				End If
			End If
			If MsgBox(Temp,vbYesNo+vbInformation,MsgList(17)) = vbNo Then
				DelDir(TempPath)
				Exit Function
			End If
		ElseIf UpdateData.Title <> WebVersion Then
			Temp = MsgList(1) & vbCrLf & Replace$(MsgList(29),"%s",WebVersion) & vbCrLf & MsgList(32)
			MsgBox(Temp,vbOkOnly+vbInformation,MsgList(0))
			DelDir(TempPath)
			Exit Function
		End If
	Else
		If Mode < 2 Then
			PSL.Output Replace$(MsgList(30),"%s",WebVersion)
		Else
			MsgBox(Replace$(MsgList(31),"%s",WebVersion),vbOkOnly+vbInformation,MsgList(16))
		End If
		DelDir(TempPath)
		Exit Function
	End If

	'安装新版本
	PSL.Output MsgList(27)
	If SetupNewVersion(TempPath,MacroLoc) = True Then
		PSL.Output MsgList(28)
		MsgBox(MsgList(18),vbOkOnly+vbInformation,MsgList(17))
		Download = 3
	End If
	DelDir(TempPath)
End Function


'从注册表中获取 RAR 扩展名的默认程序
Public Function getCMDPath(ByVal ExtName As String,CmdPath As String,Argument As String) As String
	Dim i As Long,WshShell As Object,TempArray() As String
	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		Err.Source = "WScript.Shell"
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
	ExtName = WshShell.RegRead("HKCR\" & ExtName & "\")
	If ExtName <> "" Then
		CmdPath = WshShell.RegRead("HKCR\" & ExtName & "\shell\open\command\")
	End If
	On Error GoTo 0
	Set WshShell = Nothing
	If CmdPath <> "" Then
		i = InStr(CmdPath,".")
		Argument = Trim$(Mid$(CmdPath,InStr(i,CmdPath," ")))
		CmdPath = Left$(CmdPath,Len(CmdPath) - Len(Argument))
		TempArray = ReSplit(CmdPath,"%")
		If UBound(TempArray) = 2 Then
			CmdPath = Replace$(CmdPath,"%" & TempArray(1) & "%",Environ(TempArray(1)),,1)
		End If
		CmdPath = RemoveBackslash(CmdPath,"""","""",1)

		If InStr(CmdPath,"\") = 0 Then
			If Dir$(Environ("SystemRoot") & "\system32\" & CmdPath) <> "" Then
				CmdPath = Environ("SystemRoot") & "\system32\" & CmdPath
			ElseIf Dir$(Environ("SystemRoot") & "\" & CmdPath) <> "" Then
				CmdPath = Environ("SystemRoot") & "\" & CmdPath
			End If
		End If

		If InStr(LCase$(CmdPath),"winrar.exe") Then
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e -ibck " & Replace$(Argument,"""%1""","""%1"" %2 ""%3""")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e -ibck " & Replace$(Argument,"%1","""%1"" %2 ""%3""")
				Else
					Argument = "e -ibck ""%1"" %2 ""%3"" " & Argument
				End If
			Else
				Argument = "e -ibck ""%1"" %2 ""%3"""
			End If
		ElseIf InStr(LCase$(CmdPath),"winzip.exe") Then
			CmdPath = strReplace(CmdPath,"WinZip.exe","WzunZip.exe")
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = Replace$(Argument,"""%1""","""%1"" %2 ""%3""")
				ElseIf InStr(Argument,"%1") Then
					Argument = Replace$(Argument,"%1","""%1"" %2 ""%3""")
				Else
					Argument = """%1"" %2 ""%3"" " & Argument
				End If
			Else
				Argument = " ""%1"" %2 ""%3"""
			End If
		ElseIf InStr(LCase$(CmdPath),"7z.exe") Then
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e " & Replace$(Argument,"""%1""","""%1"" -o""%3"" %2")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e " & Replace$(Argument,"%1","""%1"" -o""%3"" %2")
				Else
					Argument = "e ""%1"" -o""%3"" %2 " & Argument
				End If
			Else
				Argument = "e ""%1"" -o""%3"" %2"
			End If
		ElseIf InStr(LCase$(CmdPath),"haozip.exe") Then
			CmdPath = strReplace(CmdPath,"HaoZip.exe","HaoZipC.exe")
			If Argument <> "" Then
				If InStr(Argument,"""%1""") Then
					Argument = "e " & Replace$(Argument,"""%1""","""%1"" -r -o""%3"" %2")
				ElseIf InStr(Argument,"%1") Then
					Argument = "e " & Replace$(Argument,"%1","""%1"" -r -o""%3"" %2")
				Else
					Argument = "e ""%1"" -r -o""%3"" %2 " & Argument
				End If
			Else
				Argument = "e ""%1"" -r -o""%3"" %2"
			End If
		End If
	End If
	getCMDPath = CmdPath & JoinStr & Argument
End Function


'在 wTimes 等待时间内轮询服务器的状态
'tValue 为目标值，当 wTimes = 0 时为默认等待时间
Private Function OnReadyStateChange(xmlHttp As Object,ByVal tValue As Long,wTimes As Long) As Long
	Dim StartTime As Long
	StartTime = Timer
	If wTimes = 0 Then wTimes = DefaultWaitTimes
	OnReadyStateChange = xmlHttp.readyState
	Do While OnReadyStateChange < tValue
		OnReadyStateChange = xmlHttp.readyState
		If (Timer - StartTime) > wTimes Then Exit Do
	Loop
End Function


'转换二进制数据为指定编码格式的字符
Public Function BytesToBstr(strBody As Variant,ByVal outCode As String) As String
	Dim objStream As Object
	If LenB(strBody) = 0 Or outCode = "" Then Exit Function
	On Error GoTo ErrorMsg
	Set objStream = CreateObject("Adodb.Stream")
	If Not objStream Is Nothing Then
		With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Write strBody
			.Position = 0
			.Type = 2
			.CharSet = outCode
			BytesToBstr = .ReadText
			.Close
		End With
		Set objStream = Nothing
	End If
	Exit Function
	ErrorMsg:
	Err.Source = "Adodb.Stream"
	Call sysErrorMassage(Err,1)
End Function


'写入二进制数据到文件
Public Function BytesToFile(strBody As Variant,ByVal File As String) As Boolean
	Dim objStream As Object
	BytesToFile = False
	If LenB(strBody) = 0 Or File = "" Then Exit Function
	On Error GoTo ErrorMsg
	Set objStream = CreateObject("Adodb.Stream")
	If Not objStream Is Nothing Then
		With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Write(strBody)
			.Position = 0
			.SaveToFile File,2
			.Flush
			.Close
		End With
		Set objStream = Nothing
		BytesToFile = True
	End If
	Exit Function
	ErrorMsg:
	Err.Source = "Adodb.Stream"
	Call sysErrorMassage(Err,1)
End Function


'下载文件
'返回值，1 = 成功，0 = 失败，-1 = 文件不存在，-2 = 错误
Private Function DownloadFile(Body() As Byte,xmlHttp As Object,ByVal Url As String) As Long
	Dim FileSize As Long
	ReDim Body(0) As Byte
	If Trim$(Url) = "" Then Exit Function
	On Error GoTo ExitFunction
	xmlHttp.Open "HEAD",Url,False,"",""
	xmlHttp.send()
	If OnReadyStateChange(xmlHttp,4,DefaultWaitTimes) = 4 Then
		'FileSize = CLng(ReSplit(xmlHttp.getResponseHeader("Content-Range"),"/")(1))
		FileSize = CLng(xmlHttp.getResponseHeader("Content-Length"))
	End If
	xmlHttp.Abort
	If FileSize > 0 Then
		xmlHttp.Open "GET",Url,False,"",""
		xmlHttp.setRequestHeader "Referer", Left(Url, InStr(InStr(Url, "//") + 2, Url, "/") - 1)
		xmlHttp.setRequestHeader "Accept", "*/*"
		'xmlHttp.setRequestHeader "Range", "bytes = " & FileSize
		xmlHttp.setRequestHeader "Content-Type", "application/octet-stream"
		xmlHttp.setRequestHeader "If-Modified-Since", "0"
		xmlHttp.setRequestHeader "Pragma", "no-cache"
		xmlHttp.setRequestHeader "Cache-Control", "no-cache"
		xmlHttp.send()
		If OnReadyStateChange(xmlHttp,4,DefaultWaitTimes) = 4 Then
			If xmlHttp.Status = 200 Then
				Body = xmlHttp.responseBody
				If LenB(Body) = FileSize Then DownloadFile = 1
			End If
		End If
	Else
		DownloadFile = -1
	End If
	On Error GoTo 0
	ExitFunction:
	xmlHttp.Abort
	If Err.Number <> 0 Then DownloadFile = -2
End Function


'检查更新配置文件
Private Function CheckUpdateINIFile(Data As INIFILE_DATA,ByVal UpdateINIText As String) As Boolean
	Dim i As Long,j As Long,m As Long,n As Long
	Dim DefaultLng As String,LangName As String,DataList() As INIFILE_DATA

	If Trim$(UpdateINIText) = "" Then Exit Function
	If getINIFile(DataList,"",UpdateINIText,2) = False Then Exit Function
	For i = 0 To UBound(DataList)
		With DataList(i)
			Select Case .Title
			Case "Option"
				For j = 0 To UBound(.Item)
					If .Item(j) = "DefaultLanguage" Then
						DefaultLng = LCase$(Trim$(.Value(j)))
						Exit For
					End If
				Next j
			Case "Language"
				UpdateINIText = LCase$(OSLanguage)
				For j = 0 To UBound(.Item)
					If InStr(LCase$(.Value(j)),UpdateINIText) Then
						LangName = LCase$(.Item(j))
						Exit For
					End If
				Next j
				If LangName = "" Then LangName = DefaultLng
			Case AppName
				For j = 0 To UBound(.Item)
					UpdateINIText = LCase$(.Item(j))
					If UpdateINIText = "version" Then
						Data.Title = Trim$(.Value(j))
					ElseIf InStr(UpdateINIText,"url_") Then
						ReDim Preserve Data.Item(m)
						Data.Item(m) = Trim$(.Value(j))
						m = m + 1
					ElseIf InStr(UpdateINIText,"des_" & LangName) Then
						ReDim Preserve Data.Value(n)
						Data.Value(n) = Trim$(.Value(j))
						n = n + 1
					End If
				Next j
				Exit For
			End Select
		End With
	Next i
	If m = 0 Or n = 0 Then
		Data.Title = ""
		ReDim Data.Item(0),Data.Value(0)
	Else
		CheckUpdateINIFile = True
		ReDim Preserve Data.Item(m - 1),Data.Value(n - 1)
	End If
End Function


'解压文件
'返回值 1 = 成功，0 = 要解压的文件不存在或大小为零，-1 = 宏主程序找不到，-2 = 解压程序找不到，-3 = 解压错误
Private Function ExtractFile(ByVal File As String,ByVal Path As String,ByVal ExePath As String,ByVal Argument As String) As Long
	Dim WshShell As Object,TempList() As String

	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		Err.Source = "WScript.Shell"
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
	On Error GoTo 0

	If Dir$(File) = "" Then Exit Function
	If FileLen(File) = 0 Then Exit Function
	If ExePath <> "" Then
		TempList = ReSplit(ExePath,"%",-1)
		If UBound(TempList) >= 2 Then
			ExePath = Replace$(ExePath,"%" & TempList(1) & "%",Environ(TempList(1)),,1)
		End If
		ExePath = RemoveBackslash(ExePath,"""","""",1)
	End If
	If Argument <> "" Then
		If InStr(Argument,"""%1""") Then Argument = Replace$(Argument,"%1",File)
		If InStr(Argument,"""%2""") Then Argument = Replace$(Argument,"%2","*.*")
		If InStr(Argument,"""%3""") Then Argument = Replace$(Argument,"%3",Path)
		If InStr(Argument,"%1") Then Argument = Replace$(Argument,"%1","""" & File & """")
		If InStr(Argument,"%2") Then Argument = Replace$(Argument,"%2","*.*")
		If InStr(Argument,"%3") Then Argument = Replace$(Argument,"%3","""" & Path & """")
	End If
	If ExePath <> "" Then
		If Dir$(ExePath) <> "" Then
			If WshShell.Run("""" & ExePath & """ " & Argument,0,True) = 0 Then
				ExtractFile = IIf(Dir$(Path & updateMainFile) <> "",1,-1)
			Else
				ExtractFile = -3
			End If
		Else
			ExtractFile = -2
		End If
	Else
		ExtractFile = -2
	End If
	Set WshShell = Nothing
End Function


'获取下载的程序版本号
Private Function GetWebVersion(ByVal File As String,ByVal CheckStr As String) As String
	Dim n As Long,Temp As String,FN As Long
	On Error GoTo ExitFunction
	FN = FreeFile
	Open File For Input As #FN
	Do While Not EOF(FN)
		Line Input #FN,Temp
		n = InStr(Temp,CheckStr)
		If n > 0 Then
			GetWebVersion = Mid$(Temp,n + Len(CheckStr) + 1,10)
			Exit Do
		End If
	Loop
	ExitFunction:
	On Error Resume Next
	Close #FN
End Function


'安装新版本
Private Function SetupNewVersion(ByVal FromPath As String,ByVal TargetDir As String) As Boolean
	Dim i As Long,File As String,TempList() As String

	On Error GoTo ExitFunction
	If Right$(FromPath,1) <> "\" Then FromPath = FromPath & "\"
	'检查是否存在相应的目录
	If Dir$(FromPath & "*.lng") <> "" Or Dir$(FromPath & "*.ini") <> "" Then
		If Dir$(TargetDir & "\Data\" & "*.*") = "" Then MkDir TargetDir & "\Data\"
	End If
	If Dir$(FromPath & "*.txt") <> "" Then
		If Dir$(TargetDir & "\Doc\" & "*.*") = "" Then MkDir TargetDir & "\Doc\"
	End If
	If Dir$(FromPath & "mod*.bas") <> "" Then
		If Dir$(TargetDir & "\Module\" & "*.*") = "" Then MkDir TargetDir & "\Module\"
	End If

	'获取新版本的文件列表
	File = Dir$(FromPath & "*.*")
	Do While File <> ""
		ReDim Preserve TempList(i) As String
		TempList(i) = File
		i = i + 1
		File = Dir$()
	Loop
	If i = 0 Then Exit Function

	'复制文件到子文件文件夹
	For i = 0 To UBound(TempList)
		File = TempList(i)
		Select Case LCase$(Mid$(File,InStrRev(File,".") + 1))
		Case "bas"
			If File Like "mod*.bas" = False Then
				FileCopy FromPath & File,TargetDir & "\" & File
			Else
				FileCopy FromPath & File,TargetDir & "\Module\" & File
			End If
			Kill FromPath & File
		Case "lng", "dat"
			FileCopy FromPath & File,TargetDir & "\Data\" & File
			Kill FromPath & File
		Case "obm", "cls"
			FileCopy FromPath & File,TargetDir & "\Module\" & File
			Kill FromPath & File
		Case "txt"
			FileCopy FromPath & File,TargetDir & "\Doc\" & File
			Kill FromPath & File
			If Dir$(TargetDir & "\Data\" & File) <> "" Then
				Kill TargetDir & "\Data\" & File
			End If
		Case "ini"
			If Dir$(TargetDir & "\Data\" & File) <> "" Then
				If FileDateTime(FromPath & File) > FileDateTime(TargetDir & "\Data\" & File) Then
					FileCopy FromPath & File,TargetDir & "\Data\" & File
					Kill FromPath & File
				End If
			Else
				FileCopy FromPath & File,TargetDir & "\Data\" & File
				Kill FromPath & File
			End If
		End Select
	Next i
	On Error GoTo 0
	SetupNewVersion = True
	ExitFunction:
End Function


'删除文件夹
Public Function DelDir(ByVal DirPath As String) As Boolean
	Dim File As String
	DirPath = Trim$(DirPath)
	If DirPath = "" Then Exit Function
	If Right$(DirPath,1) <> "\" Then DirPath = DirPath & "\"
	File = Dir$(DirPath & "*.*")
	On Error GoTo errHandle
	Do While File <> ""
		Kill DirPath & File
		File = Dir$()
	Loop
	If Dir$(DirPath & "*.*") = "" Then RmDir DirPath
	DelDir = True
	errHandle:
End Function


'获取在线翻译
Function getTranslate(xmlHttp As Object,EngineData() As String,ByVal ID As Long,ByVal srcStr As String,ByVal LngPair As String,ByVal fType As Long) As String
	Dim i As Long,Pos As Long,LangFrom As String,LangTo As String,responseType As String
	Dim Code As String,Temp As String,SetsArray() As String,TempList() As String

	If Trim(srcStr) = "" Or LngPair = "" Then Exit Function
	SetsArray = ReSplit(ReSplit(EngineData(ID),JoinStr)(1),SubJoinStr)
	BodyData = SetsArray(8)
	responseType = LCase(SetsArray(10))
	Select Case responseType
	Case "responsetext"
		TranBeforeStr = SetsArray(11)
		TranAfterStr = SetsArray(12)
	Case "responsebody"
		TranBeforeStr = SetsArray(13)
		TranAfterStr = SetsArray(14)
	Case "responsestream"
		TranBeforeStr = SetsArray(15)
		TranAfterStr = SetsArray(16)
	Case "responsexml"
		TranBeforeStr = SetsArray(17)
		TranAfterStr = SetsArray(18)
	End Select

	If SetsArray(2) = "" Then
		If fType = 3 Then getTranslate = "NullUrl"
		Exit Function
	End If

	On Error GoTo ErrorHandler
	If fType = 2 Then
		xmlHttp.Open "HEAD",SetsArray(2),SetsArray(5),SetsArray(6),SetsArray(7)
		'xmlHttp.setRequestHeader("If-Modified-Since","0")
		xmlHttp.send()
		If OnReadyStateChange(xmlHttp,4,WaitTimes) = 4 Then
			getTranslate = xmlHttp.getAllResponseHeaders
		End If
		xmlHttp.Abort
		Exit Function
	End If

	If LngPair <> "" Then
		TempList = ReSplit(LngPair,LngJoinStr)
		LangFrom = TempList(0)
		LangTo = TempList(1)
	End If

	Pos = InStr(LCase(SetsArray(9)),"charset")
	If Pos > 0 Then
		TempList = ReSplit(Mid(SetsArray(9),Pos),vbCrLf)
		For i = 0 To UBound(TempList)
			Temp = TempList(i)
			If InStr(Temp,"=") Then
				Code = ExtractStr(Temp,"=",";|" & vbCrLf,1)
				If Code <> "" Then Exit For
			End If
		Next i
	Else
		xmlHttp.Open SetsArray(4),SetsArray(2),SetsArray(5),SetsArray(6),SetsArray(7)
		'xmlHttp.setRequestHeader("If-Modified-Since","0")
		xmlHttp.send()
		If OnReadyStateChange(xmlHttp,4,WaitTimes) = 4 Then
			Temp = xmlHttp.getResponseHeader("Content-Type")
			Pos = InStr(LCase(Temp),"charset")
			If Pos > 0 Then
				Temp = Mid(Temp,Pos)
				If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",";|" & vbCrLf,1)
			Else
				Temp = xmlHttp.responseText
				Pos = InStr(LCase(Temp),"charset")
				If Pos = 0 Then Pos = InStr(LCase(Temp),"lang")
				If Pos > 0 Then
					Temp = Mid(Temp,Pos)
					If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",">",1)
				End If
			End If
		End If
		xmlHttp.Abort
	End If
	Dim nodeSrcStr
	nodeSrcStr = srcStr
	If Code <> "" Then Code = RemoveBackslash(Code,"""","""",1)
	If LCase(Code) = "utf-8" Or LCase(Code) = "utf8" Then
		srcStr = Str2URLEsc(srcStr,CP_UTF8,1)	'srcStr = Utf8Encode(srcStr)
	Else
		srcStr = Str2URLEsc(srcStr,GetACP,1)	'srcStr = ANSIEncode(srcStr)
	End If

	If SetsArray(3) <> "" Then
		If InStr(LCase(SetsArray(3)),"{url}") = 0 Then SetsArray(3) = SetsArray(2) & SetsArray(3)
		SetsArray(3) = strReplace(SetsArray(3),"{url}",SetsArray(2))
		SetsArray(3) = strReplace(SetsArray(3),"{appid}",SetsArray(1))
		SetsArray(3) = strReplace(SetsArray(3),"{text}",srcStr)
		SetsArray(3) = strReplace(SetsArray(3),"{from}",LangFrom)
		SetsArray(3) = strReplace(SetsArray(3),"{to}",LangTo)
	Else
		SetsArray(3) = SetsArray(2)
	End If

	If BodyData <> "" Then
		BodyData = strReplace(BodyData,"{url}",SetsArray(2))
		BodyData = strReplace(BodyData,"{appid}",SetsArray(1))
		BodyData = strReplace(BodyData,"{text}",srcStr)
		BodyData = strReplace(BodyData,"{from}",LangFrom)
		BodyData = strReplace(BodyData,"{to}",LangTo)
	End If

    xmlHttp.Open SetsArray(4),SetsArray(3),SetsArray(5),SetsArray(6),SetsArray(7)
    If fType = 3 Then xmlHttp.setRequestHeader("If-Modified-Since","0")
    If SetsArray(9) <> "" And UCase(SetsArray(4)) <> "GET" Then
		TempList = ReSplit(SetsArray(9),vbCrLf)
		For i = 0 To UBound(TempList)
			Temp = TempList(i)
    		Pos = InStr(Temp,",")
    		If Pos = 0 Then Pos = InStr(Temp,":")
			If Pos > 0 Then
				bstrHeader = Trim(Left(Temp,Pos-1))
				bstrValue = Trim(Mid(Temp,Pos+1))
				If bstrValue <> "" Then
					bstrValue = strReplace(bstrValue,"{url}",SetsArray(2))
					bstrValue = strReplace(bstrValue,"{appid}",SetsArray(1))
					bstrValue = strReplace(bstrValue,"{text}",srcStr)
					bstrValue = strReplace(bstrValue,"{from}",LangFrom)
					bstrValue = strReplace(bstrValue,"{to}",LangTo)
					If LCase(bstrHeader) = "content-length" Then
						xmlHttp.setRequestHeader bstrHeader,LenB(bstrValue)
					Else
						xmlHttp.setRequestHeader bstrHeader,bstrValue
					End If
				End If
			End If
		Next i
	End If
   	xmlHttp.send(BodyData)
   	i = OnReadyStateChange(xmlHttp,4,WaitTimes)
   	If i < 4 Then
   		If fType = 3 Then
			getTranslate = googleTranslator(srcStr, LngPair, LngJoinStr, Code)
			If getTranslate = "nodegoogle_error" Then getTranslate = IIf(i <= 1,"NotConnected","Timeout")
   		End If
   		xmlHttp.Abort
   		Exit Function
   	End If
	If xmlHttp.Status = 200 Or xmlHttp.Status = 206 Then
		If fType = 1 Then
			getTranslate = BytesToBstr(xmlHttp.responseBody,Code)
		Else
			Select Case responseType
			Case "responsetext"
				getTranslate = ExtractStr(xmlHttp.responseText,TranBeforeStr,TranAfterStr,0)
			Case "responsexml"
				getTranslate = ReadXML(xmlHttp.responseXML,TranBeforeStr,TranAfterStr)
			Case "responsestream"
				getTranslate = ExtractStr(BytesToBstr(xmlHttp.responseStream,Code),TranBeforeStr,TranAfterStr,0)
			Case "responsebody"
				getTranslate = ExtractStr(BytesToBstr(xmlHttp.responseBody,Code),TranBeforeStr,TranAfterStr,0)
			End Select
			If getTranslate = "" Then
				getTranslate = googleTranslator(nodeSrcStr, LngPair, LngJoinStr, Code)
				If getTranslate = "nodegoogle_error" Then getTranslate = ""
			End If
			getTranslate = Convert(getTranslate)
		End If
	End If
	xmlHttp.Abort
	On Error GoTo 0
	Exit Function

    ErrorHandler:
    If Err.Number <> 0 Then
    	If fType = 3 Then getTranslate = "NotConnected"
	End If
	xmlHttp.Abort
End Function

Function googleTranslator(ByRef srcStr As String, ByRef LngPair As String, ByRef LngJoinStr As String, ByRef Code As String) As String

	'Const WshRunning = 0
	'Const WshFinished = 1
	'Const WshFailed = 2

	Dim LangFrom, LangTo

	If LngPair <> "" Then
		TempList = ReSplit(LngPair,LngJoinStr)
		LangFrom = TempList(0)
		LangTo = TempList(1)
	End If

	Dim WshShell, oExec, nodeCmd, objStream
	Set objStream = CreateObject("ADODB.Stream")
	objStream.CharSet = "utf-8"
	objStream.Mode = 3
	objStream.Open
	objStream.WriteText srcStr
	objStream.SaveToFile("d:\APPs\Tools\npm_tools\translate.txt", 2)
	objStream.Close

	Set WshShell = CreateObject("WScript.Shell")
	nodeCmd = "node.exe " & "d:\APPs\Tools\npm_tools\index.js" & " --text=d:\APPs\Tools\npm_tools\translate.txt" & " --from=" & LangFrom &  " --to=" & LangTo
	'Set oExec = WshShell.Exec(nodeCmd)
'
	'Do While oExec.Status = WshRunning
		'Wait 1
	'Loop
'
	'If oExec.Status = WshFailed Then
	'googleTranslator = oExec.StdErr.ReadAll
	'Else
	'googleTranslator = oExec.StdOut.ReadAll
	'End If
	Set oExec = WshShell.run(nodeCmd, 0, True)

	objStream.CharSet = "utf-8"
	objStream.Open
	objStream.LoadFromFile("d:\APPs\Tools\npm_tools\translate.txt")
	googleTranslator = objStream.ReadText

	objStream.Close
	Set objStream = Nothing
	Set WshShell = Nothing
	Replace(googleTranslator, "％", "%")
	ConvStr(googleTranslator, "utf-8", Code)

End Function


'不区分大小写的字符替换 (保留未替换字符的大小写)
Function strReplace(ByVal s As String,ByVal Find As String,ByVal repwith As String) As String
	Dim i As Long,fL As Long
	strReplace = s
	If s = "" Or Find = "" Then Exit Function
	s = LCase$(s)
	Find = LCase$(Find)
	i = InStr(s,Find)
	If i = 0 Then Exit Function
	fL = Len(Find)
	Do While i > 0
		strReplace = Replace$(strReplace,Mid$(strReplace,i,fL),repwith)
		i = InStr(i + fL,s,Find)
	Loop
End Function


'转换宽字符为多字节
Function UTF16ToMultiByte(ByVal UTF16 As String, ByVal CodePage As Long) As Byte()
	Dim bufSize As Long,lRet As Long
	On Error GoTo errHandle
	ReDim arr(0) As Byte
	bufSize = WideCharToMultiByte(CodePage, 0&, StrPtr(UTF16), Len(UTF16), arr(0), 0, 0, 0)
	'If CodePage = CP_UTF8 Then bufSize = 2 * LenB(UTF16) - 1 Else bufSize = LenB(UTF16)
	If bufSize < 1 Then bufSize = 1
	ReDim arr(bufSize - 1) As Byte
	lRet = WideCharToMultiByte(CodePage, 0&, StrPtr(UTF16), Len(UTF16), arr(0), bufSize, 0, 0)
	If lRet > 0 Then
		ReDim Preserve arr(lRet - 1) As Byte
	End If
	UTF16ToMultiByte = arr
	Exit Function
	errHandle:
	ReDim arr(0) As Byte
	UTF16ToMultiByte = arr
End Function


'字符串转字节数组
Function StringToByte(ByVal textStr As String,ByVal CodePage As Long) As Byte()
	If textStr = "" Then
		ReDim StringToByte(0) As Byte
		Exit Function
	End If
	Select Case CodePage
	Case CP_UNICODELITTLE
		StringToByte = textStr
	Case CP_UNICODEBIG
		StringToByte = textStr
		StringToByte = LowByte2HighByte(StringToByte,0,-1)
	Case Else
		StringToByte = UTF16ToMultiByte(textStr,CodePage)	'按指定代码页转Unicode字符为ANSI组数
		'StringToByte = StrConv$(textStr,vbFromUnicode)		'按本机代码页转Unicode字符为ANSI组数
	End Select
End Function


'字符型字节数组的高字节和低字节互换
'适用于 UNICODE LITTLE 和 UNICODE BIG 字节数组的相互转换
Function LowByte2HighByte(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long) As Byte()
	Dim i As Long,Temp() As Byte
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Temp = Bytes
	For i = StartPos To EndPos - 1 Step 2
		Temp(i) = Bytes(i + 1)
		Temp(i + 1) = Bytes(i)
	Next i
	LowByte2HighByte = Temp
End Function


'字符串转 URL 转义符
Function Str2URLEsc(ByVal textStr As String,ByVal CodePage As Long,ByVal MultibyteOnly As Long) As String
	If textStr = "" Then Exit Function
	If MultibyteOnly = 0 Then
		Str2URLEsc = Byte2URLEsc(StringToByte(textStr,CodePage),0,-1,CodePage)
		Exit Function
	End If
	Dim i As Long,Matches As Object
	Str2URLEsc = textStr
	RegExp.Global = True
	RegExp.IgnoreCase = False
	'RegExp.Pattern = "[\x00-x20\x22\x25\x3C\x3E\x5B-\x5E\x60\x7B-\x7D\x7F]+"
	RegExp.Pattern = "[^\x21\x23\x24\x26-\x3B\x3D\x3F-\x5A\x5F\x61-\x7A\x7E]+"
	Set Matches = RegExp.Execute(textStr)
	If Matches.Count = 0 Then Exit Function
	For i = Matches.Count - 1 To 0 Step -1
		With Matches(i)
			If .FirstIndex > 0 Then
				Str2URLEsc = Left$(Str2URLEsc,.FirstIndex) & _
							Byte2URLEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2URLEsc,.FirstIndex + .Length + 1)
			Else
				Str2URLEsc = Byte2URLEsc(StringToByte(.Value,CodePage),0,-1,CodePage) & _
							Mid$(Str2URLEsc,.Length + 1)
			End If
		End With
	Next i
End Function


'字节转 RUL 转义符
'StartPos <= EndPos 获取低位到高位的 Hex 代码，否则获取高位到低位的 Hex 代码
Function Byte2URLEsc(Bytes() As Byte,ByVal StartPos As Long,ByVal EndPos As Long,ByVal CodePage As Long) As String
	Dim i As Long,n As Long
	If StartPos < 0 Then StartPos = LBound(Bytes)
	If EndPos < 0 Then EndPos = UBound(Bytes)
	Select Case CodePage
	Case CP_UNICODELITTLE
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			Mid$(Byte2URLEsc,n,6) = "%u" & Right$("0" & Hex$(Bytes(i + 1)),2) & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 6
		Next i
	Case CP_UNICODEBIG
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos - 1 Step IIf(StartPos <= EndPos,2,-2)
			Mid$(Byte2URLEsc,n,6) = "%u" & Right$("0" & Hex$(Bytes(i)),2) & Right$("0" & Hex$(Bytes(i + 1)),2)
			n = n + 6
		Next i
	Case Else
		Byte2URLEsc = Space$((Abs(EndPos - StartPos) + 1) * 3)
		n = 1
		For i = StartPos To EndPos Step IIf(StartPos <= EndPos,1,-1)
			Mid$(Byte2URLEsc,n,3) = "%" & Right$("0" & Hex$(Bytes(i)),2)
			n = n + 3
		Next i
	End Select
End Function


'转换字符的编码格式
Function ConvStr(ByVal textStr As String,ByVal inCode As String,ByVal outCode As String) As String
	Dim objStream As Object
    ConvStr = textStr
    If Trim(textStr) = "" Or inCode = "" Or outCode = "" Then Exit Function
    On Error GoTo ErrorMsg
    Set objStream = CreateObject("Adodb.Stream")
    If Not objStream Is Nothing Then
	    With objStream
    		.Type = 2
    		.Mode = 3
    		.CharSet = inCode
    		.Open
    		.WriteText textStr
    		.Position = 0
    		.CharSet = outCode
    		ConvStr = .ReadText
    		.Close
    	End With
		Set objStream = Nothing
	End If
    Exit Function
    ErrorMsg:
    Err.Source = "Adodb.Stream"
    Call sysErrorMassage(Err,1)
End Function


'解析 XML 格式对象并提取翻译文本
Function ReadXML(xmlObj As Object,ByVal IdName As String,ByVal TagName As String) As String
	If xmlObj Is Nothing Then Exit Function
	If IdName = "" And TagName = "" Then Exit Function

	Dim i As Long,j As Long,xmlDoc As Object,Item As Object
	On Error GoTo ErrorMsg
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	'Set xmldoc = CreateObject("Msxml2.DOMDocument")
	If xmlDoc Is Nothing Then Exit Function
	On Error Resume Next
	With xmlDoc
		'.async = False
		'.validateOnParse = False
		'.loadXML(xmlObj)	'加载字串
		.load(xmlObj)		'加载对象
		If .readyState > 2 Then
			If IdName <> "" Then
				TempList = ReSplit(IdName,"|")
				For i = 0 To UBound(TempList)
					Set Item = .nodeFromID(TempList(i))
					If Not Item Is Nothing Then
						ReadXML = Item.Text
						If ReadXML <> "" Then
							Set xmlDoc = Nothing
							Exit Function
						End If
					End If
				Next i
			End If
			If TagName <> "" Then
				TempList = ReSplit(TagName,"|")
				For i = 0 To UBound(TempList)
					Set Item = .getElementsByTagName(TempList(i))
					If Item.Length > 0 Then
						ReDim TempList(Item.Length - 1) As String
						For j = 0 To Item.Length - 1
							TempList(j) = Item(j).Text
						Next j
						ReadXML = Join$(TempList,"")
						If ReadXML <> "" Then Exit For
					End If
				Next i
			End If
		End If
	End With
	On Error GoTo 0
	Set xmlDoc = Nothing
	Exit Function

	ErrorMsg:
	Err.Source = "Microsoft.XMLDOM"
	Call sysErrorMassage(Err,1)
End Function


'提取指定前后字符之间的值
Function ExtractStr(ByVal textStr As String,ByVal BeforeStr As String,ByVal AfterStr As String,ByVal fType As Long) As String
	Dim i As Long,j As Long,k As Long,L1 As Long,L2 As Long
	If Trim(textStr) = "" Or (BeforeStr = "" And AfterStr = "") Then Exit Function
	textStr = textStr & vbCrLf
	BeforeStrArray = ReSplit(BeforeStr,"|")
	AfterStrArray = ReSplit(AfterStr,"|")
	k = UBound(AfterStrArray)
	For i = 0 To UBound(BeforeStrArray)
		For j = 0 To k
			BeforeStr = BeforeStrArray(i)
			AfterStr = AfterStrArray(j)
			L1 = InStr(textStr,BeforeStr)
			Do While L1 > 0
				L1 = L1 + Len(BeforeStr)
				L2 = InStr(L1,textStr,AfterStr)
				If fType > 0 And L2 = 0 Then L2 = InStr(L1,textStr,vbCrLf)
				If L2 > 0 Then
					If ExtractStr <> "" Then
						ExtractStr = ExtractStr & Mid(textStr,L1,L2 - L1)
					ElseIf ExtractStr = "" Then
						ExtractStr = Mid(textStr,L1,L2 - L1)
					End If
					If fType > 0 Then Exit Function
				End If
				L1 = InStr(L1,textStr,BeforeStr)
			Loop
			If ExtractStr <> "" Then Exit Function
		Next j
	Next i
End Function


'检查字串是否包含指定字符(文本和通配符比较)
'Mode = 0 检查字串是否包含指定字符，并找出指定字符的位置
'Mode = 1 检查字串是否只包含指定字符
'Mode = 2 检查字串是否包含指定字符
'Mode = 3 检查字串是否只包含大小混写的指定字符，此时 IgnoreCase 参数无效
'Mode = 4 检查字串中是否有连续相同的字符，StrNum 为检查的字符个数
'StrRange  定义字串检查范围 (可用 [Min - Max|Min - Max] 表示范围)
Public Function CheckStr(ByVal textStr As String,ByVal StrRange As String,Optional ByVal StrNum As Long, _
				Optional ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
	Dim i As Long,Temp As String
	If StrRange = "" Then Exit Function
	If Trim$(textStr) = "" Then Exit Function
	Select Case Mode
	Case 0
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		StrRange = Replace$(StrRange,"]|[","")
		If (textStr Like "*" & StrRange & "*") = False Then Exit Function
		For i = 1 To Len(textStr)
			If (Mid$(textStr,i,1) Like StrRange) = True Then
				CheckStr = i
				Exit For
			End If
		Next i
	Case 1
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		If (textStr Like "*[!" & Replace$(Replace$(StrRange,"]|[",""),"[","") & "*") = False Then
			CheckStr = True
		End If
	Case 2
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		If (textStr Like "*[" & Replace$(Replace$(StrRange,"]|[",""),"[","") & "*") = True Then
			CheckStr = True
		End If
	Case 3
		If InStr(textStr," ") Then Exit Function
		If Len(textStr) < 2 Then Exit Function
		If LCase$(textStr) = textStr Then Exit Function
		If UCase$(textStr) = textStr Then Exit Function
		If (textStr Like "*[!" & Replace$(Replace$(StrRange,"]|[",""),"[","") & "*") = True Then Exit Function
		If (Mid$(textStr,2) Like "*" & ReSplit(StrRange,"|")(0) & "*") = True Then
			If (Mid$(textStr,2) Like "*" & ReSplit(StrRange,"|")(1) & "*") = True Then
				CheckStr = True
			End If
		End If
	Case 4
		If StrNum < 2 Then Exit Function
		If Len(textStr) < StrNum Then Exit Function
		If InStr(textStr," ") Then Exit Function
		If IsNumeric(textStr) = True Then Exit Function
		If IsDate(textStr) = True Then Exit Function
		If IgnoreCase = True Then
			textStr = LCase$(textStr)
			StrRange = LCase$(StrRange)
		End If
		If (textStr Like "*" & Replace$(StrRange,"]|[","") & "*") = False Then Exit Function
		For i = 1 To Len(textStr)
			Temp = Mid$(textStr,i,1)
			If IsNumeric(Temp) = False Then
				If Temp <> "." Then
					If InStr(textStr,String$(StrNum,Temp)) Then
						CheckStr = True
						Exit For
					End If
				End If
			End If
		Next i
	End Select
End Function


'检查字串是否包含指定字符(正则表达式比较)
'Mode = 0 检查字串是否包含指定字符，并找出指定字符的位置
'Mode = 1 检查字串是否只包含指定字符
'Mode = 2 检查字串是否包含指定字符
'Mode = 3 检查字串是否只包含大小混写的指定字符，此时 IgnoreCase 参数无效
'Mode = 4 检查字串中是否有连续相同的字符，StrNum 为最少重复字符个数
'Patrn  为正则表达式模板
Public Function CheckStrRegExp(ByVal textStr As String,ByVal Patrn As String,Optional ByVal StrNum As Long, _
				Optional ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
	Dim n As Long,Matches As Object
	If Patrn = "" Then Exit Function
	If Trim$(textStr) = "" Then Exit Function
	With RegExp
		Select Case Mode
		Case 0
			.Global = True
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count > 0 Then CheckStrRegExp = Matches(0).FirstIndex + 1
		Case 1
			.Global = True
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count = Len(textStr) Then CheckStrRegExp = True
		Case 2
			.Global = False
			.IgnoreCase = IgnoreCase
			.Pattern = Patrn
			If .Test(textStr) Then CheckStrRegExp = True
		Case 3
			If InStr(textStr," ") Then Exit Function
			n = Len(textStr)
			If n < 2 Then Exit Function
			If LCase$(textStr) = textStr Then Exit Function
			If UCase$(textStr) = textStr Then Exit Function
			.Global = True
			.IgnoreCase = False
			.Pattern = Patrn
			Set Matches = .Execute(textStr)
			If Matches.Count <> n Then Exit Function
			.Pattern = ReSplit(Patrn,"|")(1)
			Set Matches = .Execute(Mid$(textStr,2))
			If Matches.Count <> n - 1 Then CheckStrRegExp = True
		Case 4
			If StrNum < 2 Then Exit Function
			If Len(textStr) < StrNum Then Exit Function
			If InStr(textStr," ") Then Exit Function
			.Global = False
			.IgnoreCase = IgnoreCase
			.Pattern = "(" & Patrn & ")\1{" & StrNum - 1 & ",}"
			If .Test(textStr) Then CheckStrRegExp = True
		End Select
	End With
End Function


'分行翻译处理
Function SplitTran(xmlHttp As Object,EngineData() As String,srcStr As String,ByVal LangPair As String,ByVal EngineID As Long, _
	ByVal CheckID As Long,ByVal ProjectIDSrc As Long,ByVal mCheckSrc As Long,ByVal mHanding As Long,ByVal fType As Long) As String
	Dim i As Long,srcStrBak As String,trnString As String,Temp As String,Stemp As Boolean

	If Trim(srcStr) = "" Or LangPair = "" Then Exit Function

	'用替换法拆分字串
	srcStrBak = srcStr
	LineSplitChar = "\r\n,\r,\n"
	FindStrArr = ReSplit(Convert(LineSplitChar),",",-1)
	For i = 0 To UBound(FindStrArr)
		FindStr = Trim(FindStrArr(i))
		If InStr(srcStrBak,FindStr) Then
			srcStrBak = Replace(srcStrBak,FindStr,"*c!N!g*")
		End If
	Next i
	srcStrArr = ReSplit(srcStrBak,"*c!N!g*",-1)

	'获取每行的翻译
	Temp = srcStr
	Stemp = False
	For i = 0 To UBound(srcStrArr)
		If srcStrArr(i) <> "" Then
			srcStrBak = srcStrArr(i)
			If mHanding = 0 Then
				If mCheckSrc = 1 Then srcStrArr(i) = CheckHanding(CheckID,srcStrBak,srcStrArr(i),ProjectIDSrc)
				If InStr(srcStrArr(i),"&") Then srcStrArr(i) = Replace(srcStrArr(i),"&","")
				If srcStrArr(i) <> "" And srcStrArr(i) <> srcStrBak Then
					srcStr = Replace(srcStr,srcStrBak,srcStrArr(i),,1)
				End If
			End If
			trnString = getTranslate(xmlHttp,EngineData,EngineID,srcStrArr(i),LangPair,fType)
			If trnString <> "" And trnString <> srcStrBak Then
				Temp = Replace(Temp,srcStrBak,trnString,,1)
				Stemp = True
			End If
		End If
	Next i
	If Stemp = True Then SplitTran = Temp
End Function


'替换特定字符
'fType = 0 正向替换，使用第一个替换字符配置
'fType = 1 还原替换，使用第一个替换字符配置
'fType = 2 正向替换，使用第二个替换字符配置
'fType = 3 还原替换，使用第二个替换字符配置
'Record = 0 不记录替换字符
'Record = 1 记录替换字符
Function ReplaceStr(ByVal CheckID As Long,ByVal trnStr As String,ByVal fType As Long,ByVal Record As Long) As String
	Dim i As Long,PreStr As String,AppStr As String
	ReplaceStr = trnStr
	PreRepStr = ""
	AppRepStr = ""
	If Trim(trnStr) = "" Then Exit Function
	'获取选定配置的参数
	TempArray = ReSplit(CheckDataListBak(CheckID),JoinStr)
	SetsArray = ReSplit(TempArray(1),SubJoinStr)
	If fType < 2 Then AutoRepChar = SetsArray(11) Else AutoRepChar = SetsArray(12)
	If AutoRepChar <> "" Then
		FindStrArr = ReSplit(AutoRepChar,",",-1)
		For i = 0 To UBound(FindStrArr)
			FindStr = FindStrArr(i)
			PreStr = ""
			AppStr = ""
			If InStr(FindStr,"|") Then
				TempArray = ReSplit(FindStr,"|")
				If fType = 0 Or fType = 2 Then
					PreStr = TempArray(0)
					AppStr = TempArray(1)
				Else
					PreStr = TempArray(1)
					AppStr = TempArray(0)
				End If
				cPreStr = Convert(PreStr)
				cAppStr = Convert(AppStr)
			End If
			If PreStr <> "" And InStr(ReplaceStr,cPreStr) Then
				ReplaceStr = Replace(ReplaceStr,cPreStr,cAppStr)
				If Record = 1 Then
					If PreRepStr <> "" Then
						If InStr(PreRepStr,PreStr) = 0 Then PreRepStr = PreRepStr & JoinStr & PreStr
					Else
						PreRepStr = PreStr
					End If
					If AppRepStr <> "" Then
						If InStr(AppRepStr,AppStr) = 0 Then AppRepStr = AppRepStr & JoinStr & AppStr
					Else
						AppRepStr = AppStr
					End If
				End If
			End If
		Next i
	End If
End Function


'检查修正快捷键、终止符和加速器
Function CheckHanding(ByVal CheckID As Long,ByVal srcStr As String,ByVal trnStr As String,ByVal ProjectID As Long) As String
	Dim i As Long,srcStrBak As String,trnStrBak As String,LineSplitMode As Long
	Dim srcNum As Long,trnNum As Long,srcSplitNum As Long,trnSplitNum As Long,Stemp As Boolean
	Dim FindStr As String,srcStrArr() As String,trnStrArr() As String,TempArray() As String
	Dim k As Long,l As Long,m As Long

	'参数初始化
	srcNum = 0
	trnNum = 0
	srcSplitNum = 0
	trnSplitNum = 0
	CheckHanding = trnStr
	If Trim(srcStr) = "" Or Trim(trnStr) = "" Then Exit Function

	'获取选定配置的参数
	TempArray = ReSplit(CheckDataList(CheckID),JoinStr)
	SetsArray = ReSplit(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	PreInsertSplitChar = SetsArray(1)
	KeepCharPair = SetsArray(3)
	AccessKeyChar = SetsArray(13)
	LineSplitMode = StrToLong(SetsArray(15))
	AppInsertSplitChar = SetsArray(16)
	ReplaceSplitChar = SetsArray(17)

	TempArray = ReSplit(ProjectDataList(ProjectID),JoinStr)
	SetsArray = ReSplit(TempArray(1),LngJoinStr)
	EnableStringSplit = StrToLong(SetsArray(16))

	'配置参数数组化
	If ExcludeChar <> "" Then ExcludeCharArr = ReSplit(ExcludeChar,",",-1)
	If KeepCharPair <> "" Then KeepCharPairArr = ReSplit(KeepCharPair,",",-1)
	If AccessKeyChar <> "" Then AccessKeyCharArr = ReSplit(AccessKeyChar,",",-1)

	If EnableStringSplit = 1 Then
		LineSplitChar = PreInsertSplitChar & AppInsertSplitChar & ReplaceSplitChar
		Temp = PreInsertSplitChar & "," & AppInsertSplitChar & "," & ReplaceSplitChar
		If LineSplitChar <> "" Then LineSplitCharArr = ReSplit(Temp,",",-1)
		If PreInsertSplitChar <> "" Then k = UBound(ReSplit(PreInsertSplitChar,",",-1)) + 1
		If AppInsertSplitChar <> "" Then l = UBound(ReSplit(AppInsertSplitChar,",",-1)) + 1
		If ReplaceSplitChar <> "" Then m = UBound(ReSplit(ReplaceSplitChar,",",-1)) + 1
	End If

	'排除字串中的非快捷键
	If ExcludeChar <> "" Then
		For i = 0 To UBound(ExcludeCharArr)
			FindStr = LTrim(ExcludeCharArr(i))
			If FindStr <> "" Then
				srcStr = Replace(srcStr,FindStr,"*a" & i & "!N!" & i & "d*")
				trnStr = Replace(trnStr,FindStr,"*a" & i & "!N!" & i & "d*")
			End If
		Next i
	End If

	'过滤不是快捷键的快捷键
	If KeepCharPair <> "" Then
		For i = 0 To UBound(KeepCharPairArr)
			FindStr = Trim(KeepCharPairArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				ToRepStr = LFindStr & "&" & RFindStr
				BeRepStr = LFindStr & "*!N!" & i & "!M!" & i & "!N!*" & RFindStr
				srcStr = Replace(srcStr,ToRepStr,BeRepStr)
				trnStr = Replace(trnStr,ToRepStr,BeRepStr)
			End If
		Next i
	End If

	'用替换法拆分字串
	If EnableStringSplit = 1 Then
		srcStrBak = srcStr
		trnStrBak = trnStr
		If LineSplitChar <> "" Then
			For i = 0 To UBound(LineSplitCharArr)
				FindStr = Trim(LineSplitCharArr(i))
				If FindStr <> "" Then
					Stemp = False
					If LineSplitMode = 1 Then
						srcNum = UBound(ReSplit(srcStrBak,FindStr,-1))
						trnNum = UBound(ReSplit(trnStrBak,FindStr,-1))
						If srcNum = trnNum And srcNum > 0 And trnNum > 0 Then Stemp = True
					End If
					If LineSplitMode = 0 Or Stemp = True Then
						If InStr(LCase(AccessKeyChar),LCase(FindStr)) Then
							srcStrBak = Insert(srcStrBak,FindStr,"*c!N!g*",1)
							trnStrBak = Insert(trnStrBak,FindStr,"*c!N!g*",1)
						ElseIf i < k And k <> 0 Then
							srcStrBak = Replace(srcStrBak,FindStr,"*c!N!g*" & FindStr)
							trnStrBak = Replace(trnStrBak,FindStr,"*c!N!g*" & FindStr)
						ElseIf i >= k And i < k + l + 1 And l <> 0 Then
							srcStrBak = Replace(srcStrBak,FindStr,FindStr & "*c!N!g*")
							trnStrBak = Replace(trnStrBak,FindStr,FindStr & "*c!N!g*")
						ElseIf i >= k + l And i < k + l + m + 2 And m <> 0 Then
							srcStrBak = Replace(srcStrBak,FindStr,"*c!N!g*")
							trnStrBak = Replace(trnStrBak,FindStr,"*c!N!g*")
						End If
					End If
				End If
			Next i
		End If
		srcStrArr = ReSplit(srcStrBak,"*c!N!g*",-1)
		trnStrArr = ReSplit(trnStrBak,"*c!N!g*",-1)

		'字串处理
		Stemp = False
		srcNum = UBound(srcStrArr)
		trnNum = UBound(trnStrArr)
		If srcNum > 0 And trnNum > 0 Then
			If LineSplitMode = 0 Then Stemp = True
			If LineSplitMode = 1 And srcNum = trnNum Then Stemp = True
		End If
		If Stemp = True Then
			TempArray = MergeArray(srcStrArr,trnStrArr)
			trnStr = ReplaceStrSplit(CheckID,trnStr,TempArray,ProjectID)
		Else
			trnStr = StringReplace(CheckID,srcStr,trnStr,ProjectID)
		End If
	Else
		trnStr = StringReplace(CheckID,srcStr,trnStr,ProjectID)
	End If

	'计算快捷键数
	srcStrBak = srcStr
	trnStrBak = trnStr
	toRepStr = Trim(AccessKeyCharArr(0))
	If AccessKeyChar <> "" Then
		For i = 0 To UBound(AccessKeyCharArr)
			FindStr = Trim(AccessKeyCharArr(i))
			If FindStr <> "" And FindStr <> toRepStr Then
				srcStrBak = Replace(srcStrBak,FindStr,toRepStr)
				trnStrBak = Replace(trnStrBak,FindStr,toRepStr)
			End If
		Next i
	End If
	StringSrc.AccKeyNum = UBound(ReSplit(srcStrBak,toRepStr,-1))
	StringTrn.AccKeyNum = UBound(ReSplit(trnStrBak,toRepStr,-1))

	'还原不是快捷键的快捷键
	If KeepCharPair <> "" Then
		For i = 0 To UBound(KeepCharPairArr)
			FindStr = Trim(KeepCharPairArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				ToRepStr = LFindStr & "*!N!" & i & "!M!" & i & "!N!*" & RFindStr
				BeRepStr = LFindStr & "&" & RFindStr
				srcStr = Replace(srcStr,ToRepStr,BeRepStr)
				trnStr = Replace(trnStr,ToRepStr,BeRepStr)
			End If
		Next i
	End If

	'还原字串中被排除的非快捷键
	If ExcludeChar <> "" Then
		For i = 0 To UBound(ExcludeCharArr)
			FindStr = LTrim(ExcludeCharArr(i))
			If FindStr <> "" Then
				srcStr = Replace(srcStr,"*a" & i & "!N!" & i & "d*",FindStr)
				trnStr = Replace(trnStr,"*a" & i & "!N!" & i & "d*",FindStr)
			End If
		Next i
	End If
	CheckHanding = trnStr
End Function


'在快捷键后插入特定字符并以此拆分字串
Function Insert(ByVal SplitString As String,ByVal SplitStr As String,ByVal InsStr As String,ByVal Leng As Long) As String
	Dim i As Long,j As Long
	Insert = SplitString
	If UBound(ReSplit(SplitString,SplitStr)) < 2 Then Exit Function
	i = InStr(Insert,SplitStr)
	Do While i > Leng
		j = InStr(i + 1,Insert,SplitStr)
		If j > i Then
			SplitString = Mid(Insert,i - Leng,j - i)
			If SplitString <> "" Then
				Insert = Replace(Insert,SplitString,SplitString & InsStr)
			End If
		End If
		i = InStr(i + 1,Insert,SplitStr)
	Loop
	'PSL.Output "Insert = " & Insert       '调试用
End Function


'读取数组中的每个字串并替换处理
Function ReplaceStrSplit(ByVal CheckID As Long,ByVal trnStr As String,StrSplitArr() As String,ByVal ProjectID As Long) As String
	Dim Temp As String,i As Long,j As Long

	j = 1
	ReplaceStrSplit = trnStr
	For i = 0 To UBound(StrSplitArr) Step 2
		Temp = StringReplace(CheckID,StrSplitArr(i),StrSplitArr(i + 1),ProjectID)

		'处理在前后行中包含的重复字符
		If StrSplitArr(i + 1) <> Temp Then
			If j = 1 Then
				ReplaceStrSplit = Replace(ReplaceStrSplit,StrSplitArr(i + 1),Temp,j,1)
			Else
				ReplaceStrSplit = Left(ReplaceStrSplit,j - 1) & Replace(ReplaceStrSplit,StrSplitArr(i + 1),Temp,j,1)
			End If
		End If
		j = j + Len(Temp)

		'对每行的数据进行连接，用于消息输出
		TPreSpaceSrc = TPreSpaceSrc & StringSrc.PreSpace
		TPreSpaceTrn = TPreSpaceTrn & StringTrn.PreSpace
		TacckeySrc = TacckeySrc & StringSrc.AccKey
		TacckeyTrn = TacckeyTrn & StringTrn.AccKey
		TEndStringSrc = TEndStringSrc & StringSrc.EndString
		TEndStringTrn = TEndStringTrn & StringTrn.EndString
		TShortcutSrc = TShortcutSrc & StringSrc.Shortcut
		TShortcutTrn = TShortcutTrn & StringTrn.Shortcut
		TEndSpaceSrc = TEndSpaceSrc & StringSrc.EndSpace
		TEndSpaceTrn = TEndSpaceTrn & StringTrn.EndSpace
		TSpaceTrn = TSpaceTrn & StringTrn.Spaces
		TExpStringTrn = TExpStringTrn & StringTrn.ExpString
		TPreStringTrn = TPreStringTrn & StringTrn.PreString
		TMoveAcckey = TMoveAcckey & MoveAcckey
	Next i

	'为调用消息输出，用原有变量替换连接后的数据
	StringSrc.PreSpace = TPreSpaceSrc
	StringTrn.PreSpace = TPreSpaceTrn
	StringSrc.AccKey = TacckeySrc
	StringTrn.AccKey = TacckeyTrn
	StringSrc.EndString = TEndStringSrc
	StringTrn.EndString = TEndStringTrn
	StringSrc.Shortcut = TShortcutSrc
	StringTrn.Shortcut = TShortcutTrn
	StringSrc.EndSpace = TEndSpaceSrc
	StringTrn.EndSpace = TEndSpaceTrn
	StringTrn.Spaces = TSpaceTrn
	StringTrn.ExpString = TExpStringTrn
	StringTrn.PreString = TPreStringTrn
	MoveAcckey = TMoveAcckey
End Function


'按行获取字串的各个字段并替换翻译字符串
Function StringReplace(ByVal CheckID As Long,ByVal srcStr As String,ByVal trnStr As String,ByVal ProjectID As Long) As String
	Dim i As Long,j As Long,x As Long,y As Long,m As Long,n As Long
	Dim AsiaKey As Long,AddAccessKeyWithFirstChar As Long,LeadingSpaceInSource As Long
	Dim LeadingSpaceInTarget As Long,LeadingSpaceInBoth As Long,TrailingSpaceInSource As Long
	Dim TrailingSpaceInTarget As Long,TrailingSpaceInBoth As Long,AccessKeyInSource As Long
	Dim AccessKeyInTarget As Long,AccessKeyInBoth As Long,EndCharInSource As Long,EndCharInTarget As Long
	Dim EndCharInBoth As Long,ShortcutInSource As Long,ShortcutInTarget As Long,ShortcutInBoth As Long
	Dim DeleteExtraSpace As Long,TranslateEndChar As Long,AccKeyInShort As Long,TempArray() As String
	Dim FindStr As String,LastStringTrn As String,Temp As String,TempBak As String,Stemp As Boolean

	'参数初始化
	StringSrc.Length = 0
	StringTrn.Length = 0
	StringSrc.PreSpace = ""
	StringTrn.PreSpace = ""
	StringSrc.EndSpace = ""
	StringTrn.EndSpace = ""
	StringSrc.AccKey = ""
	StringTrn.AccKey = ""
	StringSrc.AccKeyIFR = ""
	StringTrn.AccKeyIFR = ""
	StringSrc.AccKeyKey = ""
	StringTrn.AccKeyKey = ""
	StringSrc.AccKeyPos = 0
	StringTrn.AccKeyPos = 0
	StringSrc.EndString = ""
	StringTrn.EndString = ""
	StringSrc.Shortcut = ""
	StringTrn.Shortcut = ""
	StringTrn.Spaces = ""
	StringTrn.ExpString = ""
	StringTrn.PreString = ""
	LastStringTrn = ""
	MoveAcckey = ""
	StringReplace = trnStr
	If Trim(srcStr) = "" Or Trim(trnStr) = "" Then Exit Function

	'获取选定配置的参数
	TempArray = ReSplit(CheckDataList(CheckID),JoinStr)
	SetsArray = ReSplit(TempArray(1),SubJoinStr)
	CheckBracket = SetsArray(2)
	AsiaKey = StrToLong(SetsArray(4))
	CheckEndChar = SetsArray(5)
	NoTrnEndChar = SetsArray(6)
	AutoTrnEndChar = SetsArray(7)
	CheckShortChar = SetsArray(8)
	CheckShortKey = SetsArray(9)
	KeepShortKey = SetsArray(10)
	AccessKeyChar = SetsArray(13)
	AddAccessKeyWithFirstChar = StrToLong(SetsArray(14))

	TempArray = ReSplit(ProjectDataList(ProjectID),JoinStr)
	SetsArray = ReSplit(TempArray(1),LngJoinStr)
	LeadingSpaceInSource = StrToLong(SetsArray(0))
	LeadingSpaceInTarget = StrToLong(SetsArray(1))
	LeadingSpaceInBoth = StrToLong(SetsArray(2))
	TrailingSpaceInSource = StrToLong(SetsArray(3))
	TrailingSpaceInTarget = StrToLong(SetsArray(4))
	TrailingSpaceInBoth = StrToLong(SetsArray(5))
	AccessKeyInSource = StrToLong(SetsArray(6))
	AccessKeyInTarget = StrToLong(SetsArray(7))
	AccessKeyInBoth = StrToLong(SetsArray(8))
	EndCharInSource = StrToLong(SetsArray(9))
	EndCharInTarget = StrToLong(SetsArray(10))
	EndCharInBoth = StrToLong(SetsArray(11))
	ShortcutInSource = StrToLong(SetsArray(12))
	ShortcutInTarget = StrToLong(SetsArray(13))
	ShortcutInBoth = StrToLong(SetsArray(14))
	DeleteExtraSpace = StrToLong(SetsArray(15))
	TranslateEndChar = StrToLong(SetsArray(17))
	AccKeyInShort = StrToLong(SetsArray(18))

	'配置参数数组化
	If CheckBracket <> "" Then CheckBracketArr = ReSplit(CheckBracket,",",-1)
	If CheckEndChar <> "" Then CheckEndCharArr = ReSplit(CheckEndChar," ",-1)
	If AutoTrnEndChar <> "" Then AutoTrnEndCharArr = ReSplit(AutoTrnEndChar," ",-1)
	If CheckShortChar <> "" Then CheckShortCharArr = ReSplit(CheckShortChar,",",-1)
	If AccessKeyChar <> "" Then AccessKeyCharArr = ReSplit(AccessKeyChar,",",-1)

	'获取来源和翻译的长度
	StringSrc.Length = Len(srcStr)
	StringTrn.Length = Len(trnStr)

	'获取来源和翻译的前置空格
	StringSrc.PreSpace = Space(StringSrc.Length - Len(LTrim(srcStr)))
	StringTrn.PreSpace = Space(StringTrn.Length - Len(LTrim(trnStr)))

	'获取来源和翻译的尾随空格
	StringSrc.EndSpace = Space(StringSrc.Length - Len(RTrim(srcStr)))
	StringTrn.EndSpace = Space(StringTrn.Length - Len(RTrim(trnStr)))

	'获取来源和翻译的加速器
	If CheckShortChar <> "" Then
		CheckShortKey = CheckShortKey & "," & KeepShortKey
		If AccessKeyChar <> "" Then m = UBound(AccessKeyCharArr)
		For i = 0 To UBound(CheckShortCharArr)
			FindStr = Trim(CheckShortCharArr(i))
			If FindStr <> "" Then
				For n = 0 To 1
					If n = 0 Then Temp = RTrim(srcStr)
					If n = 1 Then Temp = RTrim(trnStr)
					Shortcut = ""
					ShortcutKey = ""
					If InStrRev(LTrim(Temp),FindStr) > 1 Then
						y = InStrRev(Temp,FindStr)
						ShortcutKey = Trim(Mid(Temp,y + 1))
					End If
					If ShortcutKey <> "" Then
						If AccessKeyChar <> "" Then
							For j = 0 To m
								ShortcutKey =Replace(ShortcutKey,AccessKeyCharArr(j),"")
							Next j
						End If
						If ShortcutKey = "+" Then
							If CheckKeyCode(ShortcutKey,CheckShortKey) <> 0 Then
								Shortcut = Mid(Temp,y)
							End If
						ElseIf InStr(ShortcutKey,"+") Then
							x = 0
							TempArray = ReSplit(ShortcutKey,"+",-1)
							For j = 0 To UBound(TempArray)
								x = x + CheckKeyCode(TempArray(j),CheckShortKey)
							Next j
							If x > 0 And x >= UBound(TempArray) Then
								Shortcut = Mid(Temp,y)
							End If
						Else
							If CheckKeyCode(ShortcutKey,CheckShortKey) <> 0 Then
								Shortcut = Mid(Temp,y)
							End If
						End If
						If Shortcut <> "" Then
							If n = 0 And StringSrc.Shortcut = "" Then
								StringSrc.Shortcut = Shortcut
								ShortcutKeySrc = ShortcutKey
							ElseIf n = 1 And StringTrn.Shortcut = "" Then
								StringTrn.Shortcut = Shortcut
								ShortcutKeyTrn = ShortcutKey
							End If
						End If
					End If
				Next n
			End If
			If StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" Then Exit For
		Next i
	End If

	'获取来源和翻译的终止符及其前后空格
	If CheckEndChar <> "" Then
		xTemp = Left(srcStr,StringSrc.Length - Len(StringSrc.Shortcut & StringSrc.EndSpace))
		yTemp = Left(trnStr,StringTrn.Length - Len(StringTrn.Shortcut & StringTrn.EndSpace))
		If NoTrnEndChar <> "" Then
			If CheckKeyCode(xTemp,NoTrnEndChar) = 1 Then xTemp = ""
			If CheckKeyCode(yTemp,NoTrnEndChar) = 1 Then yTemp = ""
		End If
	End If
	If xTemp <> "" And yTemp <> "" Then
		If AccessKeyChar <> "" Then m = UBound(AccessKeyCharArr)
		For i = 0 To UBound(CheckEndCharArr)
			FindStr = Trim(CheckEndCharArr(i))
			If FindStr <> "" Then
				PreFindStr = Left(FindStr,1)
				AppFindStr = Right(FindStr,1)
				For j = 0 To 1
					If j = 0 Then
						Temp = xTemp
						EndSpace = StringSrc.EndSpace
						Shortcut = StringSrc.Shortcut
					Else
						Temp = yTemp
						EndSpace = StringTrn.EndSpace
						Shortcut = StringTrn.Shortcut
					End If
					n = 0
					y = InStrRev(Temp,FindStr)
					If y > 0 Then
						If Trim(Mid(Temp,y)) Like FindStr Then
							n = y
						ElseIf Right(Trim(Temp),1) = AppFindStr Then
							y = InStr(Temp,PreFindStr)
							Do While y > 0
								TempBak = Trim(Mid(Temp,y))
								If AccessKeyChar <> "" Then
									For x = 0 To m
										TempBak = Replace(TempBak,AccessKeyCharArr(x),"")
									Next x
								End If
								If TempBak Like FindStr Then
									n = y
									Exit Do
								End If
								y = InStr(y + 1,Temp,PreFindStr)
							Loop
						End If
					End If
					If n <> 0 Then
						PreStr = Left(Temp,n - 1)
						AppStr = Mid(Temp,n)
						x = Len(PreStr) - Len(RTrim(PreStr))
						If j = 0 And StringSrc.EndString = "" Then
							StringSrc.EndString = Space(x) & AppStr
						ElseIf j = 1 And StringTrn.EndString = "" Then
							StringTrn.EndString = Space(x) & AppStr
						End If
					End If
				Next j
			End If
			If StringSrc.EndString <> "" And StringTrn.EndString <> "" Then Exit For
		Next i
	End If

	'获取来源和翻译的快捷键位置及其字符
	If AccessKeyChar <> "" Then
		For i = 0 To UBound(AccessKeyCharArr)
			FindStr = Trim(AccessKeyCharArr(i))
			If FindStr <> "" Then
				For j = 0 To 1
					If j = 0 Then Temp = srcStr
					If j = 1 Then Temp = trnStr
					n = InStrRev(Temp,FindStr)
					If n > 0 Then
						If j = 0 And n > StringSrc.AccKeyPos Then
							StringSrc.AccKeyPos = n
							StringSrc.AccKeyIFR = FindStr
							StringSrc.AccKeyKey = Mid(Temp,n + Len(FindStr),1)
						ElseIf j = 1 And n > StringTrn.AccKeyPos Then
							StringTrn.AccKeyPos = n
							StringTrn.AccKeyIFR = FindStr
							StringTrn.AccKeyKey = Mid(Temp,n + Len(FindStr),1)
						End If
					End If
				Next j
			End If
		Next i
	End If
	If StringSrc.AccKeyIFR = "" Then StringSrc.AccKeyIFR = "&"
	If StringTrn.AccKeyIFR = "" Then StringTrn.AccKeyIFR = "&"

	'获取来源和翻译的快捷键 (包括快捷键前后的括号字符)
	If (StringSrc.AccKeyPos > 1 Or StringTrn.AccKeyPos > 1) And CheckBracket <> "" Then
		For i = 0 To UBound(CheckBracketArr)
			FindStr = Trim(CheckBracketArr(i))
			If FindStr <> "" Then
				PreFindStr = Trim(Left(FindStr,1))
				AppFindStr = Trim(Right(FindStr,1))
				For n = 0 To 1
					If n = 0 Then
						Temp = srcStr
						j = StringSrc.AccKeyPos
						xTemp = StringSrc.AccKeyIFR
						yTemp = StringSrc.AccKeyKey
					ElseIf n = 1 Then
						Temp = trnStr
						j = StringTrn.AccKeyPos
						xTemp = StringTrn.AccKeyIFR
						yTemp = StringTrn.AccKeyKey
					End If
					AccessKey = ""
					If j > 1 Then
						x = InStrRev(Temp,PreFindStr,j)
						y = InStr(j,Temp,AppFindStr)
						If x > 0 And y > x Then
							TempBak = Mid(Temp,x + 1,y - x - 1)
							If Trim(TempBak) = xTemp & yTemp Then
								AccessKey = Mid(Temp,x,y - x + 1)
								j = x
							End If
						End If
					ElseIf j = 1 Then
						AccessKey = xTemp
					End If
					If AccessKey <> "" Then
						If n = 0 And StringSrc.AccKey = "" Then
							StringSrc.AccKeyPos = j
							StringSrc.AccKey = AccessKey
						ElseIf n = 1 And StringTrn.AccKey = "" Then
							StringTrn.AccKeyPos = j
							StringTrn.AccKey = AccessKey
						End If
					End If
				Next n
			End If
			If StringSrc.AccKey <> "" And StringTrn.AccKey <> "" Then Exit For
		Next i
	End If
	If StringSrc.AccKey = "" And StringSrc.AccKeyPos > 0 Then StringSrc.AccKey = StringSrc.AccKeyIFR
	If StringTrn.AccKey = "" And StringTrn.AccKeyPos > 0 Then StringTrn.AccKey = StringTrn.AccKeyIFR

	'获取翻译的快捷键后面的非终止符和非加速器的字符(包括空格)
	If StringTrn.AccKeyPos > 0 Then
		x = Len(StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		'If InStr(StringTrn.Shortcut,StringTrn.AcckeyIFR) Then x = Len(StringTrn.EndSpace)
		'If InStr(StringTrn.EndString,StringTrn.AcckeyIFR) Then x = Len(StringTrn.Shortcut & StringTrn.EndSpace)
		If StringTrn.Length > x Then
			Temp = Left(trnStr,StringTrn.Length - x)
			StringTrn.ExpString = Mid(Temp,StringTrn.AccKeyPos + Len(StringTrn.AccKey))
		End If
	End If

	'获取翻译的快捷键或终止符或加速器前面的空格
	Temp = StringTrn.AccKey & StringTrn.ExpString & StringTrn.EndString & _
			StringTrn.Shortcut & StringTrn.EndSpace
	If Temp <> "" Then
		x = Len(Temp)
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			x = Len(StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		End If
		If StringTrn.Length > x Then
			Temp = Left(trnStr,StringTrn.Length - x)
			y = Len(Temp) - Len(RTrim(Temp))
			If y > 0 Then StringTrn.Spaces = Space(y)
		End If
	End If

	'获取翻译的快捷键前的终止符及其终止符前的空格
	If StringTrn.AccKey <> "" And CheckEndChar <> "" Then
		x = Len(StringTrn.Spaces & StringTrn.AccKey & StringTrn.ExpString & StringTrn.EndString & _
			StringTrn.Shortcut & StringTrn.EndSpace)
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			x = Len(StringTrn.Spaces & StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		End If
		Temp = Left(trnStr,StringTrn.Length - x)
		If Trim(Temp) <> "" Then
			If NoTrnEndChar <> "" Then
				If CheckKeyCode(Temp,NoTrnEndChar) = 1 Then Temp = ""
			End If
			If Temp <> "" Then
				For i = 0 To UBound(CheckEndCharArr)
					FindStr = Trim(CheckEndCharArr(i))
					PreFindStr = Left(FindStr,1)
					AppFindStr = Right(FindStr,1)
					n = 0
					y = InStrRev(Temp,FindStr)
					If y > 0 Then
						If Trim(Mid(Temp,y)) Like FindStr Then
							n = y
						ElseIf Right(Trim(Temp),1) = AppFindStr Then
							y = InStr(Temp,PreFindStr)
							Do While y > 0
								TempBak = Trim(Mid(Temp,y))
								If TempBak Like FindStr Then
									n = y
									Exit Do
								End If
								y = InStr(y + 1,Temp,PreFindStr)
							Loop
						End If
					End If
					If n > 0 Then
						PreStr = Left(Temp,n - 1)
						AppStr = Mid(Temp,n)
						x = Len(PreStr) - Len(RTrim(PreStr))
						StringTrn.PreString = Space(x) & AppStr
					End If
					If StringTrn.PreString <> "" Then Exit For
				Next i
			End If
		End If
	End If

	'获取翻译中除已提取字符外的其他所有字符
	Temp = StringTrn.PreString & StringTrn.Spaces & StringTrn.AccKey & StringTrn.ExpString & _
			StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace
	If Temp <> "" Then
		x = Len(Temp)
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			x = Len(StringTrn.PreString & StringTrn.Spaces & StringTrn.EndString & _
				StringTrn.Shortcut & StringTrn.EndSpace)
		End If
		Temp = LTrim(trnStr)
		y = Len(Temp)
		If y > x Then LastStringTrn = Left(Temp,y - x)
	Else
		LastStringTrn = Trim(trnStr)
	End If

	'保留符合条件的加速器翻译
	If StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" And KeepShortKey <> "" Then
		SrcKeyArr = ReSplit(ShortcutKeySrc,"+",-1)
		x = UBound(SrcKeyArr)
		TrnKeyArr = ReSplit(ShortcutKeyTrn,"+",-1)
		y = UBound(TrnKeyArr)
		If x = y Then
			For i = 0 To x
				Temp = Trim(SrcKeyArr(i))
				TempBak = Trim(TrnKeyArr(i))
				If Temp <> "" And TempBak <> "" Then
					If CheckKeyCode(TempBak,KeepShortKey) <> 0 Then
						StringSrc.Shortcut = Replace(StringSrc.Shortcut,Temp,TempBak)
					End If
				End If
			Next i
		End If
	End If

	'PSL.Output "------------------------------ "   		'调试用
	'PSL.Output "srcStr = " & srcStr                		'调试用
	'PSL.Output "trnStr = " & trnStr               			'调试用
	'PSL.Output "SpaceTrn = " & StringTrn.Spaces            '调试用
	'PSL.Output "KeySrc = " & StringTrnKeySrc               '调试用
	'PSL.Output "acckeySrc = " & StringSrc.AccKey          	'调试用
	'PSL.Output "acckeyTrn = " & StringTrn.AccKey         	'调试用
	'PSL.Output "EndStringSrc = " & StringSrc.EndString   	'调试用
	'PSL.Output "EndStringTrn = " & StringTrn.EndString   	'调试用
	'PSL.Output "ShortcutSrc = " & StringSrc.Shortcut     	'调试用
	'PSL.Output "ShortcutTrn = " & StringTrn.Shortcut     	'调试用
	'PSL.Output "ExpStringTrn = " & StringTrn.ExpString   	'调试用
	'PSL.Output "PreStringTrn = " & StringTrn.PreString   	'调试用
	'PSL.Output "LastStringTrn = " & LastStringTrn  		'调试用
	'PSL.Output "MoveAcckey = " & MoveAcckey                '调试用

	'备份参数值
	SpaceTrnBak = StringTrn.Spaces
	ExpStringTrnBak = StringTrn.ExpString
	PreStringTrnBak = StringTrn.PreString
	ShortcutSrcBak = StringSrc.Shortcut
	EndStringSrcBak = StringTrn.EndString

	'字串内容选择处理
	If AllCont <> 1 Then
		If Acceler <> 1 Then
			ShortcutInSource = 0
			ShortcutInTarget = 0
			ShortcutInBoth = 0
		End If
		If EndChar <> 1 Then
			EndCharInSource = 0
			EndCharInTarget = 0
			EndCharInBoth = 0
			TranslateEndChar = 0
		End If
		If AccKey <> 1 Then
			AccessKeyInSource = 0
			AccessKeyInTarget = 0
			AccessKeyInBoth = 0
			If AsiaKey = 1 Then
				If StringTrn.AccKey = StringTrn.AccKeyIFR Then AsiaKey = 0
			Else
				If StringTrn.AccKey <> StringTrn.AccKeyIFR Then AsiaKey = 1
			End If
		End If
	End If

	'数据集成
	If ProjectID >= 0 Then
		'执行检查规则
		If StringSrc.PreSpace <> "" And StringTrn.PreSpace = "" Then
			If LeadingSpaceInSource = 0 Then StringSrc.PreSpace = StringTrn.PreSpace
		ElseIf StringSrc.PreSpace = "" And StringTrn.PreSpace <> "" Then
			If LeadingSpaceInTarget = 0 Then StringSrc.PreSpace = StringTrn.PreSpace
		ElseIf StringSrc.PreSpace <> "" And StringTrn.PreSpace <> "" Then
			If LeadingSpaceInBoth = 0 Then StringSrc.PreSpace = StringTrn.PreSpace
			If LeadingSpaceInBoth = 2 Then StringSrc.PreSpace = ""
		End If
		If StringSrc.EndSpace <> "" And StringTrn.EndSpace = "" Then
			If LeadingSpaceInSource = 0 Then StringSrc.EndSpace = StringTrn.EndSpace
		ElseIf StringSrc.EndSpace = "" And StringTrn.EndSpace <> "" Then
			If TrailingSpaceInTarget = 0 Then StringSrc.EndSpace = StringTrn.EndSpace
		ElseIf StringSrc.EndSpace <> "" And StringTrn.EndSpace <> "" Then
			If TrailingSpaceInBoth = 0 Then StringSrc.EndSpace = StringTrn.EndSpace
			If TrailingSpaceInBoth = 2 Then StringSrc.EndSpace = ""
		End If
		If StringSrc.AccKey <> "" And StringTrn.AccKey = "" Then
			If AccessKeyInSource = 0 Then
				StringSrc.AccKey = StringTrn.AccKey
				StringSrc.AccKeyIFR = StringTrn.AccKeyIFR
				StringSrc.AccKeyKey = StringTrn.AccKeyKey
			End If
		ElseIf StringSrc.AccKey = "" And StringTrn.AccKey <> "" Then
			If AccessKeyInTarget = 0 Then
				StringSrc.AccKey = StringTrn.AccKey
				StringSrc.AccKeyIFR = StringTrn.AccKeyIFR
				StringSrc.AccKeyKey = StringTrn.AccKeyKey
			End If
		ElseIf StringSrc.AccKey <> "" And StringTrn.AccKey <> "" Then
			If AccessKeyInBoth = 0 Then
				StringSrc.AccKey = StringTrn.AccKey
				StringSrc.AccKeyIFR = StringTrn.AccKeyIFR
				StringSrc.AccKeyKey = StringTrn.AccKeyKey
			ElseIf AccessKeyInBoth = 2 Then
				StringSrc.AccKey = ""
			End If
		End If
		If StringSrc.EndString <> "" And StringTrn.EndString = "" Then
			If EndCharInSource = 0 Then StringSrc.EndString = StringTrn.EndString
		ElseIf StringSrc.EndString = "" And StringTrn.EndString <> "" Then
			If EndCharInTarget = 0 Then StringSrc.EndString = StringTrn.EndString
		ElseIf StringSrc.EndString <> "" And StringTrn.EndString <> "" Then
			If EndCharInBoth = 0 Then StringSrc.EndString = StringTrn.EndString
			If EndCharInBoth = 2 Then StringSrc.EndString = ""
		End If
		If StringSrc.Shortcut <> "" And StringTrn.Shortcut = "" Then
			If ShortcutInSource = 0 Then StringSrc.Shortcut = StringTrn.Shortcut
		ElseIf StringSrc.Shortcut = "" And StringTrn.Shortcut <> "" Then
			If ShortcutInTarget = 0 Then StringSrc.Shortcut = StringTrn.Shortcut
		ElseIf StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" Then
			If ShortcutInBoth = 0 Then StringSrc.Shortcut = StringTrn.Shortcut
			If ShortcutInBoth = 2 Then StringSrc.Shortcut = ""
		End If

		'设置快捷键方式
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			StringTrn.ExpString = ""
			If AsiaKey = 0 Then StringTrn.PreString = ""
		End If
		If StringTrn.AccKey = StringTrn.AccKeyIFR Then
			StringTrn.AccKey = StringTrn.AccKeyIFR & StringTrn.AccKeyKey
		End If
		If StringSrc.AccKey <> "" Then
			If AsiaKey = 0 Then
				StringSrc.AccKey = StringSrc.AccKeyIFR & StringSrc.AccKeyKey
				KeySrc = StringSrc.AccKeyIFR
			Else
				StringSrc.AccKey = "(" & StringSrc.AccKeyIFR & UCase(StringSrc.AccKeyKey) & ")"
				KeySrc = StringSrc.AccKey
			End If
		End If

		'确定快捷键是否被移动
		Stemp = False
		If StringSrc.AccKey <> "" Then
			i = InStr(StringSrc.Shortcut,StringSrc.AccKeyIFR)
			j = InStr(StringTrn.Shortcut,StringTrn.AccKeyIFR)
			x = InStr(StringSrc.EndString,StringSrc.AcckeyIFR)
			y = InStr(StringTrn.EndString,StringTrn.AcckeyIFR)
			If LCase(StringSrc.AccKey) = LCase(StringTrn.AccKey) Then
				If AccKeyInShort = 1 Then
					If i > 0 And j = 0 Then MoveAcckey = "ShortcutSrc"
					If x > 0 And y = 0 Then MoveAcckey = "EndStringSrc"
					If i = 0 And j > 0 Then MoveAcckey = "ShortcutTrn"
					If x = 0 And y > 0 Then MoveAcckey = "EndStringTrn"
				Else
					If j > 0 Then MoveAcckey = "ShortcutTrn"
					If y > 0 Then MoveAcckey = "EndStringTrn"
				End If
			Else
				i = InStr(ShortcutSrcBak,StringSrc.AccKeyIFR)
				x = InStr(EndStringSrcBak,StringSrc.AccKeyIFR)
				If i = 0 And j > 0 Then Stemp = True
				If x = 0 And y > 0 Then Stemp = True
			End If
		End If

		'移动或删除快捷键前的终止符
		If StringTrn.PreString <> "" And AsiaKey = 1 Then
			If StringSrc.EndString & StringTrn.EndString = "" Then
				StringSrc.EndString = StringTrn.PreString
			End If
			StringTrn.PreString = ""
		End If

		'删除所有多余空格
		If DeleteExtraSpace = 1 Then
			If StringTrn.Spaces <> "" Then
				If AsiaKey = 0 Then
					If Len(StringTrn.Spaces) > 1 Then StringTrn.Spaces = Space(1)
				Else
					StringTrn.Spaces = ""
				End If
			End If
			If StringTrn.PreString <> "" Then StringTrn.PreString = Trim(StringTrn.PreString)
			If StringTrn.ExpString <> "" Then StringTrn.ExpString = Trim(StringTrn.ExpString)
			If StringSrc.Shortcut <> "" Then StringSrc.Shortcut = Trim(StringSrc.Shortcut)
			If StringSrc.EndString <> "" Then
				If StringSrc.EndString = Space(1) & LTrim(StringSrc.EndString) Then
					StringSrc.EndString = RTrim(StringSrc.EndString)
				Else
					StringSrc.EndString = Trim(StringSrc.EndString)
				End If
			End If
		End If

		'确定快捷键的方式
		If StringSrc.AccKey <> "" And AccKeyInShort = 1 Then
			If InStr(StringSrc.EndString & StringSrc.Shortcut,StringSrc.AccKeyIFR) Then
				If Stemp = False Then
					StringSrc.AccKey = StringSrc.AccKeyIFR & StringSrc.AccKeyKey
					KeySrc = ""
				End If
			End If
		End If
		If StringSrc.AccKey = "" Or KeySrc <> "" Then
			If InStr(StringSrc.Shortcut,StringSrc.AccKeyIFR) Then
				StringSrc.Shortcut = Replace(StringSrc.Shortcut,StringSrc.AccKeyIFR,"")
			End If
			If InStr(StringSrc.EndString,StringSrc.AccKeyIFR) Then
				StringSrc.EndString = Replace(StringSrc.EndString,StringSrc.AccKeyIFR,"")
			End If
		End If

		'自动翻译符合条件的终止符
		If StringSrc.EndString <> "" And TranslateEndChar = 1 And AutoTrnEndChar <> "" Then
			Temp = Replace(Trim(StringSrc.EndString),StringSrc.AccKeyIFR,"")
			If Trim(Temp) <> "" Then
				For i = 0 To UBound(AutoTrnEndCharArr)
					FindStr = Trim(AutoTrnEndCharArr(i))
					If InStr(FindStr,"|") Then
						TempArray = ReSplit(FindStr,"|")
						If Temp = TempArray(0) Then
							StringSrc.EndString = Replace(StringSrc.EndString,Temp,TempArray(1))
							Exit For
						End If
					End If
				Next i
			End If
		End If

		'查找快捷键字符并设置快捷键
		If StringSrc.AccKey <> "" And KeySrc <> "" And AsiaKey = 0 Then
			If LCase(StringSrc.AccKey) <> LCase(StringTrn.AccKey) Or MoveAcckey <> "" Then
				For i = 0 To 3
					Temp = ""
					If AccKeyInShort = 0 Then
						If i = 0 Then Temp = LastStringTrn
						If i = 1 Then Temp = StringTrn.ExpString
					Else
						If i = 0 Then Temp = LastStringTrn
						If i = 1 Then Temp = StringTrn.ExpString
						If i = 2 Then Temp = StringSrc.Shortcut
						If i = 3 Then Temp = StringSrc.EndString
					End If
					If Trim(Temp) <> "" Then
						StringTrn.AccKeyPos = InStr(Temp,StringSrc.AccKeyKey)
						If StringTrn.AccKeyPos = 0 Then
							StringTrn.AccKeyPos = InStr(LCase(Temp),LCase(StringSrc.AccKeyKey))
						End If
						If StringTrn.AccKeyPos > 0 Then
							StringTrn.AccKeyKey = Mid(Temp,StringTrn.AccKeyPos,1)
							Temp = Replace(Temp,StringTrn.AccKeyKey,StringSrc.AccKeyIFR & StringTrn.AccKeyKey,,1)
							StringSrc.AccKey = StringSrc.AccKeyIFR & StringTrn.AccKeyKey
							KeySrc = ""
						End If
					End If
					If AccKeyInShort = 0 Then
						If i = 0 Then LastStringTrn = Temp
						If i = 1 Then StringTrn.ExpString = Temp
					Else
						If i = 0 Then LastStringTrn = Temp
						If i = 1 Then StringTrn.ExpString = Temp
						If i = 2 Then StringSrc.Shortcut = Temp
						If i = 3 Then StringSrc.EndString = Temp
					End If
					If KeySrc = "" Then Exit For
				Next i
				If KeySrc <> "" Then
					If AddAccessKeyWithFirstChar = 1 Then
						If Trim(LastStringTrn) <> "" Then
							i = CheckStrRegExp(LastStringTrn,CheckSkipStr(5),0,0)
							If i = 0 Then
								i = Len(LastStringTrn) - Len(LTrim(LastStringTrn)) + 1
							End If
							If i > 0 Then
								PreTrn = Left(LastStringTrn,i - 1)
								AppTrn = Mid(LastStringTrn,i)
								StringTrn.AccKeyKey = Mid(LastStringTrn,i,1)
								LastStringTrn = PreTrn & StringSrc.AccKeyIFR & AppTrn
								StringSrc.AccKey = StringSrc.AccKeyIFR & StringTrn.AccKeyKey
							Else
								StringSrc.AccKey = ""
							End If
							MoveAcckey = ""
							KeySrc = ""
						ElseIf Trim(StringTrn.ExpString) <> "" Then
							i = CheckStrRegExp(StringTrn.ExpString,CheckSkipStr(5),0,0)
							If i = 0 Then
								i = Len(StringTrn.ExpString) - Len(LTrim(StringTrn.ExpString)) + 1
							End If
							If i > 0 Then
								PreTrn = Left(StringTrn.ExpString,i - 1)
								AppTrn = Mid(StringTrn.ExpString,i)
								StringTrn.AccKeyKey = Mid(StringTrn.ExpString,i,1)
								StringTrn.ExpString = PreTrn & StringSrc.AccKeyIFR & AppTrn
								StringSrc.AccKey = StringSrc.AccKeyIFR & StringTrn.AccKeyKey
							Else
								StringSrc.AccKey = ""
							End If
							MoveAcckey = ""
							KeySrc = ""
						End If
					Else
						MoveAcckey = ""
						StringSrc.AccKey = ""
						KeySrc = ""
					End If
				End If
			Else
				StringTrn.AccKey = StringSrc.AccKey
			End If
		End If

		'组织替换字符
		If AsiaKey = 0 Then
			NewStringTrn = StringSrc.PreSpace & LastStringTrn & StringTrn.PreString & StringTrn.Spaces & KeySrc & _
							StringTrn.ExpString & StringSrc.EndString & StringSrc.Shortcut & StringSrc.EndSpace
		Else
			NewStringTrn = StringSrc.PreSpace & LastStringTrn & StringTrn.PreString & StringTrn.Spaces & _
							StringTrn.ExpString & KeySrc & StringSrc.EndString & StringSrc.Shortcut & StringSrc.EndSpace
		End If

		'字串替换
		If StringReplace <> NewStringTrn Then StringReplace = NewStringTrn
	End If

	'还原参数
	StringTrn.Spaces = SpaceTrnBak
	StringTrn.ExpString = ExpStringTrnBak
	StringTrn.PreString = PreStringTrnBak

	'删除终止符和加速器中的快捷键，以便可以正确比较终止符和加速器
	If InStr(StringSrc.Shortcut,StringSrc.AccKeyIFR) Then StringSrc.Shortcut = Replace(StringSrc.Shortcut,StringSrc.AccKeyIFR,"")
	If InStr(StringSrc.Shortcut,StringTrn.AccKeyIFR) Then StringSrc.Shortcut = Replace(StringSrc.Shortcut,StringTrn.AccKeyIFR,"")
	If InStr(StringTrn.Shortcut,StringSrc.AccKeyIFR) Then StringTrn.Shortcut = Replace(StringTrn.Shortcut,StringSrc.AccKeyIFR,"")
	If InStr(StringTrn.Shortcut,StringTrn.AccKeyIFR) Then StringTrn.Shortcut = Replace(StringTrn.Shortcut,StringTrn.AccKeyIFR,"")
	If InStr(StringSrc.EndString,StringSrc.AccKeyIFR) Then StringSrc.EndString = Replace(StringSrc.EndString,StringSrc.AccKeyIFR,"")
	If InStr(StringSrc.EndString,StringTrn.AccKeyIFR) Then StringSrc.EndString = Replace(StringSrc.EndString,StringTrn.AccKeyIFR,"")
	If InStr(StringTrn.EndString,acckeyIFRSrc) Then StringTrn.EndString = Replace(StringTrn.EndString,StringSrc.AccKeyIFR,"")
	If InStr(StringTrn.EndString,acckeyIFRTrn) Then StringTrn.EndString = Replace(StringTrn.EndString,StringTrn.AccKeyIFR,"")
	If InStr(MoveAcckey,StringSrc.AccKeyIFR) Then MoveAcckey = Replace(MoveAcckey,StringSrc.AccKeyIFR,"")
	If InStr(MoveAcckey,StringTrn.AccKeyIFR) Then MoveAcckey = Replace(MoveAcckey,StringTrn.AccKeyIFR,"")
End Function


' 修改消息输出
Function ReplaceMassage(ByVal CheckID As Long,ByVal ProjectID As Long) As String
	Dim i As Long,j As Long,m As Long,a As Long,v As Long,d As Long,r As Long
	Dim ModifiedMsg() As String,AddedMsg() As String,MovedMsg() As String
	Dim DeledMsg() As String,ReplacedMsg() As String,Massage() As String
	Dim sL As Long,tL As Long,sR As Long,tR As Long,MsgList() As String
	Dim AsiaKey As Long,DeleteExtraSpace As Long

	If getMsgList(UIDataList,MsgList,"ReplaceMassage",1) = False Then Exit Function

	'获取选定配置的参数
	TempArray = ReSplit(CheckDataList(CheckID),JoinStr)
	SetsArray = ReSplit(TempArray(1),SubJoinStr)
	AsiaKey = StrToLong(SetsArray(4))

	TempArray = ReSplit(ProjectDataList(ProjectID),JoinStr)
	SetsArray = ReSplit(TempArray(1),LngJoinStr)
	DeleteExtraSpace = StrToLong(SetsArray(15))
	If StrToLong(SetsArray(20)) = 0 Then
		For i = 23 To 27
			MsgList(i) = MsgList(i + 8)
		Next i
	End If

	'参数初始化
	ReDim ModifiedMsg(m),AddedMsg(a),MovedMsg(v),DeledMsg(d),ReplacedMsg(r),Massage(0)

	'计算快捷键
	If StringSrc.AccKey <> StringTrn.AccKey Then
		If StringSrc.AccKey <> "" And StringTrn.AccKey <> "" Then
			ModifiedMsg(m) = MsgList(0)
		ElseIf StringSrc.AccKey <> "" And StringTrn.AccKey = "" Then
			AddedMsg(a) = MsgList(0)
		ElseIf StringSrc.AccKey = "" And StringTrn.AccKey <> "" Then
			DeledMsg(d) = MsgList(0)
		ElseIf LCase(StringSrc.AccKey) = LCase(StringTrn.AccKey) Then
			ModifiedMsg(m) = MsgList(1)
		End If
		If ModifiedMsg(m) <> "" Then m = m + 1
		If AddedMsg(a) <> "" Then a = a + 1
		If DeledMsg(d) <> "" Then d = d + 1
	End If

	'计算终止符
	If StringSrc.EndString <> StringTrn.EndString Then
		If StringSrc.EndString <> "" And StringTrn.EndString <> "" Then
			If Trim(StringSrc.EndString) <> Trim(StringTrn.EndString) Then
				ReDim Preserve ModifiedMsg(m)
				ModifiedMsg(m) = MsgList(2)
				m = m + 1
			End If
		ElseIf StringSrc.EndString <> "" And StringTrn.EndString = "" Then
			If Trim(StringSrc.EndString) <> Trim(PreStringTrn) Then
				ReDim Preserve AddedMsg(a)
				AddedMsg(a) = MsgList(2)
				a = a + 1
			End If
		ElseIf StringSrc.EndString = "" And StringTrn.EndString <> "" Then
			ReDim Preserve DeledMsg(d)
			DeledMsg(d) = MsgList(2)
			d = d + 1
		End If
	End If

	'计算加速器
	If StringSrc.Shortcut <> StringTrn.Shortcut Then
		If StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" Then
			ReDim Preserve ModifiedMsg(m)
			ModifiedMsg(m) = MsgList(3)
			m = m + 1
		ElseIf StringSrc.Shortcut <> "" And StringTrn.Shortcut = "" Then
			ReDim Preserve AddedMsg(a)
			AddedMsg(a) = MsgList(3)
			a = a + 1
		ElseIf StringSrc.Shortcut = "" And StringTrn.Shortcut <> "" Then
			ReDim Preserve DeledMsg(d)
			DeledMsg(d) = MsgList(3)
			d = d + 1
		End If
	End If

	'计算前置空格
	If StringSrc.PreSpace <> StringTrn.PreSpace Then
		If StringSrc.PreSpace <> "" And StringTrn.PreSpace <> "" Then
			ReDim Preserve ModifiedMsg(m)
			ModifiedMsg(m) = MsgList(4)
			m = m + 1
		ElseIf StringSrc.PreSpace <> "" And StringTrn.PreSpace = "" Then
			ReDim Preserve AddedMsg(a)
			AddedMsg(a) = MsgList(4)
			a = a + 1
		ElseIf StringSrc.PreSpace = "" And StringTrn.PreSpace <> "" Then
			ReDim Preserve DeledMsg(d)
			DeledMsg(d) = MsgList(4)
			d = d + 1
		End If
	End If

	'计算后置空格
	If StringSrc.EndSpace <> StringTrn.EndSpace Then
		If StringSrc.EndSpace <> "" And StringTrn.EndSpace <> "" Then
			ReDim Preserve ModifiedMsg(m)
			ModifiedMsg(m) = MsgList(5)
			m = m + 1
		ElseIf StringSrc.EndSpace <> "" And StringTrn.EndSpace = "" Then
			ReDim Preserve AddedMsg(a)
			AddedMsg(a) = MsgList(5)
			a = a + 1
		ElseIf StringSrc.EndSpace = "" And StringTrn.EndSpace <> "" Then
			ReDim Preserve DeledMsg(d)
			DeledMsg(d) = MsgList(5)
			d = d + 1
		End If
	End If

	'计算快捷键、终止符、加速器前后的空格
	If DeleteExtraSpace = 1 Then
		If StringTrn.Spaces <> "" And StringTrn.ExpString = "" Then
			ReDim Preserve DeledMsg(d)
			If StringSrc.Shortcut <> "" Or StringTrn.Shortcut <> "" Then DeledMsg(d) = MsgList(6)
			If StringSrc.EndString <> "" Or StringTrn.EndString <> "" Then DeledMsg(d) = MsgList(7)
			If StringSrc.AccKey <> "" Or StringTrn.AccKey <> "" Then DeledMsg(d) = MsgList(10)
			If DeledMsg(d) <> "" Then d = d + 1
		ElseIf StringTrn.Spaces <> "" And StringTrn.ExpString <> "" Then
			i = Len(StringTrn.ExpString) - Len(LTrim(StringTrn.ExpString))
			j = Len(StringTrn.ExpString) - Len(RTrim(StringTrn.ExpString))
			ReDim Preserve DeledMsg(d)
			If i = 0 Then DeledMsg(d) = MsgList(10) Else DeledMsg(d) = MsgList(12)
			d = d + 1
			If Trim(StringTrn.ExpString) <> "" And j > 0 Then
				ReDim Preserve DeledMsg(d)
				If StringTrn.Shortcut <> "" Then DeledMsg(d) = MsgList(6)
				If StringTrn.EndString <> "" Then DeledMsg(d) = MsgList(7)
				If DeledMsg(d) <> "" Then d = d + 1
			End If
		ElseIf StringTrn.Spaces = "" And StringTrn.ExpString <> "" Then
			i = Len(StringTrn.ExpString) - Len(LTrim(StringTrn.ExpString))
			j = Len(StringTrn.ExpString) - Len(RTrim(StringTrn.ExpString))
			If i > 0 Then
				ReDim Preserve DeledMsg(d)
				DeledMsg(d) = MsgList(11)
				d = d + 1
			End If
			If Trim(StringTrn.ExpString) <> "" And j > 0 Then
				ReDim Preserve DeledMsg(d)
				If StringTrn.Shortcut <> "" Then DeledMsg(d) = MsgList(6)
				If StringTrn.EndString <> "" Then DeledMsg(d) = MsgList(7)
				If DeledMsg(d) <> "" Then d = d + 1
			End If
		End If
		If StringTrn.PreString <> "" And StringTrn.PreString <> LTrim(StringTrn.PreString) Then
			ReDim Preserve DeledMsg(d)
			DeledMsg(d) = MsgList(13)
			d = d + 1
		End If

		'计算终止符前后的空格
		If StringSrc.EndString <> StringTrn.EndString Then
			If StringSrc.EndString <> "" And StringTrn.EndString <> "" Then
				sL = Len(StringSrc.EndString) - Len(LTrim(StringSrc.EndString))
				tL = Len(StringTrn.EndString) - Len(LTrim(StringTrn.EndString))
				sR = Len(StringSrc.EndString) - Len(RTrim(StringSrc.EndString))
				tR = Len(StringTrn.EndString) - Len(RTrim(StringTrn.EndString))
				i = sL - tL
				j = sR - tR
				If i > 0 Or j > 0 Then
					ReDim Preserve AddedMsg(a)
					If i > 0 And j = 0 Then
						AddedMsg(a) = MsgList(7)
					ElseIf i = 0 And j > 0 Then
						AddedMsg(a) = MsgList(8)
					ElseIf i > 0 And j > 0 Then
						AddedMsg(a) = MsgList(9)
					End If
					If AddedMsg(a) <> "" Then a = a + 1
				End If
				If i < 0 Or j < 0 Then
					ReDim Preserve DeledMsg(d)
					If i < 0 And j = 0 Then
						DeledMsg(d) = MsgList(7)
					ElseIf i = 0 And j < 0 Then
						DeledMsg(d) = MsgList(8)
					ElseIf i < 0 And j < 0 Then
						DeledMsg(d) = MsgList(9)
					End If
					If DeledMsg(d) <> "" Then d = d + 1
				End If
			ElseIf StringSrc.EndString = "" And StringTrn.EndString <> "" Then
				i = Len(StringTrn.EndString) - Len(LTrim(StringTrn.EndString))
				j = Len(StringTrn.EndString) - Len(RTrim(StringTrn.EndString))
				If i + j <> 0 Then
					ReDim Preserve DeledMsg(d)
					If i > 0 And j = 0 Then
						DeledMsg(d) = MsgList(7)
					ElseIf i = 0 And j > 0 Then
						DeledMsg(d) = MsgList(8)
					ElseIf i > 0 And j > 0 Then
						DeledMsg(d) = MsgList(9)
					End If
					If DeledMsg(d) <> "" Then d = d + 1
				End If
			End If
		End If
	End If

	'计算快捷键前后的字符移动或删除
	If AsiaKey = 1 Then
		If StringTrn.ExpString <> "" And Trim(StringTrn.ExpString) <> "" Then
			ReDim Preserve MovedMsg(v)
			If StringSrc.AccKey <> "" Or StringTrn.AccKey <> "" Then MovedMsg(v) = MsgList(14)
			If StringSrc.EndSpace <> "" Or StringTrn.EndSpace <> "" Then MovedMsg(v) = MsgList(15)
			If StringSrc.Shortcut <> "" Or StringTrn.Shortcut <> "" Then MovedMsg(v) = MsgList(16)
			If StringSrc.EndString <> "" Or StringTrn.EndString <> "" Then MovedMsg(v) = MsgList(17)
			If MovedMsg(v) <> "" Then v = v + 1
		End If
		If StringTrn.PreString <> "" And Trim(StringTrn.PreString) <> "" Then
			If StringSrc.EndString = "" And StringTrn.EndString = "" Then
				ReDim Preserve MovedMsg(v)
				MovedMsg(v) = MsgList(17)
				v = v + 1
			ElseIf StringSrc.EndString = "" And StringTrn.EndString <> "" Then
				ReDim Preserve DeledMsg(d)
				DeledMsg(d) = MsgList(18)
				d = d + 1
			ElseIf StringSrc.EndString <> "" And StringTrn.EndString = "" Then
				If Trim(StringSrc.EndString) = Trim(StringTrn.PreString) Then
					ReDim Preserve MovedMsg(v)
					MovedMsg(v) = MsgList(17)
					v = v + 1
				Else
					ReDim Preserve DeledMsg(d)
					DeledMsg(d) = MsgList(18)
					d = d + 1
				End If
			ElseIf StringSrc.EndString <> "" And StringTrn.EndString <> "" Then
				ReDim Preserve DeledMsg(d)
				DeledMsg(d) = MsgList(18)
				d = d + 1
			End If
		End If
	End If
	If MoveAcckey <> "" Then
		ReDim Preserve MovedMsg(v)
		If MoveAcckey = "EndStringSrc" Then
			MovedMsg(v) = MsgList(19)
		ElseIf MoveAcckey = "ShortcutSrc" Then
			MovedMsg(v) = MsgList(20)
		ElseIf MoveAcckey = "EndStringTrn" Then
			MovedMsg(v) = MsgList(21)
		ElseIf MoveAcckey = "ShortcutTrn" Then
			MovedMsg(v) = MsgList(22)
		End If
		If MovedMsg(v) <> "" Then v = v + 1
	End If

	'计算替换字符
	If PreRepStr <> "" And AppRepStr <> "" Then
		ReDim Preserve ReplacedMsg(r + 1)
		ReplacedMsg(r) = Replace(PreRepStr,JoinStr,MsgList(28))
		ReplacedMsg(r + 1) = Replace(AppRepStr,JoinStr,MsgList(28))
	ElseIf PreRepStr <> "" And AppRepStr = "" Then
		ReDim Preserve DeledMsg(d)
		DeledMsg(d) = Replace(PreRepStr,JoinStr,MsgList(28))
	End If

	'组织消息及计数
	i = 0
	If CheckArray(ModifiedMsg) = True Then
		Massage(i) = Replace(MsgList(23),"%s",Join(ClearArray(ModifiedMsg,1),MsgList(28)))
		ModifiedCount = ModifiedCount + 1
		i = i + 1
	End If
	If CheckArray(AddedMsg) = True Then
		ReDim Preserve Massage(i)
		Massage(i) = Replace(MsgList(24),"%s",Join(ClearArray(AddedMsg,1),MsgList(28)))
		AddedCount = AddedCount + 1
		i = i + 1
	End If
	If CheckArray(MovedMsg) = True Then
		ReDim Preserve Massage(i)
		Massage(i) = Replace(MsgList(25),"%s",Join(ClearArray(MovedMsg,1),MsgList(28)))
		MovedCount = MovedCount + 1
		i = i + 1
	End If
	If CheckArray(DeledMsg) = True Then
		ReDim Preserve Massage(i)
		Massage(i) = Replace(MsgList(26),"%s",Join(ClearArray(DeledMsg,1),MsgList(28)))
		DeledCount = DeledCount + 1
		i = i + 1
	End If
	If CheckArray(ReplacedMsg) = True Then
		ReDim Preserve Massage(i)
		Massage(i) = Replace(Replace(MsgList(27),"%s",ReplacedMsg(r)),"%d",ReplacedMsg(r + 1))
		ReplacedCount = ReplacedCount + 1
	End If

	If CheckArray(Massage) = True Then ReplaceMassage = Join(Massage,MsgList(29)) & MsgList(30)
End Function


'输出行数错误消息
Function LineErrMassage(ByVal srcLineNum As Integer,ByVal trnLineNum As Integer,LineNumErrCount As Integer) As String
	Dim MsgList() As String
	If getMsgList(UIDataList,MsgList,"LineErrMassage",1) = False Then Exit Function
	If srcLineNum <> trnLineNum Then
		If srcLineNum > trnLineNum Then
			LineErrMassage = Replace(MsgList(0),"%s",CStr(srcLineNum - trnLineNum))
			LineNumErrCount = LineNumErrCount + 1
		ElseIf srcLineNum < trnLineNum Then
			LineErrMassage = Replace(MsgList(1),"%s",CStr(trnLineNum - srcLineNum))
			LineNumErrCount = LineNumErrCount + 1
		End If
	End If
End Function


'输出快捷键数错误消息
Function AccKeyErrMassage(ByVal srcAccKeyNum As Integer,ByVal trnAccKeyNum As Integer,accKeyNumErrCount As Integer) As String
	Dim MsgList() As String
	If getMsgList(UIDataList,MsgList,"AccKeyErrMassage",1) = False Then Exit Function
	If srcAccKeyNum <> trnAccKeyNum Then
		If srcAccKeyNum > trnAccKeyNum Then
			AccKeyErrMassage = Replace(MsgList(0),"%s",CStr(srcAccKeyNum - trnAccKeyNum))
			accKeyNumErrCount = accKeyNumErrCount + 1
		ElseIf srcAccKeyNum < trnAccKeyNum Then
			AccKeyErrMassage = Replace(MsgList(1),"%s",CStr(trnAccKeyNum - srcAccKeyNum))
			accKeyNumErrCount = accKeyNumErrCount + 1
		End If
	End If
End Function


'翻译消息输出
Function TranMassage(ByVal tCount As Integer,ByVal sCount As Integer,ByVal nCount As Integer,ByVal mCount As Integer,ByVal eCount As Integer) As String
	Dim MsgList() As String
	If getMsgList(UIDataList,MsgList,"TranMassage",1) = False Then Exit Function
	If tCount + sCount + nCount + mCount + eCount = 0 Then
		TranMassage = MsgList(0)
	Else
		TranMassage = Replace(MsgList(1) & MsgList(2),"%t",CStr(tCount))
		TranMassage = Replace(TranMassage,"%s",CStr(sCount))
		TranMassage = Replace(TranMassage,"%n",CStr(nCount))
		TranMassage = Replace(TranMassage,"%m",CStr(mCount))
		TranMassage = Replace(TranMassage,"%e",CStr(eCount))
	End If
End Function


'输出程序错误消息
Sub sysErrorMassage(sysError As ErrObject,ByVal fType As Long)
	Dim TempArray() As String,MsgList() As String
	Dim ErrorNumber As Long,ErrorSource As String,ErrorDescription As String
	Dim TitleMsg As String,ContinueMsg As String,Msg As String

	ErrorNumber = sysError.Number
	ErrorSource = sysError.Source
	ErrorDescription = sysError.Description

	TitleMsg = "Error"
	Select Case fType
	Case 0
		ContinueMsg = vbCrLf & vbCrLf & "The program cannot continue and will exit."
	Case 1
		ContinueMsg = vbCrLf & vbCrLf & "Do you want to continue?"
	Case 2
		ContinueMsg = vbCrLf & vbCrLf & "The program will continue to run."
	End Select

	If CheckINIArray(UIDataList) = True Then
		If getMsgList(UIDataList,MsgList,"sysErrorMassage",3) = False Then
			If getMsgList(UIDataList,MsgList,"Main",3) = False Then
				Msg = "The following file is missing [Main|sysErrorMassage] section." & vbCrLf & "%s"
				Msg = Replace(Msg,"%s",LangFile)
			Else
				TitleMsg = MsgList(42)
				If fType <> 0 Then ContinueMsg = MsgList(94) Else ContinueMsg = MsgList(95)
				Msg = Replace(Replace(MsgList(75),"%s","sysErrorMassage"),"%d",LangFile)
			End If
		Else
			TitleMsg = MsgList(0)
			Select Case fType
			Case 0
				ContinueMsg = MsgList(10)
			Case 1
				ContinueMsg = MsgList(11)
			Case 2
				ContinueMsg = MsgList(12)
			End Select

			Select Case ErrorSource
			Case ""
				If ErrorNumber = 10051 And PSL.Version >= 1500 Then
					Msg = Replace$(MsgList(13),"%s",ErrorSource)
				Else
					Msg = Replace$(Replace$(MsgList(1),"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
				End If
			Case "NotSection"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace(Replace(MsgList(3),"%s",TempArray(1)),"%d",TempArray(0))
			Case "NotValue"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace(Replace(MsgList(4),"%s",TempArray(1)),"%d",TempArray(0))
			Case "NotReadFile"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace(MsgList(5),"%s",TempArray(1))
			Case "NotWriteFile"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace(MsgList(6),"%s",TempArray(1))
			Case "NotINIFile"
				Msg = Replace(MsgList(7),"%s",ErrorDescription)
			Case "NotExitFile"
				Msg = Replace(MsgList(8),"%s",ErrorDescription)
			Case "NotVersion"
				TempArray = ReSplit(ErrorDescription,JoinStr,-1)
				Msg = Replace(MsgList(9),"%s",TempArray(0))
				Msg = Replace(Replace(Msg,"%d",TempArray(1)),"%v",TempArray(2))
			Case Else
				Msg = Replace(MsgList(2),"%s",ErrorSource)
				Msg = Replace(Replace(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
			End Select
		End If
	Else
		Select Case ErrorSource
		Case ""
			If ErrorNumber = 10051 And PSL.Version >= 1500 Then
				Msg = "Unable to open the file. Please verify the file path and file name" & _
						"contains characters in Asian languages. " & vbCrLf & _
						"Note: Passolo 2015 Version of the macro engine does not recognize" & _
						"the file path and file name contains Asian language characters."
			Else
				Msg = "An Error occurred in the program design." & vbCrLf & "Error Code: %d, Content: %v" & _
						vbCrLf & "Please restart the Passolo try and please report to the software developer."
				Msg = Replace$(Replace$(Msg,"%s",CStr(ErrorNumber)),"%v",ErrorDescription)
			End If
		Case "NotSection"
			TempArray = ReSplit(ErrorDescription,JoinStr,-1)
			Msg = "The following file is missing [%s] section." & vbCrLf & "%d"
			Msg = Replace(Replace(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotValue"
			TempArray = ReSplit(ErrorDescription,JoinStr,-1)
			Msg = "The following file is missing [%s] Value." & vbCrLf & "%d"
			Msg = Replace(Replace(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		Case "NotReadFile"
			Msg = Replace(ErrorDescription,JoinStr,vbCrLf)
		Case "NotWriteFile"
			Msg = Replace(ErrorDescription,JoinStr,vbCrLf)
		Case "NotINIFile"
			Msg = "The following contents of the file is not correct." & vbCrLf & "%s"
			Msg = Replace(Msg,"%s",ErrorDescription)
		Case "NotExitFile"
			Msg = "The following file does not exist! Please check and try again." & vbCrLf & "%s"
			Msg = Replace(Msg,"%s",ErrorDescription)
		Case "NotVersion"
			TempArray = ReSplit(ErrorDescription,JoinStr,-1)
			Msg = "The following file version is %d, requires version at least %v." & vbCrLf & "%s"
			Msg = Replace(Msg,"%s",TempArray(0))
			Msg = Replace(Replace(Msg,"%d",TempArray(1)),"%v",TempArray(2))
		Case Else
			Msg = "Your system is missing %s server." & vbCrLf & "Error Code: %d, Content: %v"
			Msg = Replace(Msg,"%s",ErrorSource)
			Msg = Replace(Replace(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
		End Select
	End If

	If Msg <> "" Then
		Msg = Msg & ContinueMsg
		Select Case fType
		Case 0
			MsgBox(Msg,vbOkOnly+vbInformation,TitleMsg)
			Call ExitMacro(1)
		Case 1
			If MsgBox(Msg,vbYesNo+vbInformation,TitleMsg) = vbNo Then
				Call ExitMacro(1)
			End If
		Case Else
			MsgBox(Msg,vbOkOnly+vbInformation,TitleMsg)
		End Select
	End If
End Sub


'进行数组合并
Function MergeArray(srcStrArr() As String,trnStrArr() As String) As String()
	Dim i As Long,srcNum As Long,trnNum As Long
	Dim srcPassNum As Long,trnPassNum As Long,TempArray() As String
	srcNum = UBound(srcStrArr)
	trnNum = UBound(trnStrArr)
	ReDim TempArray(srcNum + trnNum + 1) As String
	For i = 0 To (srcNum + trnNum + 1) Step 2
		If srcNum >= srcPassNum Then
			TempArray(i) = srcStrArr(srcPassNum)
			srcPassNum = srcPassNum + 1
		End If
		If trnNum >= trnPassNum Then
			TempArray(i + 1) = trnStrArr(trnPassNum)
			trnPassNum = trnPassNum + 1
		End If
	Next i
	MergeArray = TempArray
End Function


'字串常数反向转换
Public Function ReConvert(ByVal ConverString As String) As String
	ReConvert = ConverString
	If ReConvert = "" Then Exit Function
	If InStr(ReConvert,"\") Then ReConvert = Replace$(ReConvert,"\","\\")
	If InStr(ReConvert,vbCrLf) Then ReConvert = Replace$(ReConvert,vbCrLf,"\r\n")
	If InStr(ReConvert,vbNewLine) Then ReConvert = Replace$(ReConvert,vbNewLine,"\r\n")
	If InStr(ReConvert,vbCr) Then ReConvert = Replace$(ReConvert,vbCr,"\r")
	If InStr(ReConvert,vbLf) Then ReConvert = Replace$(ReConvert,vbLf,"\n")
	If InStr(ReConvert,vbBack) Then ReConvert = Replace$(ReConvert,vbBack,"\b")
	If InStr(ReConvert,vbFormFeed) Then ReConvert = Replace$(ReConvert,vbFormFeed,"\f")
	If InStr(ReConvert,vbVerticalTab) Then ReConvert = Replace$(ReConvert,vbVerticalTab,"\v")
	If InStr(ReConvert,vbTab) Then ReConvert = Replace$(ReConvert,vbTab,"\t")
	If InStr(ReConvert,vbNullChar) Then ReConvert = Replace$(ReConvert,vbNullChar,"\0")
	ReConvert = ReConvertBRegExp(ReConvert)
End Function


'转换拉丁文扩展字符为十六进制转义符
Private Function ReConvertB(ByVal ConverString As String) As String
	Dim i As Long,Length As Long,Dec As Long,Temp As String
	ReConvertB = ConverString
	i = 1
	Do
		Temp = Mid$(ReConvertB,i,1)
		Dec = AscW(Temp)
		If (Dec > 0 And Dec < 32) Or (Dec > 126 And Dec < 256) Then
			ReConvertB = Replace$(ReConvertB,Temp,"\x" & Right$("0" & Hex$(Dec),2))
			i = i + 3
		End If
		Length = Len(ReConvertB)
		i = i + 1
	Loop Until i > Length
End Function


'转换拉丁文扩展字符为十六进制转义符
Private Function ReConvertBRegExp(ByVal ConverString As String) As String
	Dim i As Long,ConvCode As String,Matches As Object
	ReConvertBRegExp = ConverString
	With RegExp
		.Global = True
		.IgnoreCase = True
		.Pattern = "[\x01-\x1F\x7F-\xFF]"
		Set Matches = .Execute(ConverString)
		If Matches.Count > 0 Then
			For i = 0 To Matches.Count - 1
				ConvCode = Right$("0" & Hex$(AscW(Matches(i).Value)),2)
				ReConvertBRegExp = Replace$(ReConvertBRegExp,Matches(i).Value,"\x" & ConvCode)
			Next i
		End If
	End With
End Function


'字串常数正向转换
Public Function Convert(ByVal ConverString As String) As String
	Convert = ConverString
	If Convert = "" Then Exit Function
	If InStr(Convert,"\") = 0 Then Exit Function
	If InStr(Convert,"\\") Then Convert = Replace$(Convert,"\\","*a!N!d*")
	If InStr(Convert,"\r\n") Then Convert = Replace$(Convert,"\r\n",vbCrLf)
	If InStr(Convert,"\r\n") Then Convert = Replace$(Convert,"\r\n",vbNewLine)
	If InStr(Convert,"\r") Then Convert = Replace$(Convert,"\r",vbCr)
	If InStr(Convert,"\r") Then Convert = Replace$(Convert,"\r",vbNewLine)
	If InStr(Convert,"\n") Then Convert = Replace$(Convert,"\n",vbLf)
	If InStr(Convert,"\b") Then Convert = Replace$(Convert,"\b",vbBack)
	If InStr(Convert,"\f") Then Convert = Replace$(Convert,"\f",vbFormFeed)
	If InStr(Convert,"\v") Then Convert = Replace$(Convert,"\v",vbVerticalTab)
	If InStr(Convert,"\t") Then Convert = Replace$(Convert,"\t",vbTab)
	If InStr(Convert,"\'") Then Convert = Replace$(Convert,"\'","'")
	If InStr(Convert,"\""") Then Convert = Replace$(Convert,"\""","""")
	If InStr(Convert,"\?") Then Convert = Replace$(Convert,"\?","?")
	If InStr(Convert,"\") Then Convert = ConvertBRegExp(Convert)
	If InStr(Convert,"\0") Then Convert = Replace$(Convert,"\0",vbNullChar)
	If InStr(Convert,"*a!N!d*") Then Convert = Replace$(Convert,"*a!N!d*","\")
End Function


'转换八进制或十六进制转义符
Private Function ConvertB(ByVal ConverString As String) As String
	Dim i As Long,j As Long,EscStr As String,ConvCode As String
	ConvertB = ConverString
	i = InStr(ConvertB,"\")
	Do While i > 0
		EscStr = Mid$(ConvertB,i,2)
		Select Case EscStr
		Case "\x", "\X"
			ConvCode = Mid$(ConvertB,i + 2,2)
			If CheckStr(UCase$(ConvCode),CheckHexStr(0).AscRange,0,1) = True Then
				j = Val("&H" & ConvCode)
				ConvertB = Replace$(ConvertB,EscStr & ConvCode,Val2Bytes(j,2))
			End If
		Case "\u", "\U"
			ConvCode = Mid$(ConvertB,i + 2,4)
			If CheckStr(UCase$(ConvCode),CheckHexStr(1).AscRange,0,1) = True Then
				j = Val("&H" & ConvCode)
				ConvertB = Replace$(ConvertB,EscStr & ConvCode,Val2Bytes(j,2))
			End If
		Case Is <> ""
			EscStr = "\"
			For j = 3 To 1 Step -1
				ConvCode = Mid$(ConvertB,i + 1,j)
				If CheckStr(ConvCode,CheckHexStr(2).AscRange,0,1) = True Then
					j = Val("&O" & ConvCode)
					If j > 256 Then
						ConvCode = Left$(ConvCode,2)
						j = Val("&O" & ConvCode)
					End If
					ConvertB = Replace$(ConvertB,EscStr & ConvCode,Val2Bytes(j,2))
					Exit For
				End If
			Next j
		End Select
		i = InStr(i + 1,ConvertB,"\")
	Loop
End Function


'转换八进制或十六进制转义符
Private Function ConvertBRegExp(ByVal ConverString As String) As String
	Dim i As Long,j As Long,CodeVal As Long,Matches As Object
	ConvertBRegExp = ConverString
	With RegExp
		.Global = True
		.IgnoreCase = True
		For i = 0 To UBound(CheckHexStr)
			.Pattern = CheckHexStr(i).Range
			Set Matches = .Execute(ConverString)
			If Matches.Count > 0 Then
				For j = 0 To Matches.Count - 1
					If i = 0 Then
						If Matches(j).Length = 4 Then
							CodeVal = Val("&H" & Mid$(Matches(j).Value,3))
							ConvertBRegExp = Replace$(ConvertBRegExp,Matches(j).Value,Val2Bytes(CodeVal,2))
						End If
					ElseIf i = 1 Then
						If Matches(j).Length = 6 Then
							CodeVal = Val("&H" & Mid$(Matches(j).Value,3))
							ConvertBRegExp = Replace$(ConvertBRegExp,Matches(j).Value,Val2Bytes(CodeVal,2))
						End If
					ElseIf Matches(j).Length > 1 And Matches(j).Length < 5 Then
						CodeVal = Val("&O" & Replace$(Matches(j).Value,"\",""))
						If CodeVal > 256 Then
							Matches(j).Value = Left$(Matches(j).Value,3)
							CodeVal = Val("&O" & Replace$(Matches(j).Value,"\",""))
						End If
						ConvertBRegExp = Replace$(ConvertBRegExp,Matches(j).Value,Val2Bytes(CodeVal,2))
					End If
				Next j
			End If
		Next i
	End With
End Function


'转换数值为字节数组
Public Function Val2Bytes(ByVal Value As Long,ByVal Length As Integer) As Byte()
	On Error GoTo errHandle
	ReDim Bytes(Length - 1) As Byte
	CopyMemory Bytes(0), Value, Length
	Val2Bytes = Bytes
	Exit Function
	errHandle:
	ReDim Bytes(0) As Byte
	Val2Bytes = Bytes
End Function


'转换字符为整数数值
Public Function StrToLong(ByVal mStr As String,Optional ByVal DefaultValue As Long) As Long
	On Error GoTo errHandle
	StrToLong = CLng(mStr)
	Exit Function
	errHandle:
	StrToLong = DefaultValue
End Function


'自定义参数
Function Projects(ProjectID As Long) As Long
	Dim ActionList(2) As String,MsgList() As String
	Projects = ProjectID
	If getMsgList(UIDataList,MsgList,"Projects",1) = False Then Exit Function
	ActionList(0) = MsgList(23)
	ActionList(1) = MsgList(24)
	ActionList(2) = MsgList(25)
	Begin Dialog UserDialog 660,518,MsgList(0),.ProjectFunc ' %GRID:10,7,1,1
		Text 20,7,620,28,MsgList(1),.MainText

		GroupBox 20,42,360,77,MsgList(2),.GroupBox1
		DropListBox 40,63,290,21,ProjectList(),.ProjectList
		PushButton 330,63,30,21,MsgList(3),.NextSetType
		PushButton 40,91,90,21,MsgList(4),.AddButton
		PushButton 155,91,90,21,MsgList(5),.ChangButton
		PushButton 270,91,90,21,MsgList(6),.DelButton

		GroupBox 390,42,250,77,MsgList(7),.GroupBox2
		OptionGroup .cWriteType
			OptionButton 410,66,100,14,MsgList(8),.cWriteToFile
			OptionButton 520,66,100,14,MsgList(9),.cWriteToRegistry
		PushButton 410,91,90,21,MsgList(10),.ImportButton
		PushButton 530,91,90,21,MsgList(11),.ExportButton

		GroupBox 20,133,620,161,MsgList(12),.GroupBox3
		Text 40,154,160,14,MsgList(13),.ItemText
		Text 210,154,130,14,MsgList(14),.inSourceText
		Text 350,154,130,14,MsgList(15),.inTargetText
		Text 490,154,130,14,MsgList(16),.inBothText
		Text 40,178,160,14,MsgList(17),.PreSpaceText
		Text 40,199,160,14,MsgList(18),.EndSpaceText
		Text 40,220,160,14,MsgList(19),.acckeyText
		Text 40,241,160,14,MsgList(20),.EndStringText
		Text 40,262,160,14,MsgList(21),.ShortcutText
		CheckBox 210,175,120,21,MsgList(22),.PreSpaceinSourceBox
		CheckBox 350,175,120,21,MsgList(25),.PreSpaceinTargetBox
		DropListBox 490,175,130,21,ActionList(),.PreSpaceinBothBox
		CheckBox 210,196,120,21,MsgList(22),.EndSpaceinSourceBox
		CheckBox 350,196,120,21,MsgList(25),.EndSpaceinTargetBox
		DropListBox 490,196,130,21,ActionList(),.EndSpaceinBothBox
		CheckBox 210,217,120,21,MsgList(22),.AcckeyinSourceBox
		CheckBox 350,217,120,21,MsgList(25),.AcckeyinTargetBox
		DropListBox 490,217,130,21,ActionList(),.AcckeyinBothBox
		CheckBox 210,238,120,21,MsgList(22),.EndStringinSourceBox
		CheckBox 350,238,120,21,MsgList(25),.EndStringinTargetBox
		DropListBox 490,238,130,21,ActionList(),.EndStringinBothBox
		CheckBox 210,259,120,21,MsgList(22),.ShortcutinSourceBox
		CheckBox 350,259,120,21,MsgList(25),.ShortcutinTargetBox
		DropListBox 490,259,130,21,ActionList(),.ShortcutinBothBox

		GroupBox 20,308,620,168,MsgList(26),.GroupBox4
		CheckBox 40,329,580,14,MsgList(27),.DelSpaceBox
		CheckBox 40,350,580,14,MsgList(28),.EnableStringSplitBox
		CheckBox 40,371,580,14,MsgList(29),.TranEndStringBox
		CheckBox 40,392,580,14,MsgList(30),.AcckeyInShortcutBox
		CheckBox 40,420,580,14,MsgList(32),.ApplyCheckResultBox
		CheckBox 40,441,580,14,MsgList(31),.ShowOriginalTranBox

		PushButton 20,490,90,21,MsgList(33),.HelpButton
		PushButton 120,490,100,21,MsgList(34),.ResetButton
		PushButton 330,490,90,21,MsgList(35),.TestButton
		PushButton 230,490,90,21,MsgList(36),.CleanButton
		OKButton 460,490,90,21,.OKButton
		CancelButton 560,490,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.ProjectList = ProjectID
	If Dialog(dlg) = 0 Then Exit Function
	Projects = dlg.ProjectList
End Function


Private Function ProjectFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim HeaderID As Long,NewData As String
	Dim i As Long,j As Long,n As Long,CheckID As Long,Path As String,Temp As String
	Dim TempArray() As String,TempList() As String,MsgList() As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgEnable "ShowOriginalTranBox",False
		If CheckArray(ProjectList) = True Then
			HeaderID = DlgValue("ProjectList")
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),LngJoinStr)
			DlgValue "PreSpaceinSourceBox",StrToLong(SetsArray(0))
			DlgValue "PreSpaceinTargetBox",StrToLong(SetsArray(1))
			DlgValue "PreSpaceinBothBox",StrToLong(SetsArray(2))
			DlgValue "EndSpaceinSourceBox",StrToLong(SetsArray(3))
			DlgValue "EndSpaceinTargetBox",StrToLong(SetsArray(4))
			DlgValue "EndSpaceinBothBox",StrToLong(SetsArray(5))
			DlgValue "AcckeyinSourceBox",StrToLong(SetsArray(6))
			DlgValue "AcckeyinTargetBox",StrToLong(SetsArray(7))
			DlgValue "AcckeyinBothBox",StrToLong(SetsArray(8))
			DlgValue "EndStringinSourceBox",StrToLong(SetsArray(9))
			DlgValue "EndStringinTargetBox",StrToLong(SetsArray(10))
			DlgValue "EndStringinBothBox",StrToLong(SetsArray(11))
			DlgValue "ShortcutinSourceBox",StrToLong(SetsArray(12))
			DlgValue "ShortcutinTargetBox",StrToLong(SetsArray(13))
			DlgValue "ShortcutinBothBox",StrToLong(SetsArray(14))
			DlgValue "DelSpaceBox",StrToLong(SetsArray(15))
			DlgValue "EnableStringSplitBox",StrToLong(SetsArray(16))
			DlgValue "TranEndStringBox",StrToLong(SetsArray(17))
			DlgValue "AcckeyInShortcutBox",StrToLong(SetsArray(18))
			DlgValue "ShowOriginalTranBox",StrToLong(SetsArray(19))
			DlgValue "ApplyCheckResultBox",StrToLong(SetsArray(20))
		End If
		If cWriteLoc = CheckFilePath Then
			DlgValue "cWriteType",0
		ElseIf cWriteLoc = CheckRegKey Then
			DlgValue "cWriteType",1
		ElseIf cWriteLoc = "" Then
			DlgValue "cWriteType",0
		End If
		For i = LBound(DefaultProjectList) To UBound(DefaultProjectList)
			If DefaultProjectList(i) = ProjectList(DlgValue("ProjectList")) Then
				DlgEnable "ChangButton",False
				DlgEnable "DelButton",False
				Exit For
			End If
		Next i
		If DlgEnable("DelButton") = True Then
			If UBound(CheckList) = 0 Then DlgEnable "DelButton",False
		End If
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		ProjectFunc = True '防止按下按钮关闭对话框窗口
		If getMsgList(UIDataList,MsgList,"SettingsDlgFunc",1) = False Then Exit Function
		Select Case DlgItem$
		Case "HelpButton"
			Call Help("ProjectHelp")
			Exit Function
		Case "OKButton"
			Path = IIf(DlgValue("cWriteType") = 0,CheckFilePath,CheckRegKey)
			If WriteCheckSet(Path,"Project") = False Then
				MsgBox Replace$(MsgList(20),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
			Else
				ProjectListBak = ProjectList
				ProjectDataListBak = ProjectDataList
				ProjectFunc = False
			End If
			Exit Function
		Case "CancelButton"
			ProjectList = ProjectListBak
			ProjectDataList = ProjectDataListBak
			ProjectFunc = False
			Exit Function
		Case "NextSetType"
			i = DlgValue("ProjectList")
			If i < UBound(ProjectList) Then i = i + 1 Else i = 0
			DlgValue "ProjectList",i
			DlgItem$ = "ProjectList"
		Case "AddButton"
			NewData = AddSet(ProjectList)
			If NewData = "" Then Exit Function
			HeaderID = DlgValue("ProjectList")
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),LngJoinStr)
			ReDim SetsArray(UBound(SetsArray)) As String
			CreateArray(NewData,NewData & JoinStr & Join(SetsArray,LngJoinStr),ProjectList,ProjectDataList)
			DlgListBoxArray "ProjectList",ProjectList()
			DlgText "ProjectList",NewData
			DlgItem$ = "ProjectList"
		Case "ChangButton"
			HeaderID = DlgValue("ProjectList")
			NewData = EditSet(ProjectList,HeaderID)
			If NewData = "" Then Exit Function
			ProjectList(HeaderID) = NewData
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			MainArray(0) = NewData
			ProjectDataList(HeaderID) = Join(MainArray,JoinStr)
			DlgListBoxArray "ProjectList",ProjectList()
			DlgValue "ProjectList",HeaderID
			Exit Function
    	Case "DelButton"
			n = DlgValue("ProjectList")
			If MsgBox(Replace(MsgList(12),"%s",DlgText("ProjectList")),vbYesNo+vbInformation,MsgList(11)) = vbNo Then Exit Function
			i = UBound(ProjectList)
			Call DelArrays(ProjectList,ProjectDataList,n)
			If n > 0 And n = i Then n = n - 1
			DlgListBoxArray "ProjectList",ProjectList()
			DlgValue "ProjectList",n
			DlgItem$ = "ProjectList"
		Case "ResetButton"
			HeaderID = DlgValue("ProjectList")
			ReDim TempArray(1)
			For i = LBound(DefaultProjectList) To UBound(DefaultProjectList)
				If DefaultProjectList(i) = ProjectList(HeaderID) Then
					TempArray(0) = MsgList(1)
					Exit For
				End If
			Next i
			TempArray(1) = MsgList(2)
			For i = LBound(ProjectList) To UBound(ProjectList)
				If i <> HeaderID Then
					ReDim Preserve TempArray(i + 2)
					TempArray(i + 2) = MsgList(3) & " - " & ProjectList(i)
				End If
			Next i
			i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),LngJoinStr)
			If i = 0 Then
				SetsArray = ReSplit(CheckSettings(ProjectList(HeaderID),1),LngJoinStr)
			ElseIf i = 1 Then
				For n = LBound(ProjectDataListBak) To UBound(ProjectDataListBak)
					TempArray = ReSplit(ProjectDataListBak(n),JoinStr)
					If TempArray(0) = ProjectList(HeaderID) Then
						SetsArray = ReSplit(TempArray(1),LngJoinStr)
						Exit For
					End If
				Next n
			ElseIf i >= 2 Then
				Temp = Mid(TempArray(i),InStr(TempArray(i),MsgList(3) & " - ") + Len(MsgList(3) & " - "))
				For n = LBound(ProjectList) To UBound(ProjectList)
					If ProjectList(n) = Temp Then
						SetsArray = ReSplit(ReSplit(ProjectDataList(n),JoinStr)(1),LngJoinStr)
						Exit For
					End If
				Next n
			End If
			MainArray(1) = Join(SetsArray,LngJoinStr)
			ProjectDataList(HeaderID) = Join(MainArray,JoinStr)
			DlgItem$ = "ProjectList"
	    Case "CleanButton"
	    	HeaderID = DlgValue("ProjectList")
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),LngJoinStr)
			ReDim SetsArray(UBound(SetsArray)) As String
			MainArray(1) = Join(SetsArray,LngJoinStr)
			ProjectDataList(HeaderID) = Join(MainArray,JoinStr)
			DlgItem$ = "ProjectList"
		Case "ImportButton"
			If PSL.SelectFile(Path,True,Replace$(MsgList(25),"%s",MsgList(15)),MsgList(23)) = False Then Exit Function
			n = GetCheckSet("Project",Path)
			If n = 3 Then
				DlgListBoxArray "ProjectList",ProjectList()
				Header = ProjectList(UBound(ProjectList))
				HeaderID = -1
				For i = LBound(ProjectListBak) To UBound(ProjectListBak)
					If ProjectListBak(i) = Header Then
						HeaderID = i
						Exit For
					End If
				Next i
				If HeaderID < 0 Then HeaderID = UBound(ProjectList)
				DlgValue "ProjectList",HeaderID
				MsgBox MsgList(19),vbOkOnly+vbInformation,MsgList(17)
				DlgItem$ = "ProjectList"
			ElseIf n = 0 Then
				MsgBox Replace$(MsgList(21),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
		Case "ExportButton"
			If PSL.SelectFile(Path,False,Replace$(MsgList(25),"%s",MsgList(15)),MsgList(24)) = True Then
				If InStr(Path,".dat") = 0 Then Path = Path & ".dat"
				If WriteCheckSet(Path,"All") = False Then
					MsgBox Replace$(MsgList(22),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
				Else
					MsgBox MsgList(18),vbOkOnly+vbInformation,MsgList(17)
				End If
			End If
			Exit Function
		Case "TestButton"
			If CheckNullData("",CheckDataList,"1,4,14-17",1) = True Then
				If MsgBox(Replace$(MsgList(10),"%s",MsgList(15)) & MsgList(6),vbYesNo+vbInformation,MsgList(5)) = vbNo Then
					Exit Function
				End If
			End If
			'转换检查配置中的转义符
			TempArray = CheckDataList
			For i = LBound(CheckDataList) To UBound(CheckDataList)
				MainArray = ReSplit(CheckDataList(i),JoinStr)
				SetsArray = ReSplit(MainArray(1),SubJoinStr)
				For j = 0 To UBound(SetsArray)
					If j <> 4 And j <> 14 And j <> 15 And j < 18 Then
						If SetsArray(j) <> "" Then
							If j = 1 Or j = 5 Or j = 13 Or j = 16 Or j = 17 Then
								If j = 5 Or j = 7 Then Temp = " " Else Temp = ","
								TempList = ReSplit(SetsArray(j),Temp,-1)
								Call SortArrayByLength(TempList,0,UBound(TempList),True)
								SetsArray(j) = Convert(Join(TempList,Temp))
							Else
								SetsArray(j) = Convert(SetsArray(j))
							End If
						End If
					End If
				Next j
				MainArray(1) = Join(SetsArray,SubJoinStr)
				CheckDataList(i) = Join(MainArray,JoinStr)
			Next i
			For i = LBound(CheckDataList) To UBound(CheckDataList)
				If ReSplit(CheckDataList(i),JoinStr)(0) = cSelected(2) Then
					CheckID = i
					Exit For
				End If
			Next i
			TempList = CheckDataListBak
			CheckDataListBak = TempArray
			Call CheckTest(CheckID,HeaderID)
			CheckDataList = TempArray
			CheckDataListBak = TempList
			Exit Function
		End Select

		HeaderID = DlgValue("ProjectList")
		If HeaderID < 0 Then Exit Function
		If DlgItem$ = "ProjectList" Then
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),LngJoinStr)
			DlgValue "PreSpaceinSourceBox",StrToLong(SetsArray(0))
			DlgValue "PreSpaceinTargetBox",StrToLong(SetsArray(1))
			DlgValue "PreSpaceinBothBox",StrToLong(SetsArray(2))
			DlgValue "EndSpaceinSourceBox",StrToLong(SetsArray(3))
			DlgValue "EndSpaceinTargetBox",StrToLong(SetsArray(4))
			DlgValue "EndSpaceinBothBox",StrToLong(SetsArray(5))
			DlgValue "AcckeyinSourceBox",StrToLong(SetsArray(6))
			DlgValue "AcckeyinTargetBox",StrToLong(SetsArray(7))
			DlgValue "AcckeyinBothBox",StrToLong(SetsArray(8))
			DlgValue "EndStringinSourceBox",StrToLong(SetsArray(9))
			DlgValue "EndStringinTargetBox",StrToLong(SetsArray(10))
			DlgValue "EndStringinBothBox",StrToLong(SetsArray(11))
			DlgValue "ShortcutinSourceBox",StrToLong(SetsArray(12))
			DlgValue "ShortcutinTargetBox",StrToLong(SetsArray(13))
			DlgValue "ShortcutinBothBox",StrToLong(SetsArray(14))
			DlgValue "DelSpaceBox",StrToLong(SetsArray(15))
			DlgValue "EnableStringSplitBox",StrToLong(SetsArray(16))
			DlgValue "TranEndStringBox",StrToLong(SetsArray(17))
			DlgValue "AcckeyInShortcutBox",StrToLong(SetsArray(18))
			DlgValue "ShowOriginalTranBox",StrToLong(SetsArray(19))
			DlgValue "ApplyCheckResultBox",StrToLong(SetsArray(20))
		Else
			MainArray = ReSplit(ProjectDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),LngJoinStr)
			SetsArray(0) = DlgValue("PreSpaceinSourceBox")
			SetsArray(1) = DlgValue("PreSpaceinTargetBox")
			SetsArray(2) = DlgValue("PreSpaceinBothBox")
			SetsArray(3) = DlgValue("EndSpaceinSourceBox")
			SetsArray(4) = DlgValue("EndSpaceinTargetBox")
			SetsArray(5) = DlgValue("EndSpaceinBothBox")
			SetsArray(6) = DlgValue("AcckeyinSourceBox")
			SetsArray(7) = DlgValue("AcckeyinTargetBox")
			SetsArray(8) = DlgValue("AcckeyinBothBox")
			SetsArray(9) = DlgValue("EndStringinSourceBox")
			SetsArray(10) = DlgValue("EndStringinTargetBox")
			SetsArray(11) = DlgValue("EndStringinBothBox")
			SetsArray(12) = DlgValue("ShortcutinSourceBox")
			SetsArray(13) = DlgValue("ShortcutinTargetBox")
			SetsArray(14) = DlgValue("ShortcutinBothBox")
			SetsArray(15) = DlgValue("DelSpaceBox")
			SetsArray(16) = DlgValue("EnableStringSplitBox")
			SetsArray(17) = DlgValue("TranEndStringBox")
			SetsArray(18) = DlgValue("AcckeyInShortcutBox")
			SetsArray(19) = DlgValue("ShowOriginalTranBox")
			SetsArray(20) = DlgValue("ApplyCheckResultBox")
			MainArray(1) = Join(SetsArray,LngJoinStr)
			ProjectDataList(HeaderID) = Join(MainArray,JoinStr)
		End If
		j = 0
		For i = LBound(DefaultProjectList) To UBound(DefaultProjectList)
			If DefaultProjectList(i) = ProjectList(HeaderID) Then
				j = j + 1
				Exit For
			End If
		Next i
		If j > 0 Then
			DlgEnable "ChangButton",False
			DlgEnable "DelButton",False
		Else
			DlgEnable "ChangButton",True
			DlgEnable "DelButton",IIf(UBound(ProjectList) = 0,False,True)
		End If
	End Select
End Function


'自定义参数
Function Settings(EngineID As Long,CheckID As Long,OptionID As Long) As Long
	Dim MsgList() As String,TempList() As String
	Settings = EngineID
	If getMsgList(UIDataList,MsgList,"Settings",1) = False Then Exit Function
	Begin Dialog UserDialog 660,518,MsgList(0),.SettingsDlgFunc ' %GRID:10,7,1,1
		TextBox 0,0,0,21,.SuppValueBox
		Text 20,7,620,28,MsgList(1),.MainText
		OptionGroup .Options
			OptionButton 80,42,130,14,MsgList(75),.TrnEngine
			OptionButton 220,42,130,14,MsgList(2),.StrHandle
			OptionButton 360,42,130,14,MsgList(3),.AutoUpdate
			OptionButton 500,42,130,14,MsgList(4),.UILangListSet

		GroupBox 20,70,360,77,MsgList(5),.GroupBox1
		DropListBox 40,91,290,21,EngineList(),.EngineList
		DropListBox 40,91,290,21,CheckList(),.CheckList
		PushButton 330,91,30,21,MsgList(6),.LevelButton
		PushButton 40,119,90,21,MsgList(7),.AddButton
		PushButton 155,119,90,21,MsgList(8),.ChangButton
		PushButton 270,119,90,21,MsgList(9),.DelButton

		GroupBox 390,70,250,77,MsgList(10),.GroupBox2
		OptionGroup .tWriteType
			OptionButton 410,94,110,14,MsgList(11),.tWriteToFile
			OptionButton 520,94,110,14,MsgList(12),.tWriteToRegistry
		OptionGroup .cWriteType
			OptionButton 410,94,100,14,MsgList(11),.cWriteToFile
			OptionButton 520,94,100,14,MsgList(12),.cWriteToRegistry
		PushButton 410,119,90,21,MsgList(13),.ImportButton
		PushButton 530,119,90,21,MsgList(14),.ExportButton

		GroupBox 20,189,620,287,MsgList(15),.GroupBox4
		OptionGroup .Engine
			OptionButton 110,161,160,14,MsgList(76),.EngineArgument
			OptionButton 280,161,160,14,MsgList(77),.LangCodePair
			OptionButton 450,161,160,14,MsgList(78),.EngineEnable

		Text 40,213,120,14,MsgList(79),.ObjectNameText
		Text 40,234,120,14,MsgList(80),.AppIDText
		Text 40,255,120,14,MsgList(81),.UrlText
		Text 40,276,120,14,MsgList(82),.UrlTemplateText
		Text 40,297,120,14,MsgList(83),.bstrMethodText
		Text 340,297,120,14,MsgList(84),.varAsyncText
		Text 40,318,120,14,MsgList(85),.bstrUserText
		Text 340,318,120,14,MsgList(86),.bstrPasswordText
		Text 40,339,120,14,MsgList(87),.varBodyText
		Text 40,360,120,14,MsgList(88),.setRequestHeaderText
		Text 40,378,120,21,MsgList(89),.setRequestHeaderText2
		Text 40,402,120,14,MsgList(90),.responseTypeText
		Text 40,423,120,14,MsgList(91),.TranBeforeStrText
		Text 40,444,120,14,MsgList(92),.TranAfterStrText
		TextBox 170,210,450,21,.ObjectNameBox
		TextBox 170,231,450,21,.AppIDBox
		TextBox 170,252,450,21,.UrlBox
		TextBox 170,273,420,21,.UrlTemplateBox
		TextBox 170,294,130,21,.bstrMethodBox
		TextBox 460,294,130,21,.varAsyncBox
		TextBox 170,315,160,21,.bstrUserBox
		TextBox 460,315,160,21,.bstrPasswordBox,-1
		TextBox 170,336,420,21,.varBodyBox
		TextBox 170,357,420,42,.setRequestHeaderBox,1
		TextBox 170,399,420,21,.responseTypeBox
		TextBox 170,420,420,21,.TranBeforeStrBox
		TextBox 170,441,420,21,.TranAfterStrBox
		PushButton 590,273,30,21,MsgList(69),.UrlTemplateButton
		PushButton 300,294,30,21,MsgList(69),.bstrMethodButton
		PushButton 590,294,30,21,MsgList(69),.varAsyncButton
		PushButton 590,336,30,21,MsgList(69),.varBodyButton
		PushButton 590,357,30,21,MsgList(69),.RequestButton
		PushButton 590,399,30,21,MsgList(69),.responseTypeButton
		PushButton 590,420,30,21,MsgList(69),.TranBeforeStrButton
		PushButton 590,441,30,21,MsgList(69),.TranAfterStrButton

		Text 40,203,220,14,MsgList(93),.LngNameText
		Text 270,203,100,14,MsgList(94),.SrcLngText
		Text 380,203,100,14,MsgList(95),.TranLngText
		ListBox 40,217,220,245,TempList(),.LngNameList
		ListBox 270,217,100,245,TempList(),.SrcLngList
		ListBox 380,217,100,245,TempList(),.TranLngList
		PushButton 500,217,120,21,MsgList(07),.AddLngButton
		PushButton 500,238,120,21,MsgList(09),.DelLngButton
		PushButton 500,259,120,21,MsgList(96),.DelAllButton
		PushButton 500,294,120,21,MsgList(97),.EditLngButton
		PushButton 500,315,120,21,MsgList(98),.ExtEditButton
		PushButton 500,336,120,21,MsgList(99),.NullLngButton
		PushButton 500,357,120,21,MsgList(100),.ResetLngButton
		PushButton 500,399,120,21,MsgList(101),.ShowNoNullLngButton
		PushButton 500,420,120,21,MsgList(102),.ShowNullLngButton
		PushButton 500,441,120,21,MsgList(103),.ShowAllLngButton

		Text 40,210,580,70,MsgList(104),.EnableText1
		Text 40,287,580,70,MsgList(105),.EnableText2
		CheckBox 220,378,330,14,MsgList(106),.EnableBox

		GroupBox 20,189,620,287,MsgList(15),.GroupBox3
		Text 30,164,190,14,MsgList(16),.SetItemText,1
		DropListBox 240,161,250,21,TempList(),.SetType
		PushButton 490,161,30,21,MsgList(17),.NextSetType

		Text 40,210,580,14,MsgList(18),.AccKeyBoxTxt
		Text 40,259,580,14,MsgList(19),.ExCrBoxTxt
		Text 40,308,580,14,MsgList(20),.ChkBktBoxTxt
		Text 40,357,580,14,MsgList(21),.KpPairBoxTxt
		TextBox 40,231,580,21,.AccKeyBox
		TextBox 40,280,580,21,.ExCrBox,1
		TextBox 40,329,580,21,.ChkBktBox
		TextBox 40,378,580,42,.KpPairBox,1
		CheckBox 40,427,580,14,MsgList(22),.AsiaKeyBox
		CheckBox 40,448,580,14,MsgList(23),.AddAcckeyBox

		Text 40,210,580,14,MsgList(24),.ChkEndBoxTxt
		Text 40,294,580,14,MsgList(25),.NoTrnEndBoxTxt
		Text 40,378,580,14,MsgList(26),.AutoTrnEndBoxTxt
		TextBox 40,231,580,56,.ChkEndBox,1
		TextBox 40,315,580,56,.NoTrnEndBox,1
		TextBox 40,399,580,63,.AutoTrnEndBox,1

		Text 40,210,580,14,MsgList(27),.ShortBoxTxt
		Text 40,259,580,14,MsgList(28),.ShortKeyBoxTxt
		Text 40,371,580,14,MsgList(29),.KpShortKeyBoxTxt
		TextBox 40,231,580,21,.ShortBox
		TextBox 40,280,580,84,.ShortKeyBox,1
		TextBox 40,392,580,70,.KpShortKeyBox,1

		Text 40,210,580,14,MsgList(30),.PreRepStrBoxTxt
		Text 40,336,580,14,MsgList(31),.AutoWebFlagBoxTxt
		TextBox 40,231,580,98,.PreRepStrBox,1
		TextBox 40,357,580,105,.AutoWebFlagBox,1

		Text 40,210,580,77,MsgList(32),.LineSplitBoxTxt
		Text 40,294,580,14,MsgList(33),.PreInsertSplitBoxTxt
		Text 40,343,580,14,MsgList(34),.AppInsertSplitBoxTxt
		Text 40,392,580,14,MsgList(35),.ReplaceSplitBoxTxt
		TextBox 40,315,580,21,.PreInsertSplitBox
		TextBox 40,364,580,21,.AppInsertSplitBox
		TextBox 40,413,580,21,.ReplaceSplitBox
		CheckBox 40,448,580,14,MsgList(36),.LineSplitModeBox

		Text 40,203,200,14,MsgList(37),.AppLngText
		Text 420,203,200,14,MsgList(38),.UseLngText
		ListBox 40,217,200,245,TempList(),.AppLngList
		ListBox 420,217,200,245,TempList(),.UseLngList
		PushButton 260,217,140,21,MsgList(39),.AddLangButton
		PushButton 260,238,140,21,MsgList(40),.AddAllLangButton
		PushButton 260,266,140,21,MsgList(41),.DelLangButton
		PushButton 260,287,140,21,MsgList(42),.DelAllLangButton
		PushButton 260,322,140,21,MsgList(43),.SetAppLangButton
		PushButton 260,343,140,21,MsgList(44),.EditAppLangButton
		PushButton 260,364,140,21,MsgList(45),.DelAppLangButton
		PushButton 260,399,140,21,MsgList(46),.SetUseLangButton
		PushButton 260,420,140,21,MsgList(47),.EditUseLangButton
		PushButton 260,441,140,21,MsgList(48),.DelUseLangButton

		GroupBox 20,70,620,91,MsgList(49),.UpdateSetGroup
		OptionGroup .UpdateSet
			OptionButton 40,91,580,14,MsgList(50),.AutoButton
			OptionButton 40,112,580,14,MsgList(51),.ManualButton
			OptionButton 40,133,580,14,MsgList(52),.OffButton
		GroupBox 20,175,620,49,MsgList(53),.CheckGroup
		Text 40,196,100,14,MsgList(54),.UpdateCycleText
		TextBox 150,192,40,21,.UpdateCycleBox
		Text 200,196,60,14,MsgList(55),.UpdateDatesText
		Text 280,196,130,14,MsgList(56),.UpdateDateText
		TextBox 420,192,100,21,.UpdateDateBox
		PushButton 540,192,80,21,MsgList(57),.CheckButton
		GroupBox 20,238,620,98,MsgList(58),.WebSiteGroup
		TextBox 40,259,580,63,.WebSiteBox,1
		GroupBox 20,350,620,126,MsgList(59),.CmdGroup
		Text 40,371,550,14,MsgList(60),.CmdPathBoxText
		Text 40,420,550,14,MsgList(61),.ArgumentBoxText
		TextBox 40,392,550,21,.CmdPathBox
		TextBox 40,441,550,21,.ArgumentBox
		PushButton 590,392,30,21,MsgList(6),.ExeBrowseButton
		PushButton 590,441,30,21,MsgList(62),.ArgumentButton

		GroupBox 20,70,620,280,MsgList(63),.UILangSetGroup
		Text 40,91,580,147,MsgList(64),.UILangSetText1
		Text 40,255,130,14,MsgList(65),.UILangSetText2
		DropListBox 180,252,360,21,TempList(),.UILangList
		Text 40,294,580,42,MsgList(66),.UILangSetText3

		GroupBox 20,364,620,112,MsgList(67),.UIFontSetGroup
		Text 40,388,130,14,MsgList(68),.MainFontText
		TextBox 180,385,360,21,.MainFontBox
		PushButton 540,385,30,21,MsgList(69),.MainFontButton

		PushButton 20,490,90,21,MsgList(70),.HelpButton
		PushButton 120,490,100,21,MsgList(71),.ResetButton
		PushButton 330,490,90,21,MsgList(72),.TestButton
		PushButton 230,490,90,21,MsgList(73),.CleanButton
		PushButton 120,490,160,21,MsgList(74),.EditUILangButton
		OKButton 460,490,90,21,.OKButton
		CancelButton 560,490,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.EngineList = EngineID
	dlg.CheckList = CheckID
	dlg.Options = OptionID
	If Dialog(dlg) = 0 Then Exit Function
	Settings = dlg.EngineList
End Function


'请务必查看对话框帮助主题以了解更多信息。
Private Function SettingsDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,n As Long,HeaderID As Long
	Dim NewData As String,Path As String,TempDataList() As String
	Dim Header As String,Temp As String,LngName As String,TempList() As String,TempArray() As String
	Dim AppLngList() As String,UseLngList() As String,MsgList() As String
	Dim SrcLngCode As String,TranLngCode As String,tStemp As Boolean,cStemp As Boolean
	Dim LngNameList() As String,SrcLngList() As String,TranLngList() As String,Dic As Object
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		If getMsgList(UIDataList,MsgList,"SettingsDlgFunc",1) = False Then Exit Function
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgEnable "MainFontBox",False

		ReDim TempList(5) As String
		For i = 0 To 5
			TempList(i) = MsgList(i + 53)
		Next i
		DlgListBoxArray "SetType",TempList()
		DlgValue "SetType",0

		ReDim TempList(UBound(UIFileList) + 2) As String
		n = 1
		For i = 0 To UBound(UIFileList) + 2
			If i < 2 Then
				TempList(i) = MsgList(i + 59)
			ElseIf UIFileList(i - 2).FilePath <> "" Then
				TempList(i) = UIFileList(i - 2).LangName
				If HeaderID = 0 Then
					If UIFileList(i - 2).FilePath = LangFile Then HeaderID = i
				End If
				n = n + 1
			End If
		Next i
		ReDim Preserve TempList(n) As String
		If tSelected(0) = "" Or tSelected(0) = "0" Then
			HeaderID = 0
		ElseIf tSelected(0) = "1" Then
			HeaderID = 1
		End If
		DlgListBoxArray "UILangList",TempList()
		DlgValue "UILangList",HeaderID
		DlgText "MainFontBox",GetFontText(SuppValue,LFList(0))

		If CheckArray(EngineList) = True Then
			MainArray = ReSplit(EngineDataList(DlgValue("EngineList")),JoinStr)
			SetsArray = ReSplit(MainArray(1),SubJoinStr)
			DlgText "ObjectNameBox",SetsArray(0)
			DlgText "AppIDBox",SetsArray(1)
			DlgText "UrlBox",SetsArray(2)
			DlgText "UrlTemplateBox",SetsArray(3)
			DlgText "bstrMethodBox",SetsArray(4)
			DlgText "varAsyncBox",SetsArray(5)
			DlgText "bstrUserBox",SetsArray(6)
			DlgText "bstrPasswordBox",SetsArray(7)
			DlgText "varBodyBox",SetsArray(8)
			DlgText "setRequestHeaderBox",SetsArray(9)
			DlgText "responseTypeBox",SetsArray(10)
			Select Case DlgText("responseTypeBox")
			Case "responseText"
				DlgText "TranBeforeStrBox",SetsArray(11)
				DlgText "TranAfterStrBox",SetsArray(12)
			Case "responseBody"
				DlgText "TranBeforeStrBox",SetsArray(13)
				DlgText "TranAfterStrBox",SetsArray(14)
			Case "responseStream"
				DlgText "TranBeforeStrBox",SetsArray(15)
				DlgText "TranAfterStrBox",SetsArray(16)
			Case "responseXML"
				DlgText "TranBeforeStrBox",SetsArray(17)
				DlgText "TranAfterStrBox",SetsArray(18)
			End Select
			DlgValue "EnableBox",StrToLong(SetsArray(19))
			If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject
			SplitData(MainArray(2),LngNameList,SrcLngList,TranLngList)
			DlgListBoxArray "LngNameList",LngNameList()
			DlgListBoxArray "SrcLngList",SrcLngList()
			DlgListBoxArray "TranLngList",TranLngList()
			DlgValue "LngNameList",0
			DlgValue "SrcLngList",0
			DlgValue "TranLngList",0
		End If

		If CheckArray(CheckList) = True Then
			MainArray = ReSplit(CheckDataList(DlgValue("CheckList")),JoinStr)
			SetsArray = ReSplit(MainArray(1),SubJoinStr)
			DlgText "ExCrBox",SetsArray(0)
			DlgText "PreInsertSplitBox",SetsArray(1)
			DlgText "ChkBktBox",SetsArray(2)
			DlgText "KpPairBox",SetsArray(3)
			DlgValue "AsiaKeyBox",StrToLong(SetsArray(4))
			DlgText "ChkEndBox",SetsArray(5)
			DlgText "NoTrnEndBox",SetsArray(6)
			DlgText "AutoTrnEndBox",SetsArray(7)
			DlgText "ShortBox",SetsArray(8)
			DlgText "ShortKeyBox",SetsArray(9)
			DlgText "KpShortKeyBox",SetsArray(10)
			DlgText "PreRepStrBox",SetsArray(11)
			DlgText "AutoWebFlagBox",SetsArray(12)
			DlgText "AccKeyBox",SetsArray(13)
			DlgValue "AddAcckeyBox",StrToLong(SetsArray(14))
			DlgValue "LineSplitModeBox",StrToLong(SetsArray(15))
			DlgText "AppInsertSplitBox",SetsArray(16)
			DlgText "ReplaceSplitBox",SetsArray(17)
			getLngNameList(MainArray(2),AppLngList,UseLngList)
			DlgListBoxArray "AppLngList",AppLngList()
			DlgListBoxArray "UseLngList",UseLngList()
			DlgValue "AppLngList",0
			DlgValue "UseLngList",0
		End If

		If tWriteLoc = "" Or tWriteLoc = EngineFilePath Then
			DlgValue "tWriteType",0
		ElseIf tWriteLoc = EngineRegKey Then
			DlgValue "tWriteType",1
		End If
		If cWriteLoc = "" Or cWriteLoc = CheckFilePath Then
			DlgValue "cWriteType",0
		ElseIf cWriteLoc = CheckRegKey Then
			DlgValue "cWriteType",1
		End If

		If DlgText("LngNameList") = "" Then
			DlgText "LngNameText",Replace$(MsgList(50),"%s","0")
		Else
			DlgText "LngNameText",Replace$(MsgList(50),"%s",CStr$(UBound(LngNameList) + 1))
		End If
		If DlgText("AppLngList") = "" Then
			DlgText "AppLngText",Replace$(MsgList(31),"%s","0")
		Else
			DlgText "AppLngText",Replace$(MsgList(31),"%s",CStr$(UBound(AppLngList) + 1))
		End If
		If DlgText("UseLngList") = "" Then
			DlgText "UseLngText",Replace$(MsgList(32),"%s","0")
		Else
			DlgText "UseLngText",Replace$(MsgList(32),"%s",CStr$(UBound(UseLngList) + 1))
		End If

		If CheckArray(tUpdateSet) = True Then
			DlgValue "UpdateSet",StrToLong(tUpdateSet(0))
			DlgText "WebSiteBox",tUpdateSet(1)
			DlgText "CmdPathBox",tUpdateSet(2)
			DlgText "ArgumentBox",tUpdateSet(3)
			DlgText "UpdateCycleBox",tUpdateSet(4)
			DlgText "UpdateDateBox",tUpdateSet(5)
		End If
		If DlgText("UpdateDateBox") = "" Then DlgText "UpdateDateBox",MsgList(4)
		DlgEnable "UpdateDateBox",False

		If DlgValue("Options") = 0 Then
			tStemp = False
			For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
				If DefaultEngineList(i) = EngineList(DlgValue("EngineList")) Then
					tStemp = True
					Exit For
				End If
			Next i
			If tStemp = True Then
				DlgEnable "ChangButton",False
				DlgEnable "DelButton",False
			Else
				If UBound(EngineList) = 0 Then DlgEnable "DelButton",False
			End If
			If DlgText("responseTypeBox") = "responseXML" Then
				DlgText "TranBeforeStrText",MsgList(29)
				DlgText "TranAfterStrText",MsgList(30)
			Else
				DlgText "TranBeforeStrText",MsgList(27)
				DlgText "TranAfterStrText",MsgList(28)
			End If
			DlgEnable "ShowAllLngButton",False
			DlgEnable "ResetLngButton",False
		ElseIf DlgValue("Options") = 1 Then
			If DlgText("AppLngList") = "" Then
				DlgEnable "AddLangButton",False
				DlgEnable "AddAllLangButton",False
				DlgEnable "EditAppLangButton",False
				DlgEnable "DelAppLangButton",False
			End If
			If DlgText("UseLngList") = "" Then
				DlgEnable "DelLangButton",False
				DlgEnable "DelAllLangButton",False
				DlgEnable "EditUseLangButton",False
				DlgEnable "DelUseLangButton",False
			End If
			If DlgValue("AsiaKeyBox") = 1 Then
				DlgEnable "AddAcckeyBox",False
				DlgValue "AddAcckeyBox",0
			End If
			cStemp = False
			For i = LBound(CheckList) To UBound(CheckList)
				If CheckList(i) = CheckList(DlgValue("CheckList")) Then
					cStemp = True
					Exit For
				End If
			Next i
			If cStemp = True Then
				DlgEnable "ChangButton",False
				DlgEnable "DelButton",False
			Else
				If UBound(CheckList) = 0 Then DlgEnable "DelButton",False
			End If
		End If

		If DlgValue("Options") > 1 Then
			DlgVisible "GroupBox1",False
			DlgVisible "LevelButton",False
			DlgVisible "AddButton",False
			DlgVisible "ChangButton",False
			DlgVisible "DelButton",False

			DlgVisible "GroupBox2",False
			DlgVisible "ImportButton",False
			DlgVisible "ExportButton",False
		End If
		If DlgValue("Options") > 2 Then
			DlgVisible "ResetButton",False
			DlgVisible "TestButton",False
			DlgVisible "CleanButton",False
		Else
			DlgVisible "EditUILangButton",False
		End If
		If DlgValue("Options") <> 0 Then
			DlgVisible "EngineList",False

			DlgVisible "tWriteType",False
			DlgVisible "GroupBox4",False
			DlgVisible "Engine",False

			DlgVisible "ObjectNameText",False
			DlgVisible "AppIDText",False
			DlgVisible "UrlText",False
			DlgVisible "UrlTemplateText",False
			DlgVisible "bstrMethodText",False
			DlgVisible "varAsyncText",False
			DlgVisible "bstrUserText",False
			DlgVisible "bstrPasswordText",False
			DlgVisible "varBodyText",False
			DlgVisible "setRequestHeaderText",False
			DlgVisible "setRequestHeaderText2",False
			DlgVisible "responseTypeText",False
			DlgVisible "TranBeforeStrText",False
			DlgVisible "TranAfterStrText",False
			DlgVisible "ObjectNameBox",False
			DlgVisible "AppIDBox",False
			DlgVisible "UrlBox",False
			DlgVisible "UrlTemplateBox",False
			DlgVisible "bstrMethodBox",False
			DlgVisible "varAsyncBox",False
			DlgVisible "bstrUserBox",False
			DlgVisible "bstrPasswordBox",False
			DlgVisible "varBodyBox",False
			DlgVisible "setRequestHeaderBox",False
			DlgVisible "responseTypeBox",False
			DlgVisible "TranBeforeStrBox",False
			DlgVisible "TranAfterStrBox",False
			DlgVisible "UrlTemplateButton",False
			DlgVisible "bstrMethodButton",False
			DlgVisible "varAsyncButton",False
			DlgVisible "varBodyButton",False
			DlgVisible "RequestButton",False
			DlgVisible "responseTypeButton",False
			DlgVisible "TranBeforeStrButton",False
			DlgVisible "TranAfterStrButton",False

			DlgVisible "LngNameText",False
			DlgVisible "SrcLngText",False
			DlgVisible "TranLngText",False
			DlgVisible "LngNameList",False
			DlgVisible "SrcLngList",False
			DlgVisible "TranLngList",False
			DlgVisible "AddLngButton",False
			DlgVisible "DelLngButton",False
			DlgVisible "DelAllButton",False
			DlgVisible "EditLngButton",False
			DlgVisible "ExtEditButton",False
			DlgVisible "NullLngButton",False
			DlgVisible "ResetLngButton",False
			DlgVisible "ShowNoNullLngButton",False
			DlgVisible "ShowNullLngButton",False
			DlgVisible "ShowAllLngButton",False

			DlgVisible "EnableText1",False
			DlgVisible "EnableText2",False
			DlgVisible "EnableBox",False
		Else
			If DlgValue("Engine") <> 0 Then
				DlgVisible "ObjectNameText",False
				DlgVisible "AppIDText",False
				DlgVisible "UrlText",False
				DlgVisible "UrlTemplateText",False
				DlgVisible "bstrMethodText",False
				DlgVisible "varAsyncText",False
				DlgVisible "bstrUserText",False
				DlgVisible "bstrPasswordText",False
				DlgVisible "varBodyText",False
				DlgVisible "setRequestHeaderText",False
				DlgVisible "setRequestHeaderText2",False
				DlgVisible "responseTypeText",False
				DlgVisible "TranBeforeStrText",False
				DlgVisible "TranAfterStrText",False
				DlgVisible "ObjectNameBox",False
				DlgVisible "AppIDBox",False
				DlgVisible "UrlBox",False
				DlgVisible "UrlTemplateBox",False
				DlgVisible "bstrMethodBox",False
				DlgVisible "varAsyncBox",False
				DlgVisible "bstrUserBox",False
				DlgVisible "bstrPasswordBox",False
				DlgVisible "varBodyBox",False
				DlgVisible "setRequestHeaderBox",False
				DlgVisible "responseTypeBox",False
				DlgVisible "TranBeforeStrBox",False
				DlgVisible "TranAfterStrBox",False
				DlgVisible "UrlTemplateButton",False
				DlgVisible "bstrMethodButton",False
				DlgVisible "varAsyncButton",False
				DlgVisible "varBodyButton",False
				DlgVisible "RequestButton",False
				DlgVisible "responseTypeButton",False
				DlgVisible "TranBeforeStrButton",False
				DlgVisible "TranAfterStrButton",False
			End If
			If DlgValue("Engine") <> 1 Then
				DlgVisible "LngNameText",False
				DlgVisible "SrcLngText",False
				DlgVisible "TranLngText",False
				DlgVisible "LngNameList",False
				DlgVisible "SrcLngList",False
				DlgVisible "TranLngList",False
				DlgVisible "AddLngButton",False
				DlgVisible "DelLngButton",False
				DlgVisible "DelAllButton",False
				DlgVisible "EditLngButton",False
				DlgVisible "ExtEditButton",False
				DlgVisible "NullLngButton",False
				DlgVisible "ResetLngButton",False
				DlgVisible "ShowNoNullLngButton",False
				DlgVisible "ShowNullLngButton",False
				DlgVisible "ShowAllLngButton",False
			End If
			If DlgValue("Engine") <> 2 Then
				DlgVisible "EnableText1",False
				DlgVisible "EnableText2",False
				DlgVisible "EnableBox",False
			End If
		End If
		If DlgValue("Options") <> 1 Then
			DlgVisible "CheckList",False
			DlgVisible "cWriteType",False

			DlgVisible "GroupBox3",False
			DlgVisible "SetType",False
			DlgVisible "NextSetType",False
			DlgVisible "SetItemText",False

			DlgVisible "AccKeyBoxTxt",False
			DlgVisible "ExCrBoxTxt",False
			DlgVisible "ChkBktBoxTxt",False
			DlgVisible "KpPairBoxTxt",False
			DlgVisible "AccKeyBox",False
			DlgVisible "ExCrBox",False
			DlgVisible "ChkBktBox",False
			DlgVisible "KpPairBox",False
			DlgVisible "AsiaKeyBox",False
			DlgVisible "AddAcckeyBox",False

			DlgVisible "ChkEndBoxTxt",False
			DlgVisible "NoTrnEndBoxTxt",False
			DlgVisible "AutoTrnEndBoxTxt",False
			DlgVisible "ChkEndBox",False
			DlgVisible "NoTrnEndBox",False
			DlgVisible "AutoTrnEndBox",False

			DlgVisible "ShortBoxTxt",False
			DlgVisible "ShortKeyBoxTxt",False
			DlgVisible "KpShortKeyBoxTxt",False
			DlgVisible "ShortBox",False
			DlgVisible "ShortKeyBox",False
			DlgVisible "KpShortKeyBox",False

			DlgVisible "PreRepStrBoxTxt",False
			DlgVisible "AutoWebFlagBoxTxt",False
			DlgVisible "PreRepStrBox",False
			DlgVisible "AutoWebFlagBox",False

			DlgVisible "LineSplitBoxTxt",False
			DlgVisible "PreInsertSplitBoxTxt",False
			DlgVisible "AppInsertSplitBoxTxt",False
			DlgVisible "ReplaceSplitBoxTxt",False
			DlgVisible "PreInsertSplitBox",False
			DlgVisible "AppInsertSplitBox",False
			DlgVisible "ReplaceSplitBox",False
			DlgVisible "LineSplitModeBox",False

			DlgVisible "AppLngText",False
			DlgVisible "UseLngText",False
			DlgVisible "AppLngList",False
			DlgVisible "UseLngList",False
			DlgVisible "AddLangButton",False
			DlgVisible "AddAllLangButton",False
			DlgVisible "DelLangButton",False
			DlgVisible "DelAllLangButton",False
			DlgVisible "SetAppLangButton",False
			DlgVisible "EditAppLangButton",False
			DlgVisible "DelAppLangButton",False
			DlgVisible "SetUseLangButton",False
			DlgVisible "EditUseLangButton",False
			DlgVisible "DelUseLangButton",False
		Else
			If DlgValue("SetType") <> 0 Then
				DlgVisible "AccKeyBoxTxt",False
				DlgVisible "ExCrBoxTxt",False
				DlgVisible "ChkBktBoxTxt",False
				DlgVisible "KpPairBoxTxt",False
				DlgVisible "AccKeyBox",False
				DlgVisible "ExCrBox",False
				DlgVisible "ChkBktBox",False
				DlgVisible "KpPairBox",False
				DlgVisible "AsiaKeyBox",False
				DlgVisible "AddAcckeyBox",False
			End If
			If DlgValue("SetType") <> 1 Then
				DlgVisible "ChkEndBoxTxt",False
				DlgVisible "NoTrnEndBoxTxt",False
				DlgVisible "AutoTrnEndBoxTxt",False
				DlgVisible "ChkEndBox",False
				DlgVisible "NoTrnEndBox",False
				DlgVisible "AutoTrnEndBox",False
			End If
			If DlgValue("SetType") <> 2 Then
				DlgVisible "ShortBoxTxt",False
				DlgVisible "ShortKeyBoxTxt",False
				DlgVisible "KpShortKeyBoxTxt",False
				DlgVisible "ShortBox",False
				DlgVisible "ShortKeyBox",False
				DlgVisible "KpShortKeyBox",False
			End If
			If DlgValue("SetType") <> 3 Then
				DlgVisible "PreRepStrBoxTxt",False
				DlgVisible "AutoWebFlagBoxTxt",False
				DlgVisible "PreRepStrBox",False
				DlgVisible "AutoWebFlagBox",False
			End If
			If DlgValue("SetType") <> 4 Then
				DlgVisible "LineSplitBoxTxt",False
				DlgVisible "PreInsertSplitBoxTxt",False
				DlgVisible "AppInsertSplitBoxTxt",False
				DlgVisible "ReplaceSplitBoxTxt",False
				DlgVisible "PreInsertSplitBox",False
				DlgVisible "AppInsertSplitBox",False
				DlgVisible "ReplaceSplitBox",False
				DlgVisible "LineSplitModeBox",False
			End If
			If DlgValue("SetType") <> 5 Then
				DlgVisible "AppLngText",False
				DlgVisible "UseLngText",False
				DlgVisible "AppLngList",False
				DlgVisible "UseLngList",False
				DlgVisible "AddLangButton",False
				DlgVisible "AddAllLangButton",False
				DlgVisible "DelLangButton",False
				DlgVisible "DelAllLangButton",False
				DlgVisible "SetAppLangButton",False
				DlgVisible "EditAppLangButton",False
				DlgVisible "DelAppLangButton",False
				DlgVisible "SetUseLangButton",False
				DlgVisible "EditUseLangButton",False
				DlgVisible "DelUseLangButton",False
			End If
		End If
		If DlgValue("Options") <> 2 Then
			DlgVisible "UpdateSetGroup",False
			DlgVisible "UpdateSet",False
			DlgVisible "AutoButton",False
			DlgVisible "ManualButton",False
			DlgVisible "OffButton",False
			DlgVisible "CheckGroup",False
			DlgVisible "UpdateCycleText",False
			DlgVisible "UpdateCycleBox",False
			DlgVisible "UpdateDatesText",False
			DlgVisible "UpdateDateText",False
			DlgVisible "UpdateDateBox",False
			DlgVisible "CheckButton",False
			DlgVisible "WebSiteGroup",False
			DlgVisible "WebSiteBox",False
			DlgVisible "CmdGroup",False
			DlgVisible "CmdPathBoxText",False
			DlgVisible "ArgumentBoxText",False
			DlgVisible "CmdPathBox",False
			DlgVisible "ArgumentBox",False
			DlgVisible "ExeBrowseButton",False
			DlgVisible "ArgumentButton",False
		End If
		If DlgValue("Options") <> 3 Then
			DlgVisible "UILangSetGroup",False
			DlgVisible "UILangSetText1",False
			DlgVisible "UILangSetText2",False
			DlgVisible "UILangList",False
			DlgVisible "UILangSetText3",False

			DlgVisible "UIFontSetGroup",False
			DlgVisible "MainFontText",False
			DlgVisible "MainFontBox",False
			DlgVisible "MainFontButton",False
		End If

		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		SettingsDlgFunc = True '防止按下按钮关闭对话框窗口
		If getMsgList(UIDataList,MsgList,"SettingsDlgFunc",1) = False Then Exit Function
		If DlgItem$ = "NextSetType" Then
			i = DlgValue("SetType")
			If i < 5 Then i = i + 1 Else i = 0
			DlgValue "SetType",i
		End If
		Select Case DlgItem$
		Case "Options","SetType","NextSetType","Engine"
			If DlgValue("Options") < 2 Then
				DlgVisible "GroupBox1",True
				DlgVisible "LevelButton",True
				DlgVisible "AddButton",True
				DlgVisible "ChangButton",True
				DlgVisible "DelButton",True

				DlgVisible "GroupBox2",True
				DlgVisible "ImportButton",True
				DlgVisible "ExportButton",True
			Else
				DlgVisible "GroupBox1",False
				DlgVisible "LevelButton",False
				DlgVisible "AddButton",False
				DlgVisible "ChangButton",False
				DlgVisible "DelButton",False

				DlgVisible "GroupBox2",False
				DlgVisible "ImportButton",False
				DlgVisible "ExportButton",False
			End If
			If DlgValue("Options") < 3 Then
				DlgVisible "ResetButton",True
				DlgVisible "TestButton",True
				DlgVisible "CleanButton",True
				DlgVisible "EditUILangButton",False
			Else
				DlgVisible "ResetButton",False
				DlgVisible "TestButton",False
				DlgVisible "CleanButton",False
				DlgVisible "EditUILangButton",True
			End If
			If DlgValue("Options") = 0 Then
				DlgVisible "EngineList",True
				DlgVisible "tWriteType",True
				DlgVisible "GroupBox4",True
				DlgVisible "Engine",True
				If DlgValue("Engine") = 0 Then
					DlgVisible "ObjectNameText",True
					DlgVisible "AppIDText",True
					DlgVisible "UrlText",True
					DlgVisible "UrlTemplateText",True
					DlgVisible "bstrMethodText",True
					DlgVisible "varAsyncText",True
					DlgVisible "bstrUserText",True
					DlgVisible "bstrPasswordText",True
					DlgVisible "varBodyText",True
					DlgVisible "setRequestHeaderText",True
					DlgVisible "setRequestHeaderText2",True
					DlgVisible "responseTypeText",True
					DlgVisible "TranBeforeStrText",True
					DlgVisible "TranAfterStrText",True
					DlgVisible "ObjectNameBox",True
					DlgVisible "AppIDBox",True
					DlgVisible "UrlBox",True
					DlgVisible "UrlTemplateBox",True
					DlgVisible "bstrMethodBox",True
					DlgVisible "varAsyncBox",True
					DlgVisible "bstrUserBox",True
					DlgVisible "bstrPasswordBox",True
					DlgVisible "varBodyBox",True
					DlgVisible "setRequestHeaderBox",True
					DlgVisible "responseTypeBox",True
					DlgVisible "TranBeforeStrBox",True
					DlgVisible "TranAfterStrBox",True
					DlgVisible "UrlTemplateButton",True
					DlgVisible "bstrMethodButton",True
					DlgVisible "varAsyncButton",True
					DlgVisible "varBodyButton",True
					DlgVisible "RequestButton",True
					DlgVisible "responseTypeButton",True
					DlgVisible "TranBeforeStrButton",True
					DlgVisible "TranAfterStrButton",True
				Else
					DlgVisible "ObjectNameText",False
					DlgVisible "AppIDText",False
					DlgVisible "UrlText",False
					DlgVisible "UrlTemplateText",False
					DlgVisible "bstrMethodText",False
					DlgVisible "varAsyncText",False
					DlgVisible "bstrUserText",False
					DlgVisible "bstrPasswordText",False
					DlgVisible "varBodyText",False
					DlgVisible "setRequestHeaderText",False
					DlgVisible "setRequestHeaderText2",False
					DlgVisible "responseTypeText",False
					DlgVisible "TranBeforeStrText",False
					DlgVisible "TranAfterStrText",False
					DlgVisible "ObjectNameBox",False
					DlgVisible "AppIDBox",False
					DlgVisible "UrlBox",False
					DlgVisible "UrlTemplateBox",False
					DlgVisible "bstrMethodBox",False
					DlgVisible "varAsyncBox",False
					DlgVisible "bstrUserBox",False
					DlgVisible "bstrPasswordBox",False
					DlgVisible "varBodyBox",False
					DlgVisible "setRequestHeaderBox",False
					DlgVisible "responseTypeBox",False
					DlgVisible "TranBeforeStrBox",False
					DlgVisible "TranAfterStrBox",False
					DlgVisible "UrlTemplateButton",False
					DlgVisible "bstrMethodButton",False
					DlgVisible "varAsyncButton",False
					DlgVisible "varBodyButton",False
					DlgVisible "RequestButton",False
					DlgVisible "responseTypeButton",False
					DlgVisible "TranBeforeStrButton",False
					DlgVisible "TranAfterStrButton",False
				End If
				If DlgValue("Engine") = 1 Then
					DlgVisible "LngNameText",True
					DlgVisible "SrcLngText",True
					DlgVisible "TranLngText",True
					DlgVisible "LngNameList",True
					DlgVisible "SrcLngList",True
					DlgVisible "TranLngList",True
					DlgVisible "AddLngButton",True
					DlgVisible "DelLngButton",True
					DlgVisible "DelAllButton",True
					DlgVisible "EditLngButton",True
					DlgVisible "ExtEditButton",True
					DlgVisible "NullLngButton",True
					DlgVisible "ResetLngButton",True
					DlgVisible "ShowNoNullLngButton",True
					DlgVisible "ShowNullLngButton",True
					DlgVisible "ShowAllLngButton",True
				Else
					DlgVisible "LngNameText",False
					DlgVisible "SrcLngText",False
					DlgVisible "TranLngText",False
					DlgVisible "LngNameList",False
					DlgVisible "SrcLngList",False
					DlgVisible "TranLngList",False
					DlgVisible "AddLngButton",False
					DlgVisible "DelLngButton",False
					DlgVisible "DelAllButton",False
					DlgVisible "EditLngButton",False
					DlgVisible "ExtEditButton",False
					DlgVisible "NullLngButton",False
					DlgVisible "ResetLngButton",False
					DlgVisible "ShowNoNullLngButton",False
					DlgVisible "ShowNullLngButton",False
					DlgVisible "ShowAllLngButton",False
				End If
				If DlgValue("Engine") = 2 Then
					DlgVisible "EnableText1",True
					DlgVisible "EnableText2",True
					DlgVisible "EnableBox",True
				Else
					DlgVisible "EnableText1",False
					DlgVisible "EnableText2",False
					DlgVisible "EnableBox",False
				End If
			Else
				DlgVisible "EngineList",False

				DlgVisible "tWriteType",False
				DlgVisible "GroupBox4",False
				DlgVisible "Engine",False

				DlgVisible "ObjectNameText",False
				DlgVisible "AppIDText",False
				DlgVisible "UrlText",False
				DlgVisible "UrlTemplateText",False
				DlgVisible "bstrMethodText",False
				DlgVisible "varAsyncText",False
				DlgVisible "bstrUserText",False
				DlgVisible "bstrPasswordText",False
				DlgVisible "varBodyText",False
				DlgVisible "setRequestHeaderText",False
				DlgVisible "setRequestHeaderText2",False
				DlgVisible "responseTypeText",False
				DlgVisible "TranBeforeStrText",False
				DlgVisible "TranAfterStrText",False
				DlgVisible "ObjectNameBox",False
				DlgVisible "AppIDBox",False
				DlgVisible "UrlBox",False
				DlgVisible "UrlTemplateBox",False
				DlgVisible "bstrMethodBox",False
				DlgVisible "varAsyncBox",False
				DlgVisible "bstrUserBox",False
				DlgVisible "bstrPasswordBox",False
				DlgVisible "varBodyBox",False
				DlgVisible "setRequestHeaderBox",False
				DlgVisible "responseTypeBox",False
				DlgVisible "TranBeforeStrBox",False
				DlgVisible "TranAfterStrBox",False
				DlgVisible "UrlTemplateButton",False
				DlgVisible "bstrMethodButton",False
				DlgVisible "varAsyncButton",False
				DlgVisible "varBodyButton",False
				DlgVisible "RequestButton",False
				DlgVisible "responseTypeButton",False
				DlgVisible "TranBeforeStrButton",False
				DlgVisible "TranAfterStrButton",False

				DlgVisible "LngNameText",False
				DlgVisible "SrcLngText",False
				DlgVisible "TranLngText",False
				DlgVisible "LngNameList",False
				DlgVisible "SrcLngList",False
				DlgVisible "TranLngList",False
				DlgVisible "AddLngButton",False
				DlgVisible "DelLngButton",False
				DlgVisible "DelAllButton",False
				DlgVisible "EditLngButton",False
				DlgVisible "ExtEditButton",False
				DlgVisible "NullLngButton",False
				DlgVisible "ResetLngButton",False
				DlgVisible "ShowNoNullLngButton",False
				DlgVisible "ShowNullLngButton",False
				DlgVisible "ShowAllLngButton",False

				DlgVisible "EnableText1",False
				DlgVisible "EnableText2",False
				DlgVisible "EnableBox",False
			End If
			If DlgValue("Options") = 1 Then
				DlgVisible "CheckList",True
				DlgVisible "cWriteType",True

				DlgVisible "GroupBox3",True
				DlgVisible "SetType",True
				DlgVisible "NextSetType",True
				DlgVisible "SetItemText",True
				If DlgValue("SetType") = 0 Then
					DlgVisible "AccKeyBoxTxt",True
					DlgVisible "ExCrBoxTxt",True
					DlgVisible "ChkBktBoxTxt",True
					DlgVisible "KpPairBoxTxt",True
					DlgVisible "AccKeyBox",True
					DlgVisible "ExCrBox",True
					DlgVisible "ChkBktBox",True
					DlgVisible "KpPairBox",True
					DlgVisible "AsiaKeyBox",True
					DlgVisible "AddAcckeyBox",True
				Else
					DlgVisible "AccKeyBoxTxt",False
					DlgVisible "ExCrBoxTxt",False
					DlgVisible "ChkBktBoxTxt",False
					DlgVisible "KpPairBoxTxt",False
					DlgVisible "AccKeyBox",False
					DlgVisible "ExCrBox",False
					DlgVisible "ChkBktBox",False
					DlgVisible "KpPairBox",False
					DlgVisible "AsiaKeyBox",False
					DlgVisible "AddAcckeyBox",False
				End If
				If DlgValue("SetType") = 1 Then
					DlgVisible "ChkEndBoxTxt",True
					DlgVisible "NoTrnEndBoxTxt",True
					DlgVisible "AutoTrnEndBoxTxt",True
					DlgVisible "ChkEndBox",True
					DlgVisible "NoTrnEndBox",True
					DlgVisible "AutoTrnEndBox",True
				Else
					DlgVisible "ChkEndBoxTxt",False
					DlgVisible "NoTrnEndBoxTxt",False
					DlgVisible "AutoTrnEndBoxTxt",False
					DlgVisible "ChkEndBox",False
					DlgVisible "NoTrnEndBox",False
					DlgVisible "AutoTrnEndBox",False
				End If
				If DlgValue("SetType") = 2 Then
					DlgVisible "ShortBoxTxt",True
					DlgVisible "ShortKeyBoxTxt",True
					DlgVisible "KpShortKeyBoxTxt",True
					DlgVisible "ShortBox",True
					DlgVisible "ShortKeyBox",True
					DlgVisible "KpShortKeyBox",True
				Else
					DlgVisible "ShortBoxTxt",False
					DlgVisible "ShortKeyBoxTxt",False
					DlgVisible "KpShortKeyBoxTxt",False
					DlgVisible "ShortBox",False
					DlgVisible "ShortKeyBox",False
					DlgVisible "KpShortKeyBox",False
				End If
				If DlgValue("SetType") = 3 Then
					DlgVisible "PreRepStrBoxTxt",True
					DlgVisible "AutoWebFlagBoxTxt",True
					DlgVisible "PreRepStrBox",True
					DlgVisible "AutoWebFlagBox",True
				Else
					DlgVisible "PreRepStrBoxTxt",False
					DlgVisible "AutoWebFlagBoxTxt",False
					DlgVisible "PreRepStrBox",False
					DlgVisible "AutoWebFlagBox",False
				End	If
				If DlgValue("SetType") = 4 Then
					DlgVisible "LineSplitBoxTxt",True
					DlgVisible "PreInsertSplitBoxTxt",True
					DlgVisible "AppInsertSplitBoxTxt",True
					DlgVisible "ReplaceSplitBoxTxt",True
					DlgVisible "PreInsertSplitBox",True
					DlgVisible "AppInsertSplitBox",True
					DlgVisible "ReplaceSplitBox",True
					DlgVisible "LineSplitModeBox",True
				Else
					DlgVisible "LineSplitBoxTxt",False
					DlgVisible "PreInsertSplitBoxTxt",False
					DlgVisible "AppInsertSplitBoxTxt",False
					DlgVisible "ReplaceSplitBoxTxt",False
					DlgVisible "PreInsertSplitBox",False
					DlgVisible "AppInsertSplitBox",False
					DlgVisible "ReplaceSplitBox",False
					DlgVisible "LineSplitModeBox",False
				End If
				If DlgValue("SetType") = 5 Then
					DlgVisible "AppLngText",True
					DlgVisible "UseLngText",True
					DlgVisible "AppLngList",True
					DlgVisible "UseLngList",True
					DlgVisible "AddLangButton",True
					DlgVisible "AddAllLangButton",True
					DlgVisible "DelLangButton",True
					DlgVisible "DelAllLangButton",True
					DlgVisible "SetAppLangButton",True
					DlgVisible "EditAppLangButton",True
					DlgVisible "DelAppLangButton",True
					DlgVisible "SetUseLangButton",True
					DlgVisible "EditUseLangButton",True
					DlgVisible "DelUseLangButton",True
				Else
					DlgVisible "AppLngText",False
					DlgVisible "UseLngText",False
					DlgVisible "AppLngList",False
					DlgVisible "UseLngList",False
					DlgVisible "AddLangButton",False
					DlgVisible "AddAllLangButton",False
					DlgVisible "DelLangButton",False
					DlgVisible "DelAllLangButton",False
					DlgVisible "SetAppLangButton",False
					DlgVisible "EditAppLangButton",False
					DlgVisible "DelAppLangButton",False
					DlgVisible "SetUseLangButton",False
					DlgVisible "EditUseLangButton",False
					DlgVisible "DelUseLangButton",False
				End If
			Else
				DlgVisible "CheckList",False
				DlgVisible "cWriteType",False

				DlgVisible "GroupBox3",False
				DlgVisible "SetType",False
				DlgVisible "NextSetType",False
				DlgVisible "SetItemText",False

				DlgVisible "AccKeyBoxTxt",False
				DlgVisible "ExCrBoxTxt",False
				DlgVisible "ChkBktBoxTxt",False
				DlgVisible "KpPairBoxTxt",False
				DlgVisible "AccKeyBox",False
				DlgVisible "ExCrBox",False
				DlgVisible "ChkBktBox",False
				DlgVisible "KpPairBox",False
				DlgVisible "AsiaKeyBox",False
				DlgVisible "AddAcckeyBox",False

				DlgVisible "ChkEndBoxTxt",False
				DlgVisible "NoTrnEndBoxTxt",False
				DlgVisible "AutoTrnEndBoxTxt",False
				DlgVisible "ChkEndBox",False
				DlgVisible "NoTrnEndBox",False
				DlgVisible "AutoTrnEndBox",False

				DlgVisible "ShortBoxTxt",False
				DlgVisible "ShortKeyBoxTxt",False
				DlgVisible "KpShortKeyBoxTxt",False
				DlgVisible "ShortBox",False
				DlgVisible "ShortKeyBox",False
				DlgVisible "KpShortKeyBox",False

				DlgVisible "PreRepStrBoxTxt",False
				DlgVisible "AutoWebFlagBoxTxt",False
				DlgVisible "PreRepStrBox",False
				DlgVisible "AutoWebFlagBox",False

				DlgVisible "LineSplitBoxTxt",False
				DlgVisible "PreInsertSplitBoxTxt",False
				DlgVisible "AppInsertSplitBoxTxt",False
				DlgVisible "ReplaceSplitBoxTxt",False
				DlgVisible "PreInsertSplitBox",False
				DlgVisible "AppInsertSplitBox",False
				DlgVisible "ReplaceSplitBox",False
				DlgVisible "LineSplitModeBox",False

				DlgVisible "AppLngText",False
				DlgVisible "UseLngText",False
				DlgVisible "AppLngList",False
				DlgVisible "UseLngList",False
				DlgVisible "AddLangButton",False
				DlgVisible "AddAllLangButton",False
				DlgVisible "DelLangButton",False
				DlgVisible "DelAllLangButton",False
				DlgVisible "SetAppLangButton",False
				DlgVisible "EditAppLangButton",False
				DlgVisible "DelAppLangButton",False
				DlgVisible "SetUseLangButton",False
				DlgVisible "EditUseLangButton",False
				DlgVisible "DelUseLangButton",False
			End If
			If DlgValue("Options") = 2 Then
				DlgVisible "UpdateSetGroup",True
				DlgVisible "UpdateSet",True
				DlgVisible "AutoButton",True
				DlgVisible "ManualButton",True
				DlgVisible "OffButton",True
				DlgVisible "CheckGroup",True
				DlgVisible "UpdateCycleText",True
				DlgVisible "UpdateCycleBox",True
				DlgVisible "UpdateDatesText",True
				DlgVisible "UpdateDateText",True
				DlgVisible "UpdateDateBox",True
				DlgVisible "CheckButton",True
				DlgVisible "WebSiteGroup",True
				DlgVisible "WebSiteBox",True
				DlgVisible "CmdGroup",True
				DlgVisible "CmdPathBoxText",True
				DlgVisible "ArgumentBoxText",True
				DlgVisible "CmdPathBox",True
				DlgVisible "ArgumentBox",True
				DlgVisible "ExeBrowseButton",True
				DlgVisible "ArgumentButton",True
			Else
				DlgVisible "UpdateSetGroup",False
				DlgVisible "UpdateSet",False
				DlgVisible "AutoButton",False
				DlgVisible "ManualButton",False
				DlgVisible "OffButton",False
				DlgVisible "CheckGroup",False
				DlgVisible "UpdateCycleText",False
				DlgVisible "UpdateCycleBox",False
				DlgVisible "UpdateDatesText",False
				DlgVisible "UpdateDateText",False
				DlgVisible "UpdateDateBox",False
				DlgVisible "CheckButton",False
				DlgVisible "WebSiteGroup",False
				DlgVisible "WebSiteBox",False
				DlgVisible "CmdGroup",False
				DlgVisible "CmdPathBoxText",False
				DlgVisible "ArgumentBoxText",False
				DlgVisible "CmdPathBox",False
				DlgVisible "ArgumentBox",False
				DlgVisible "ExeBrowseButton",False
				DlgVisible "ArgumentButton",False
			End If
			If DlgValue("Options") = 3 Then
				DlgVisible "UILangSetGroup",True
				DlgVisible "UILangSetText1",True
				DlgVisible "UILangSetText2",True
				DlgVisible "UILangList",True
				DlgVisible "UILangSetText3",True

				DlgVisible "UIFontSetGroup",True
				DlgVisible "MainFontText",True
				DlgVisible "MainFontBox",True
				DlgVisible "MainFontButton",True
			Else
				DlgVisible "UILangSetGroup",False
				DlgVisible "UILangSetText1",False
				DlgVisible "UILangSetText2",False
				DlgVisible "UILangList",False
				DlgVisible "UILangSetText3",False

				DlgVisible "UIFontSetGroup",False
				DlgVisible "MainFontText",False
				DlgVisible "MainFontBox",False
				DlgVisible "MainFontButton",False
			End If
		Case "HelpButton"
			Select Case DlgValue("Options")
			Case 0
				Call Help("EngineSetHelp")
			Case 1
				Call Help("CheckSetHelp")
			Case 2
				Call Help("UpdateSetHelp")
			Case 3
				Call Help("UILangSetHelp")
			End Select
			Exit Function
		Case "OKButton"
			If DlgText("CmdPathBox") = "" Or DlgText("ArgumentBox") = "" Then
	    		MsgBox(MsgList(35) & Path,vbOkOnly+vbInformation,MsgList(0))
				Exit Function
	    	End If
			tStemp = CheckNullData("",EngineDataList,"1,6-9,15-19",1)
			If tStemp = False Then tStemp = CheckTargetValue(EngineDataList,-1)
			cStemp = CheckNullData("",CheckDataList,"1,4,14-17",1)
			If tStemp = True Or cStemp = True Then
				If tStemp = True And cStemp <> True Then
					Temp = Replace$(MsgList(10),"%s",MsgList(14)) & MsgList(6)
				ElseIf tStemp <> True And cStemp = True Then
					Temp = Replace$(MsgList(10),"%s",MsgList(15)) & MsgList(6)
				ElseIf tStemp = True And cStemp = True Then
					Temp = Replace$(MsgList(10),"%s",MsgList(16)) & MsgList(6)
				End If
				If MsgBox(Temp,vbYesNo+vbInformation,MsgList(5)) = vbNo Then
					Exit Function
				End If
			End If
			If DlgValue("tWriteType") = 0 Then Path = EngineFilePath Else Path = EngineRegKey
			If DlgValue("cWriteType") = 0 Then Temp = CheckFilePath Else Temp = CheckRegKey
			tStemp = WriteEngineSet(Path,"Sets")
			cStemp = WriteCheckSet(Temp,"Sets")
			If tStemp = False Or cStemp = False Then
				If tStemp = False And cStemp <> False Then
					Temp = MsgList(14) & Replace$(MsgList(20),"%s",Path)
				ElseIf tStemp <> False And cStemp = False Then
					Temp = MsgList(15) & Replace$(MsgList(20),"%s",Temp)
				ElseIf tStemp = False And cStemp = False Then
					Temp = MsgList(16) & Replace$(MsgList(20),"%s",Path & vbCrLf & Temp)
				End If
				MsgBox(Temp,vbOkOnly+vbInformation,MsgList(0))
			Else
				EngineListBak = EngineList
				EngineDataListBak = EngineDataList
				CheckListBak = CheckList
				CheckDataListBak = CheckDataList
				tUpdateSetBak = tUpdateSet
				LFListBak = LFList
				SettingsDlgFunc = False
			End If
			Exit Function
		Case "CancelButton"
			EngineList = EngineListBak
			EngineDataList = EngineDataListBak
			CheckList = CheckListBak
			CheckDataList = CheckDataListBak
			tUpdateSet = tUpdateSetBak
			SettingsDlgFunc = False
			Exit Function
		End Select

		Select Case DlgValue("Options")
		Case 0
			HeaderID = DlgValue("EngineList")
			MainArray = ReSplit(EngineDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),SubJoinStr)
			SplitData(MainArray(2),LngNameList,SrcLngList,TranLngList)
			Select Case DlgItem$
			Case "LevelButton"
				Header = EngineList(HeaderID)
				If SetLevel(EngineList,HeaderID,MsgList(52)) = True Then
					DlgListBoxArray "EngineList",EngineList()
					DlgText "EngineList",Header
					Set Dic = CreateObject("Scripting.Dictionary")
					For i = 0 To UBound(EngineDataList)
						TempList = ReSplit(EngineDataList(i),JoinStr)
						If Not Dic.Exists(TempList(0)) Then
							Dic.Add(TempList(0),i)
						End If
					Next i
					TempArray = EngineDataList
					For i = 0 To UBound(EngineList)
						If Dic.Exists(EngineList(i)) Then
							j = Dic.Item(EngineList(i))
							TempArray(n) = EngineDataList(j)
							n = n + 1
						End If
					Next i
					EngineDataList = TempArray
					Set Dic = Nothing
				End If
				Exit Function
			Case "AddButton"
				NewData = AddSet(EngineList)
				If NewData = "" Then Exit Function
				ReDim SetsArray(UBound(SetsArray)) As String
				SetsArray(0) = DefaultObject
				Data = Join(SetsArray,SubJoinStr)
				LangPairList = LangCodeList("engine",0,-1)
				Temp = NewData & JoinStr & Data & JoinStr & Join(LangPairList,SubLngJoinStr)
				CreateArray(NewData,Temp,EngineList,EngineDataList)
				DlgListBoxArray "EngineList",EngineList()
				DlgText "EngineList",NewData
				HeaderID = DlgValue("EngineList")
				MainArray = ReSplit(EngineDataList(HeaderID),JoinStr)
				SetsArray = ReSplit(MainArray(1),SubJoinStr)
				SplitData(MainArray(2),LngNameList,SrcLngList,TranLngList)
				DlgItem$ = "EngineList"
			Case "ChangButton"
				NewData = EditSet(EngineList,HeaderID)
				If NewData <> "" Then
					EngineList(HeaderID) = NewData
					MainArray(0) = NewData
					EngineDataList(HeaderID) = Join(MainArray,JoinStr)
					DlgListBoxArray "EngineList",EngineList()
					DlgValue "EngineList",HeaderID
				End If
				Exit Function
	    	Case "DelButton"
				Header = EngineList(HeaderID)
				If MsgBox(Replace(MsgList(12),"%s",Header),vbYesNo+vbInformation,MsgList(11)) = vbNo Then
					Exit Function
				End If
				i = UBound(EngineList)
				Call DelArrays(EngineList,EngineDataList,HeaderID)
				If HeaderID > 0 And HeaderID = i Then HeaderID = HeaderID - 1
				DlgListBoxArray "EngineList",EngineList()
				DlgValue "EngineList",HeaderID
				MainArray = ReSplit(EngineDataList(HeaderID),JoinStr)
				SetsArray = ReSplit(MainArray(1),SubJoinStr)
				SplitData(MainArray(2),LngNameList,SrcLngList,TranLngList)
				DlgItem$ = "EngineList"
			Case "ResetButton"
				Header = EngineList(HeaderID)
				ReDim TempArray(1)
				For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
					If DefaultEngineList(i) = Header Then
						TempArray(0) = MsgList(1)
						Exit For
					End If
				Next i
				tStemp = CheckNullData(Header,EngineDataListBak,"1,6-9,15-19",0)
				If tStemp = False Then TempArray(1) = MsgList(2)
				For i = LBound(EngineList) To UBound(EngineList)
					If i <> HeaderID Then
						ReDim Preserve TempArray(i+2)
						TempArray(i+2) = MsgList(3) & " - " & EngineList(i)
					End If
				Next i
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				If i = 0 Then
					SetsArray = ReSplit(EngineSettings(Header),SubJoinStr)
					TempArray = LangCodeList(Header,0,-1)
					TempArray(2) = Join(TempArray,SubLngJoinStr)
				ElseIf i = 1 Then
					For n = LBound(EngineDataListBak) To UBound(EngineDataListBak)
						TempArray = ReSplit(EngineDataListBak(n),JoinStr)
						If TempArray(0) = Header Then
							SetsArray = ReSplit(TempArray(1),SubJoinStr)
							Exit For
						End If
					Next n
				ElseIf i >= 2 Then
					Temp = Mid(TempArray(i),InStr(TempArray(i),MsgList(3) & " - ") + Len(MsgList(3) & " - "))
					For n = LBound(EngineList) To UBound(EngineList)
						If EngineList(n) = Temp Then
							TempArray = ReSplit(EngineDataList(n),JoinStr)
							SetsArray = ReSplit(TempArray(1),SubJoinStr)
							Exit For
						End If
					Next n
				End If
				Select Case DlgValue("Engine")
				Case 0
					DlgText "ObjectNameBox",SetsArray(0)
					DlgText "AppIDBox",SetsArray(1)
					DlgText "UrlBox",SetsArray(2)
					DlgText "UrlTemplateBox",SetsArray(3)
					DlgText "bstrMethodBox",SetsArray(4)
					DlgText "varAsyncBox",SetsArray(5)
					DlgText "bstrUserBox",SetsArray(6)
					DlgText "bstrPasswordBox",SetsArray(7)
					DlgText "varBodyBox",SetsArray(8)
					DlgText "setRequestHeaderBox",SetsArray(9)
					DlgText "responseTypeBox",SetsArray(10)
					Select Case DlgText("responseTypeBox")
					Case "responseText"
						DlgText "TranBeforeStrBox",SetsArray(11)
						DlgText "TranAfterStrBox",SetsArray(12)
					Case "responseBody"
						DlgText "TranBeforeStrBox",SetsArray(13)
						DlgText "TranAfterStrBox",SetsArray(14)
					Case "responseStream"
						DlgText "TranBeforeStrBox",SetsArray(15)
						DlgText "TranAfterStrBox",SetsArray(16)
					Case "responseXML"
						DlgText "TranBeforeStrBox",SetsArray(17)
						DlgText "TranAfterStrBox",SetsArray(18)
					Case Else
						DlgText "TranBeforeStrBox",""
						DlgText "TranAfterStrBox",""
					End Select
					If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject
				Case 1
					If TempArray(2) = "" Then Exit Function
					n = 0
					LngName = DlgText("LngNameList")
					SplitData(TempArray(2),LngNameList,SrcLngList,TranLngList)
					DlgListBoxArray "LngNameList",LngNameList()
					DlgListBoxArray "SrcLngList",SrcLngList()
					DlgListBoxArray "TranLngList",TranLngList()
					For i = 0 To UBound(LngNameList)
						If LngNameList(i) = LngName Then
							n = i
							Exit For
						End If
					Next i
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
					DlgEnable "ShowNoNullLngButton",True
					DlgEnable "ShowNullLngButton",True
					DlgEnable "ShowAllLngButton",False
				Case 2
					DlgValue "EnableBox",StrToLong(SetsArray(19))
				End Select
			Case "CleanButton"
				Select Case DlgValue("Engine")
				Case 0
					'DlgText "ObjectNameBox",""
					DlgText "AppIDBox",""
					DlgText "UrlBox",""
					DlgText "UrlTemplateBox",""
					DlgText "bstrMethodBox",""
					DlgText "varAsyncBox",""
					DlgText "bstrUserBox",""
					DlgText "bstrPasswordBox",""
					DlgText "varBodyBox",""
					DlgText "setRequestHeaderBox",""
					DlgText "responseTypeBox",""
					DlgText "TranBeforeStrBox",""
					DlgText "TranAfterStrBox",""
				Case 1
					n = DlgValue("TranLngList")
					For i = 0 To UBound(LngNameList)
						TranLngList(i) = NullValue
					Next i
					If DlgEnable("ShowAllLngButton") = False Then
						DlgListBoxArray "TranLngList",TranLngList()
					Else
						For i = 0 To UBound(DelLngNameList)
							If DelLngNameList(i) <> "" Then
								DelTranLngList(i) = NullValue
							End If
						Next i
						DlgListBoxArray "TranLngList",DelTranLngList()
					End If
					DlgValue "TranLngList",n
				Case 2
					DlgValue "EnableBox",0
				End Select
			Case "LngNameList"
				DlgValue "SrcLngList",DlgValue("LngNameList")
				DlgValue "TranLngList",DlgValue("LngNameList")
			Case "SrcLngList"
				DlgValue "LngNameList",DlgValue("SrcLngList")
				DlgValue "TranLngList",DlgValue("SrcLngList")
			Case "TranLngList"
				DlgValue "LngNameList",DlgValue("TranLngList")
				DlgValue "SrcLngList",DlgValue("TranLngList")
			Case "AddLngButton","DelLngButton","DelAllButton"
				Select Case DlgItem$
				Case "AddLngButton"
					TempList = ReSplit(MainArray(2),SubLngJoinStr)
					NewData = EditLang(TempList,"","","")
					If NewData <> "" Then
						n = UBound(LngNameList)
						If n > 0 Or LngNameList(n) <> "" Then n = n + 1
						ReDim Preserve LngNameList(n),SrcLngList(n),TranLngList(n)
						LangPairList = ReSplit(NewData,LngJoinStr)
						LngNameList(n) = LangPairList(0)
						SrcLngList(n) = LangPairList(1)
						TranLngList(n) = LangPairList(2)
					End If
				Case "DelLngButton"
					LngName = DlgText("LngNameList")
					If LngName = "" Then Exit Function
					If MsgBox(Replace(MsgList(13),"%s",LngName),vbYesNo+vbInformation,MsgList(11)) = vbNo Then Exit Function
					n = DlgValue("LngNameList")
					NewData = DlgText("LngNameList")
					TempList = ReSplit(MainArray(2),SubLngJoinStr)
					If n > 0 And n = UBound(TempList) Then n = n - 1
					Call DelArray(TempList,NewData,LngJoinStr)
					SplitData(Join(TempList,SubLngJoinStr),LngNameList,SrcLngList,TranLngList)
				Case Else
					If MsgBox(MsgList(26),vbYesNo+vbInformation,MsgList(11)) = vbNo Then Exit Function
					n = 0
					NewData = DlgText("LngNameList")
					ReDim LngNameList(0),SrcLngList(0),TranLngList(0)
				End Select
				If NewData = "" Then Exit Function
				If DlgEnable("ShowAllLngButton") = False Then
					DlgListBoxArray "LngNameList",LngNameList()
					DlgListBoxArray "SrcLngList",SrcLngList()
					DlgListBoxArray "TranLngList",TranLngList()
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
				Else
					j = UBound(LngNameList)
					ReDim DelLngNameList(j),DelSrcLngList(j),DelTranLngList(j)
					n = 0
					If DlgEnable("ShowNoNullLngButton") = False Then
						For i = 0 To UBound(LngNameList)
							If TranLngList(i) <> NullValue Then
								DelLngNameList(n) = LngNameList(i)
								DelSrcLngList(n) = SrcLngList(i)
								DelTranLngList(n) = TranLngList(i)
								n = n + 1
							End If
						Next i
					Else
						For i = 0 To UBound(LngNameList)
							If TranLngList(i) = NullValue Then
								DelLngNameList(n) = LngNameList(i)
								DelSrcLngList(n) = SrcLngList(i)
								DelTranLngList(n) = TranLngList(i)
								n = n + 1
							End If
						Next i
					End If
					If n > 0 Then n = n - 1
					ReDim Preserve DelLngNameList(n),DelSrcLngList(n),DelTranLngList(n)
					If DlgItem$ = "DelLngButton" Then
						n = j
						If n > 0 And n = UBound(DelLngNameList) + 1 Then n = n - 1
					End If
					DlgListBoxArray "LngNameList",DelLngNameList()
					DlgListBoxArray "SrcLngList",DelSrcLngList()
					DlgListBoxArray "TranLngList",DelTranLngList()
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
				End If
			Case "EditLngButton","NullLngButton","ResetLngButton"
				Header = EngineList(HeaderID)
				LngName = DlgText("LngNameList")
				SrcLngCode = DlgText("SrcLngList")
				TranLngCode = DlgText("TranLngList")
				Select Case DlgItem$
				Case "EditLngButton"
					TempList = ReSplit(MainArray(2),SubLngJoinStr)
					NewData = EditLang(TempList,LngName,SrcLngCode,TranLngCode)
					If NewData <> "" Then
						LangPairList = ReSplit(NewData,LngJoinStr)
						NewLngName = LangPairList(0)
						NewSrcLngCode = LangPairList(1)
						NewTranLngCode = LangPairList(2)
					End If
				Case "NullLngButton"
					TranLngCode = DlgText("TranLngList")
					If TranLngCode <> NullValue Then
						NewData = NullValue
						NewLngName = LngName
						NewSrcLngCode = DlgText("SrcLngList")
						NewTranLngCode = NullValue
					End If
				Case Else
					For i = LBound(EngineDataListBak) To UBound(EngineDataListBak)
						TempArray = ReSplit(EngineDataListBak(i),JoinStr)
						If TempArray(0) = Header Then
							TempList = ReSplit(TempArray(2),SubLngJoinStr)
							For j = 0 To UBound(TempList)
								LangPairList = ReSplit(TempList(j),LngJoinStr)
								If LangPairList(0) = LngName Then
									NewData = NullValue
									NewLngName = LangPairList(0)
									NewSrcLngCode = LangPairList(1)
									NewTranLngCode = LangPairList(2)
									Exit For
								End If
							Next j
							Exit For
						End If
					Next i
				End Select
				If NewData = "" Then Exit Function
				If NewSrcLngCode = "" Then NewSrcLngCode = NullValue
				If NewTranLngCode = "" Then NewTranLngCode = NullValue
				n = DlgValue("LngNameList")
				If DlgEnable("ShowAllLngButton") = False Then
					LngNameList(n) = NewLngName
					SrcLngList(n) = NewSrcLngCode
					TranLngList(n) = NewTranLngCode
					If n < UBound(LngNameList) Then n = n + 1
					DlgListBoxArray "LngNameList",LngNameList()
					DlgListBoxArray "SrcLngList",SrcLngList()
					DlgListBoxArray "TranLngList",TranLngList()
				Else
					For i = 0 To UBound(LngNameList)
						If LngNameList(i) = LngName Then
							LngNameList(i) = NewLngName
							SrcLngList(i) = NewSrcLngCode
							TranLngList(i) = NewTranLngCode
							Exit For
						End If
					Next i
					DelLngNameList(n) = NewLngName
					DelSrcLngList(n) = NewSrcLngCode
					DelTranLngList(n) = NewTranLngCode
					If n < UBound(DelLngNameList) Then n = n + 1
					DlgListBoxArray "LngNameList",DelLngNameList()
					DlgListBoxArray "SrcLngList",DelSrcLngList()
					DlgListBoxArray "TranLngList",DelTranLngList()
				End If
				DlgValue "LngNameList",n
				DlgValue "SrcLngList",n
				DlgValue "TranLngList",n
			Case "ShowNoNullLngButton","ShowNullLngButton"
				LngName = DlgText("LngNameList")
				j = UBound(LngNameList)
				ReDim DelLngNameList(j),DelSrcLngList(j),DelTranLngList(j)
				n = 0
				If DlgItem$ = "ShowNoNullLngButton" Then
					For i = 0 To UBound(LngNameList)
						If TranLngList(i) <> NullValue Then
							DelLngNameList(n) = LngNameList(i)
							DelSrcLngList(n) = SrcLngList(i)
							DelTranLngList(n) = TranLngList(i)
							If LngNameList(i) = LngName Then j = n
							n = n + 1
						End If
					Next i
				Else
					For i = 0 To UBound(LngNameList)
						If TranLngList(i) = NullValue Then
							DelLngNameList(n) = LngNameList(i)
							DelSrcLngList(n) = SrcLngList(i)
							DelTranLngList(n) = TranLngList(i)
							If LngNameList(i) = LngName Then j = n
							n = n + 1
						End If
					Next i
				End If
				If n > 0 Then n = n - 1
				ReDim Preserve DelLngNameList(n),DelSrcLngList(n),DelTranLngList(n)
				If DlgItem$ = "ShowNoNullLngButton" Then
					DlgEnable "ShowNoNullLngButton",False
					DlgEnable "ShowNullLngButton",True
					DlgEnable "ShowAllLngButton",True
				Else
					DlgEnable "ShowNoNullLngButton",True
					DlgEnable "ShowNullLngButton",False
					DlgEnable "ShowAllLngButton",True
				End If
				DlgListBoxArray "LngNameList",DelLngNameList()
				DlgListBoxArray "SrcLngList",DelSrcLngList()
				DlgListBoxArray "TranLngList",DelTranLngList()
				DlgValue "LngNameList",n
				DlgValue "SrcLngList",n
				DlgValue "TranLngList",n
			Case "ShowAllLngButton"
				LngName = DlgText("LngNameList")
				n = DlgValue("LngNameList")
				If DlgEnable("ShowAllLngButton") = True Then
					For i = 0 To UBound(LngNameList)
						If LngNameList(i) = LngName Then
							n = i
							Exit For
						End If
					Next i
				End If
				DlgListBoxArray "LngNameList",LngNameList()
				DlgListBoxArray "SrcLngList",SrcLngList()
				DlgListBoxArray "TranLngList",TranLngList()
				DlgValue "LngNameList",n
				DlgValue "SrcLngList",n
				DlgValue "TranLngList",n
				DlgEnable "ShowNoNullLngButton",True
				DlgEnable "ShowNullLngButton",True
				DlgEnable "ShowAllLngButton",False
			Case "ExtEditButton"
				ReDim TempList(UBound(Tools))
				For i = 0 To UBound(Tools)
					TempList(i) = Tools(i).sName
				Next i
				n = ShowPopupMenu(TempList,vbPopupUseRightButton)
				If n < 0 Then Exit Function
				TempList = LngNameList
				For i = 0 To UBound(LngNameList)
					TempList(i) = LngNameList(i) & vbTab & SrcLngList(i) & vbTab & TranLngList(i)
				Next i
				Temp = Join(TempList,vbCrLf)
				FilePath = trn.Project.Location & "\~temp.xls"
				If WriteToFile(FilePath,Temp,"ANSI") = True Then
					If Dir(FilePath) <> "" Then
						ReDim FileDataList(0)
						FileDataList(0) = FilePath & JoinStr & "ANSI"
						If OpenFile(FilePath,FileDataList,n,True) = True Then
							textStr = ReadFile(FilePath,"ANSI")
						End If
						On Error Resume Next
						Kill FilePath
						On Error GoTo 0
					End If
				End If
				If textStr = "" Or Temp = textStr Then Exit Function
				ReDim LngNameList(0),SrcLngList(0),TranLngList(0)
				FileLines = ReSplit(textStr,vbCrLf)
				j = UBound(FileLines)
				ReDim LngNameList(j),SrcLngList(j),TranLngList(j)
				n = 0
				For i = 0 To j
					TempArray = ReSplit(FileLines(i),vbTab)
					If UBound(TempArray) > 1 Then
						If TempArray(1) = "" Then TempArray(1) = NullValue
						If TempArray(2) = "" Then TempArray(2) = NullValue
						LngNameList(n) = TempArray(0)
						SrcLngList(n) = TempArray(1)
						TranLngList(n) = TempArray(2)
						n = n + 1
					End If
				Next i
				If n > 0 Then n = n - 1
				ReDim Preserve LngNameList(n),SrcLngList(n),TranLngList(n)
				DlgListBoxArray "LngNameList",LngNameList()
				DlgListBoxArray "SrcLngList",SrcLngList()
				DlgListBoxArray "TranLngList",TranLngList()
				DlgValue "LngNameList",0
				DlgValue "SrcLngList",0
				DlgValue "TranLngList",0
				DlgEnable "ShowNoNullLngButton",True
				DlgEnable "ShowNullLngButton",True
				DlgEnable "ShowAllLngButton",False
			Case "UrlTemplateButton","varBodyButton","RequestButton"
				ReDim TempArray(4)
				TempArray(0) = MsgList(39)
				TempArray(1) = MsgList(40)
				TempArray(2) = MsgList(41)
				TempArray(3) = MsgList(42)
				TempArray(4) = MsgList(43)
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				TempArray(i) = Mid(TempArray(i),InStr(TempArray(i),"{"))
				Select Case DlgItem$
				Case "UrlTemplateButton"
					DlgText "UrlTemplateBox",DlgText("UrlTemplateBox") & TempArray(i)
				Case "varBodyButton"
					DlgText "varBodyBox",DlgText("varBodyBox") & TempArray(i)
				Case "RequestButton"
					DlgText "setRequestHeaderBox",DlgText("setRequestHeaderBox") & TempArray(i)
				End Select
			Case "bstrMethodButton"
				ReDim TempArray(1)
				TempArray(0) = "GET"
				TempArray(1) = "POST"
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				DlgText "bstrMethodBox",TempArray(i)
			Case "varAsyncButton"
				ReDim TempArray(1)
				TempArray(0) = MsgList(44)
				TempArray(1) = MsgList(45)
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				TempArray(i) = Mid(TempArray(i),InStr(TempArray(i),"{") + 1,InStr(TempArray(i),"}") - InStr(TempArray(i),"{") - 1)
				DlgText "varAsyncBox",TempArray(i)
			Case "responseTypeButton"
				ReDim TempArray(3)
				TempArray(0) = MsgList(46)
				TempArray(1) = MsgList(47)
				TempArray(2) = MsgList(48)
				TempArray(3) = MsgList(49)
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				TempArray(i) = Mid(TempArray(i),InStr(TempArray(i),"{") + 1,InStr(TempArray(i),"}") - InStr(TempArray(i),"{") - 1)
				DlgText "responseTypeBox",TempArray(i)
				Select Case TempArray(i)
				Case "responseText"
					DlgText "TranBeforeStrBox",SetsArray(11)
					DlgText "TranAfterStrBox",SetsArray(12)
				Case "responseBody"
					DlgText "TranBeforeStrBox",SetsArray(13)
					DlgText "TranAfterStrBox",SetsArray(14)
				Case "responseStream"
					DlgText "TranBeforeStrBox",SetsArray(15)
					DlgText "TranAfterStrBox",SetsArray(16)
				Case "responseXML"
					DlgText "TranBeforeStrBox",SetsArray(17)
					DlgText "TranAfterStrBox",SetsArray(18)
				Case Else
					DlgText "TranBeforeStrBox",""
					DlgText "TranAfterStrBox",""
				End Select
			Case "TranBeforeStrButton","TranAfterStrButton"
				Call TranTest(DlgValue("EngineList"),1)
				Exit Function
			Case "ImportButton"
				If PSL.SelectFile(Path,True,Replace(MsgList(25),"%s",MsgList(14)),MsgList(23)) = False Then
					Exit Function
				End If
				n = GetEngineSet("Sets",Path)
				If n > 3 Then
					DlgListBoxArray "EngineList",EngineList()
					HeaderID = UBound(EngineList)
					Header = EngineList(HeaderID)
					For i = LBound(EngineListBak) To UBound(EngineListBak)
						If EngineListBak(i) = Header Then
							HeaderID = i
							Exit For
						End If
					Next i
					DlgValue "EngineList",HeaderID
					MainArray = ReSplit(EngineDataList(HeaderID),JoinStr)
					SetsArray = ReSplit(MainArray(1),SubJoinStr)
					SplitData(MainArray(2),LngNameList,SrcLngList,TranLngList)
					DlgItem$ = "EngineList"
					MsgBox MsgList(19),vbOkOnly+vbInformation,MsgList(17)
				ElseIf n = 0 Then
					MsgBox Replace$(MsgList(21),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
			Case "ExportButton"
				tStemp = CheckNullData("",EngineDataList,"1,6-9,15-19",1)
				If tStemp = False Then tStemp = CheckTargetValue(EngineDataList,-1)
				If tStemp = True Then
					If MsgBox(Replace$(MsgList(10),"%s",MsgList(14)) & MsgList(6),vbYesNo+vbInformation,MsgList(5)) = vbNo Then
						Exit Function
					End If
				End If
				If PSL.SelectFile(Path,False,Replace(MsgList(25),"%s",MsgList(14)),MsgList(24)) = True Then
					If InStr(Path,".dat") = 0 Then Path = Path & ".dat"
					If WriteEngineSet(Path,"All") = False Then
						MsgBox Replace$(MsgList(22),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
					Else
						MsgBox MsgList(18),vbOkOnly+vbInformation,MsgList(17)
					End If
				End If
				Exit Function
			Case "TestButton"
				tStemp = CheckNullData("",EngineDataList,"1,6-9,15-19",1)
				If tStemp = False Then tStemp = CheckTargetValue(EngineDataList,-1)
				If tStemp = True Then
					If MsgBox(Replace$(MsgList(10),"%s",MsgList(14)) & MsgList(6),vbYesNo+vbInformation,MsgList(5)) = vbNo Then
						Exit Function
					End If
				End If
				'转换翻译引擎的配置中的转义符
				TempDataList = EngineDataList
				For i = LBound(EngineDataList) To UBound(EngineDataList)
					MainArray = ReSplit(EngineDataList(i),JoinStr)
					SetsArray = ReSplit(MainArray(1),SubJoinStr)
					For j = 0 To UBound(SetsArray)
						If j > 10 And j < 19 Then
							If SetsArray(j) <> "" Then
								SetsArray(j) = Convert(SetsArray(j))
							End If
						End If
					Next j
					MainArray(1) = Join(SetsArray,SubJoinStr)
					EngineDataList(i) = Join(MainArray,JoinStr)
				Next i
				'转换检查配置中的转义符
				TempArray = CheckDataList
				For i = LBound(CheckDataList) To UBound(CheckDataList)
					MainArray = ReSplit(CheckDataList(i),JoinStr)
					SetsArray = ReSplit(MainArray(1),SubJoinStr)
					For j = 0 To UBound(SetsArray)
						If j <> 4 And j <> 14 And j <> 15 And j < 18 Then
							If SetsArray(j) <> "" Then
								If j = 1 Or j = 5 Or j = 13 Or j = 16 Or j = 17 Then
									If j = 5 Or j = 7 Then Temp = " " Else Temp = ","
									TempList = ReSplit(SetsArray(j),Temp,-1)
									Call SortArrayByLength(TempList,0,UBound(TempList),True)
									SetsArray(j) = Convert(Join(TempList,Temp))
								Else
									SetsArray(j) = Convert(SetsArray(j))
								End If
							End If
						End If
					Next j
					MainArray(1) = Join(SetsArray,SubJoinStr)
					CheckDataList(i) = Join(MainArray,JoinStr)
				Next i
				TempList = CheckDataListBak
				CheckDataListBak = TempArray
				Call TranTest(DlgValue("EngineList"),0)
				EngineDataList = TempDataList
				CheckDataList = TempArray
				CheckDataListBak = TempList
				Exit Function
			End Select

			If DlgItem$ = "EngineList" Then
				DlgText "ObjectNameBox",SetsArray(0)
				DlgText "AppIDBox",SetsArray(1)
				DlgText "UrlBox",SetsArray(2)
				DlgText "UrlTemplateBox",SetsArray(3)
				DlgText "bstrMethodBox",SetsArray(4)
				DlgText "varAsyncBox",SetsArray(5)
				DlgText "bstrUserBox",SetsArray(6)
				DlgText "bstrPasswordBox",SetsArray(7)
				DlgText "varBodyBox",SetsArray(8)
				DlgText "setRequestHeaderBox",SetsArray(9)
				DlgText "responseTypeBox",SetsArray(10)
				Select Case DlgText("responseTypeBox")
				Case "responseText"
					DlgText "TranBeforeStrBox",SetsArray(11)
					DlgText "TranAfterStrBox",SetsArray(12)
				Case "responseBody"
					DlgText "TranBeforeStrBox",SetsArray(13)
					DlgText "TranAfterStrBox",SetsArray(14)
				Case "responseStream"
					DlgText "TranBeforeStrBox",SetsArray(15)
					DlgText "TranAfterStrBox",SetsArray(16)
				Case "responseXML"
					DlgText "TranBeforeStrBox",SetsArray(17)
					DlgText "TranAfterStrBox",SetsArray(18)
				Case Else
					DlgText "TranBeforeStrBox",""
					DlgText "TranAfterStrBox",""
				End Select
				DlgValue "EnableBox",StrToLong(SetsArray(19))
				If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject
				n = 0
				Temp = DlgText("LngNameList")
				DlgListBoxArray "LngNameList",LngNameList()
				DlgListBoxArray "SrcLngList",SrcLngList()
				DlgListBoxArray "TranLngList",TranLngList()
				For i = 0 To UBound(LngNameList)
					If LngNameList(i) = Temp Then
						n = i
						Exit For
					End If
				Next i
				DlgValue "LngNameList",n
				DlgValue "SrcLngList",n
				DlgValue "TranLngList",n
				DlgEnable "ShowNoNullLngButton",True
				DlgEnable "ShowNullLngButton",True
				DlgEnable "ShowAllLngButton",False
			End If

			HeaderID = DlgValue("EngineList")
			'MainArray = ReSplit(EngineDataList(HeaderID),JoinStr)
			If DlgValue("Engine") <> 1 Then
				'SetsArray = ReSplit(MainArray(1),SubJoinStr)
				SetsArray(0) = DlgText("ObjectNameBox")
				SetsArray(1) = DlgText("AppIDBox")
				SetsArray(2) = DlgText("UrlBox")
				SetsArray(3) = DlgText("UrlTemplateBox")
				SetsArray(4) = DlgText("bstrMethodBox")
				SetsArray(5) = DlgText("varAsyncBox")
				SetsArray(6) = DlgText("bstrUserBox")
				SetsArray(7) = DlgText("bstrPasswordBox")
				SetsArray(8) = DlgText("varBodyBox")
				SetsArray(9) = DlgText("setRequestHeaderBox")
				SetsArray(10) = DlgText("responseTypeBox")
				Select Case DlgText("responseTypeBox")
				Case "responseText"
					SetsArray(11) = DlgText("TranBeforeStrBox")
					SetsArray(12) = DlgText("TranAfterStrBox")
				Case "responseBody"
					SetsArray(13) = DlgText("TranBeforeStrBox")
					SetsArray(14) = DlgText("TranAfterStrBox")
				Case "responseStream"
					SetsArray(15) = DlgText("TranBeforeStrBox")
					SetsArray(16) = DlgText("TranAfterStrBox")
				Case "responseXML"
					SetsArray(17) = DlgText("TranBeforeStrBox")
					SetsArray(18) = DlgText("TranAfterStrBox")
				End Select
				SetsArray(19) = DlgValue("EnableBox")
				MainArray(1) = Join(SetsArray,SubJoinStr)
			Else
				TempList = LngNameList
				For i = 0 To UBound(LngNameList)
					TempList(i) = LngNameList(i) & LngJoinStr & _
							IIf(SrcLngList(i) = NullValue,"",SrcLngList(i)) & _
							LngJoinStr & IIf(TranLngList(i) = NullValue,"",TranLngList(i))
				Next i
				MainArray(2) = Join(TempList,SubLngJoinStr)
			End If
			EngineDataList(HeaderID) = Join(MainArray,JoinStr)

			tStemp = False
			For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
				If DefaultEngineList(i) = EngineList(HeaderID) Then
					tStemp = True
					Exit For
				End If
			Next i
			If tStemp = True Then
				DlgEnable "ChangButton",False
				DlgEnable "DelButton",False
			Else
				DlgEnable "ChangButton",True
				DlgEnable "DelButton",IIf(UBound(EngineList) = 0,False,True)
			End If

			If DlgText("responseTypeBox") = "responseXML" Then
				DlgText "TranBeforeStrText",MsgList(29)
				DlgText "TranAfterStrText",MsgList(30)
			Else
				DlgText "TranBeforeStrText",MsgList(27)
				DlgText "TranAfterStrText",MsgList(28)
			End If

			If DlgValue("Engine") = 1 Then
				Temp = DlgText("LngNameList")
				SrcLngCode = IIf(DlgText("SrcLngList") = NullValue,"",DlgText("SrcLngList"))
				TranLngCode = IIf(DlgText("TranLngList") = NullValue,"",DlgText("TranLngList"))
				DlgEnable "NullLngButton",IIf(TranLngCode = "",False,True)
				DlgEnable "ResetLngButton",False
				For i = LBound(EngineDataListBak) To UBound(EngineDataListBak)
					TempArray = ReSplit(EngineDataListBak(i),JoinStr)
					If TempArray(0) = EngineList(HeaderID) Then
						TempList = ReSplit(TempArray(2),SubLngJoinStr)
						For n = 0 To UBound(TempList)
							LangPairList = ReSplit(TempList(n),LngJoinStr)
							If LangPairList(0) = Temp Then
								If LCase(LangPairList(2)) <> LCase(TranLngCode) Then
									DlgEnable "ResetLngButton",True
								ElseIf LCase(LangPairList(1)) <> LCase(SrcLngCode) Then
									DlgEnable "ResetLngButton",True
								End If
								Exit For
							End If
						Next n
						Exit For
					End If
				Next i
				If DlgText("LngNameList") = "" Then
					DlgText "LngNameText",Replace$(MsgList(50),"%s","0")
				ElseIf DlgEnable("ShowAllLngButton") = False Then
					DlgText "LngNameText",Replace$(MsgList(50),"%s",CStr$(UBound(LngNameList) + 1 ))
				Else
					DlgText "LngNameText",Replace$(MsgList(50),"%s",CStr$(UBound(DelLngNameList) + 1))
				End If
			End If
		Case 1
			HeaderID = DlgValue("CheckList")
			MainArray = ReSplit(CheckDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),SubJoinStr)
			getLngNameList(MainArray(2),AppLngList,UseLngList)
			Select Case DlgItem$
			Case "LevelButton"
				Header = CheckList(HeaderID)
				If SetLevel(CheckList,HeaderID,MsgList(51)) = True Then
					DlgListBoxArray "CheckList",CheckList()
					DlgText "CheckList",Header
					Set Dic = CreateObject("Scripting.Dictionary")
					For i = 0 To UBound(CheckDataList)
						TempList = ReSplit(CheckDataList(i),JoinStr)
						If Not Dic.Exists(TempList(0)) Then
							Dic.Add(TempList(0),i)
						End If
					Next i
					TempArray = CheckDataList
					For i = 0 To UBound(CheckList)
						If Dic.Exists(CheckList(i)) Then
							j = Dic.Item(CheckList(i))
							TempArray(n) = CheckDataList(j)
							n = n + 1
						End If
					Next i
					CheckDataList = TempArray
					Set Dic = Nothing
				End If
				Exit Function
			Case "AddButton"
				NewData = AddSet(CheckList)
				If NewData = "" Then Exit Function
				ReDim SetsArray(UBound(SetsArray)) As String
				Data = Join(SetsArray,SubJoinStr)
				LangPairList = LangCodeList("check",1,-1)
				Temp = NewData & JoinStr & Data & JoinStr & Join(LangPairList,SubLngJoinStr)
				CreateArray(NewData,Temp,CheckList,CheckDataList)
				DlgListBoxArray "CheckList",CheckList()
				DlgText "CheckList",NewData
				HeaderID = DlgValue("CheckList")
				MainArray = ReSplit(CheckDataList(HeaderID),JoinStr)
				SetsArray = ReSplit(MainArray(1),SubJoinStr)
				getLngNameList(MainArray(2),AppLngList,UseLngList)
				DlgItem$ = "CheckList"
			Case "ChangButton"
				NewData = EditSet(CheckList,HeaderID)
				If NewData <> "" Then
					CheckList(HeaderID) = NewData
					MainArray(0) = NewData
					CheckDataList(HeaderID) = Join(MainArray,JoinStr)
					DlgListBoxArray "CheckList",CheckList()
					DlgValue "CheckList",HeaderID
				End If
				Exit Function
	    	Case "DelButton"
				Header = CheckList(HeaderID)
				If MsgBox(Replace(MsgList(12),"%s",Header),vbYesNo+vbInformation,MsgList(11)) = vbNo Then
					Exit Function
				End If
				i = UBound(CheckList)
				Call DelArrays(CheckList,CheckDataList,HeaderID)
				If HeaderID > 0 And HeaderID = i Then HeaderID = HeaderID - 1
				DlgListBoxArray "CheckList",CheckList()
				DlgValue "CheckList",HeaderID
				MainArray = ReSplit(CheckDataList(HeaderID),JoinStr)
				SetsArray = ReSplit(MainArray(1),SubJoinStr)
				getLngNameList(MainArray(2),AppLngList,UseLngList)
				DlgItem$ = "CheckList"
			Case "ResetButton"
				Header = CheckList(HeaderID)
				ReDim TempArray(1)
				For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
					If DefaultCheckList(i) = Header Then
						TempArray(0) = MsgList(1)
						Exit For
					End If
				Next i
				cStemp = CheckNullData(Header,CheckDataListBak,"1,4,14-17",0)
				If cStemp = False Then TempArray(1) = MsgList(2)
				For i = LBound(CheckList) To UBound(CheckList)
					If i <> HeaderID Then
						ReDim Preserve TempArray(i + 2)
						TempArray(i + 2) = MsgList(3) & " - " & CheckList(i)
					End If
				Next i
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				If i = 0 Then
					SetsArray = ReSplit(CheckSettings(Header,0),SubJoinStr)
					TempArray = LangCodeList(Header,1,-1)
					Temp = Join(TempArray,SubLngJoinStr)
				ElseIf i = 1 Then
					For n = LBound(CheckDataListBak) To UBound(CheckDataListBak)
						TempArray = ReSplit(CheckDataListBak(n),JoinStr)
						If TempArray(0) = Header Then
							SetsArray = ReSplit(TempArray(1),SubJoinStr)
							Temp = TempArray(2)
							Exit For
						End If
					Next n
				ElseIf i >= 2 Then
					Temp = Mid(TempArray(i),InStr(TempArray(i),MsgList(3) & " - ") + Len(MsgList(3) & " - "))
					For n = LBound(CheckList) To UBound(CheckList)
						If CheckList(n) = Temp Then
							TempArray = ReSplit(CheckDataList(n),JoinStr)
							SetsArray = ReSplit(TempArray(1),SubJoinStr)
							Temp = TempArray(2)
							Exit For
						End If
					Next n
				End If
				Select Case DlgValue("SetType")
				Case 0
					DlgText "AccKeyBox",SetsArray(13)
					DlgText "ExCrBox",SetsArray(0)
					DlgText "ChkBktBox",SetsArray(2)
					DlgText "KpPairBox",SetsArray(3)
					DlgValue "AsiaKeyBox",StrToLong(SetsArray(4))
					DlgValue "AddAcckeyBox",StrToLong(SetsArray(14))
				Case 1
					DlgText "ChkEndBox",SetsArray(5)
					DlgText "NoTrnEndBox",SetsArray(6)
					DlgText "AutoTrnEndBox",SetsArray(7)
				Case 2
					DlgText "ShortBox",SetsArray(8)
					DlgText "ShortKeyBox",SetsArray(9)
					DlgText "KpShortKeyBox",SetsArray(10)
				Case 3
					DlgText "PreRepStrBox",SetsArray(11)
					DlgText "AutoWebFlagBox",SetsArray(12)
				Case 4
					DlgText "PreInsertSplitBox",SetsArray(1)
					DlgValue "LineSplitModeBox",StrToLong(SetsArray(15))
					DlgText "AppInsertSplitBox",SetsArray(16)
					DlgText "ReplaceSplitBox",SetsArray(17)
				Case 5
					MainArray(2) = Temp
					getLngNameList(Temp,AppLngList,UseLngList)
					DlgListBoxArray "AppLngList",AppLngList()
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "AppLngList",0
					DlgValue "UseLngList",0
				End Select
	    	Case "CleanButton"
	    		Select Case DlgValue("SetType")
	    		Case 0
	    			DlgText "AccKeyBox",""
					DlgText "ExCrBox",""
					DlgText "ChkBktBox",""
					DlgText "KpPairBox",""
					DlgValue "AsiaKeyBox",0
					DlgValue "AddAcckeyBox",0
				Case 1
					DlgText "ChkEndBox",""
					DlgText "NoTrnEndBox",""
					DlgText "AutoTrnEndBox",""
				Case 2
					DlgText "ShortBox",""
					DlgText "ShortKeyBox",""
					DlgText "KpShortKeyBox",""
				Case 3
					DlgText "PreRepStrBox",""
					DlgText "AutoWebFlagBox",""
				Case 4
					DlgText "PreInsertSplitBox",""
					DlgValue "LineSplitModeBox",0
					DlgText "AppInsertSplitBox",""
					DlgText "ReplaceSplitBox",""
				Case 5
					ReDim UseLngList(0),AppLngList(0)
					MainArray(2) = LngJoinStr & LngJoinStr
					'AppLngList = ChangeList(MainArray(2),UseLngList)
					DlgListBoxArray "AppLngList",AppLngList()
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "AppLngList",0
					DlgValue "UseLngList",0
				End Select
			Case "AddLangButton","DelLangButton"
				If DlgItem$ = "AddLangButton" Then
					LngName = DlgText("AppLngList")
					If LngName = "" Then Exit Function
					n = DlgValue("AppLngList")
					i = UBound(AppLngList)
					Call DelArray(AppLngList,n)
					UseLngList = ChangeList(MainArray(2),AppLngList)
				Else
					LngName = DlgText("UseLngList")
					If LngName = "" Then Exit Function
					n = DlgValue("UseLngList")
					i = UBound(UseLngList)
					Call DelArray(UseLngList,n)
					AppLngList = ChangeList(MainArray(2),UseLngList)
				End If
				If n > 0 And n = i Then n = n - 1
				DlgListBoxArray "AppLngList",AppLngList()
				DlgListBoxArray "UseLngList",UseLngList()
				If DlgItem$ = "AddLangButton" Then
					DlgValue "AppLngList",n
					DlgText "UseLngList",LngName
				Else
					DlgText "AppLngList",LngName
					DlgValue "UseLngList",n
				End If
			Case "AddAllLangButton","DelAllLangButton"
				If DlgItem$ = "AddAllLangButton" Then
					ReDim AppLngList(0)
					UseLngList = ChangeList(MainArray(2),AppLngList)
				Else
					ReDim UseLngList(0)
					AppLngList = ChangeList(MainArray(2),UseLngList)
				End If
				DlgListBoxArray "AppLngList",AppLngList()
				DlgListBoxArray "UseLngList",UseLngList()
				DlgValue "AppLngList",0
				DlgValue "UseLngList",0
			Case "SetAppLangButton","SetUseLangButton"
				TempList = ReSplit(MainArray(2),SubLngJoinStr)
				NewData = SetLang(TempList,"","")
				If NewData = "" Then Exit Function
				n = UBound(TempList)
				If n > 0 Or TempList(n) <> LngJoinStr & LngJoinStr Then n = n + 1
				ReDim Preserve TempList(n)
				If DlgItem$ = "SetAppLangButton" Then
					TempList(n) = NewData
				Else
					TempList(n) = NewData & ReSplit(NewData,LngJoinStr)(1)
				End If
				MainArray(2) = Join(TempList,SubLngJoinStr)
				getLngNameList(MainArray(2),AppLngList,UseLngList)
				If DlgItem$ = "SetAppLangButton" Then
					n = UBound(AppLngList)
					DlgListBoxArray "AppLngList",AppLngList()
					DlgValue "AppLngList",n
				Else
					n = UBound(UseLngList)
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "UseLngList",n
				End If
			Case "EditAppLangButton","EditUseLangButton"
				If DlgItem$ = "EditAppLangButton" Then
					LngName = DlgText("AppLngList")
					n = DlgValue("AppLngList")
				Else
					LngName = DlgText("UseLngList")
					n = DlgValue("UseLngList")
				End If
				If LngName = "" Then Exit Function
				j = -1
				TempList = ReSplit(MainArray(2),SubLngJoinStr)
				For i = 0 To UBound(TempList)
					LangPairList = ReSplit(TempList(i),LngJoinStr)
					If LangPairList(0) = LngName Then
						LngCode = LangPairList(1)
						j = i
						Exit For
					End If
				Next i
				NewData = SetLang(TempList,LngName,LngCode)
				If NewData = "" Then Exit Function
				If j > -1 Then
					TempList(j) = NewData
				Else
					j = UBound(TempList) + 1
					ReDim Preserve TempList(j)
					TempList(j) = NewData
				End If
				MainArray(2) = Join(TempList,SubLngJoinStr)
				NewData = ReSplit(NewData,LngJoinStr)(0)
				If DlgItem$ = "EditAppLangButton" Then
					AppLngList(n) = NewData
					DlgListBoxArray "AppLngList",AppLngList()
					DlgValue "AppLngList",n
				Else
					UseLngList(n) = NewData
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "UseLngList",n
				End If
			Case "DelAppLangButton","DelUseLangButton"
				If DlgItem$ = "DelAppLangButton" Then
					LngName = DlgText("AppLngList")
					n = DlgValue("AppLngList")
					i = UBound(AppLngList)
				Else
					LngName = DlgText("UseLngList")
					n = DlgValue("UseLngList")
					i = UBound(UseLngList)
				End If
				If LngName = "" Then Exit Function
				If MsgBox(Replace(MsgList(13),"%s",LngName),vbYesNo+vbInformation,MsgList(11)) = vbNo Then
					Exit Function
				End If
				TempList = ReSplit(MainArray(2),SubLngJoinStr)
				Call DelArray(TempList,LngName,LngJoinStr)
				MainArray(2) = Join(TempList,SubLngJoinStr)
				getLngNameList(MainArray(2),AppLngList,UseLngList)
				If n > 0 And n = i Then n = n - 1
				If DlgItem$ = "DelAppLangButton" Then
					DlgListBoxArray "AppLngList",AppLngList()
					DlgValue "AppLngList",n
				Else
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "UseLngList",n
				End If
			Case "ImportButton"
				If PSL.SelectFile(Path,True,Replace(MsgList(25),"%s",MsgList(15)),MsgList(23)) = False Then
					Exit Function
				End If
				n = GetCheckSet("Sets",Path)
				If n = 4 Then
					DlgListBoxArray "CheckList",CheckList()
					HeaderID = UBound(CheckList)
					Header = CheckList(HeaderID)
					For i = LBound(CheckListBak) To UBound(CheckListBak)
						If CheckListBak(i) = Header Then
							HeaderID = i
							Exit For
						End If
					Next i
					DlgValue "CheckList",HeaderID
					MainArray = ReSplit(CheckDataList(HeaderID),JoinStr)
					SetsArray = ReSplit(MainArray(1),SubJoinStr)
					getLngNameList(MainArray(2),AppLngList,UseLngList)
					DlgItem$ = "CheckList"
					MsgBox MsgList(19),vbOkOnly+vbInformation,MsgList(17)
				ElseIf n = 0 Then
					MsgBox Replace$(MsgList(21),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
			Case "ExportButton"
				If CheckNullData("",CheckDataList,"1,4,14-17",1) = True Then
					If MsgBox(MsgList(10) & MsgList(6),vbYesNo+vbInformation,MsgList(5)) = vbNo Then
						Exit Function
					End If
				End If
				If PSL.SelectFile(Path,False,Replace(MsgList(25),"%s",MsgList(15)),MsgList(24)) = True Then
					If InStr(Path,".dat") = 0 Then Path = Path & ".dat"
					If WriteCheckSet(Path,"All") = False Then
						MsgBox Replace$(MsgList(22),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
					Else
						MsgBox MsgList(18),vbOkOnly+vbInformation,MsgList(17)
					End If
				End If
				Exit Function
			Case "TestButton"
				If CheckNullData("",CheckDataList,"1,4,14-17",1) = True Then
					If MsgBox(Replace(MsgList(10),"%s",MsgList(15)) & MsgList(6),vbYesNo+vbInformation,MsgList(5)) = vbNo Then
						Exit Function
					End If
				End If
				'转换检查配置中的转义符
				TempArray = CheckDataList
				For i = LBound(CheckDataList) To UBound(CheckDataList)
					MainArray = ReSplit(CheckDataList(i),JoinStr)
					SetsArray = ReSplit(MainArray(1),SubJoinStr)
					For j = 0 To UBound(SetsArray)
						If j <> 4 And j <> 14 And j <> 15 And j < 18 Then
							If SetsArray(j) <> "" Then
								If j = 1 Or j = 5 Or j = 13 Or j = 16 Or j = 17 Then
									If j = 5 Or j = 7 Then Temp = " " Else Temp = ","
									TempList = ReSplit(SetsArray(j),Temp,-1)
									Call SortArrayByLength(TempList,0,UBound(TempList),True)
									SetsArray(j) = Convert(Join(TempList,Temp))
								Else
									SetsArray(j) = Convert(SetsArray(j))
								End If
							End If
						End If
					Next j
					MainArray(1) = Join(SetsArray,SubJoinStr)
					CheckDataList(i) = Join(MainArray,JoinStr)
				Next i
				TempList = CheckDataListBak
				CheckDataListBak = TempArray
				Call CheckTest(DlgValue("CheckList"),0)
				CheckDataList = TempArray
				CheckDataListBak = TempList
				Exit Function
			End Select

			If DlgItem$ = "CheckList" Then
				DlgText "ExCrBox",SetsArray(0)
				DlgText "PreInsertSplitBox",SetsArray(1)
				DlgText "ChkBktBox",SetsArray(2)
				DlgText "KpPairBox",SetsArray(3)
				DlgValue "AsiaKeyBox",StrToLong(SetsArray(4))
				DlgText "ChkEndBox",SetsArray(5)
				DlgText "NoTrnEndBox",SetsArray(6)
				DlgText "AutoTrnEndBox",SetsArray(7)
				DlgText "ShortBox",SetsArray(8)
				DlgText "ShortKeyBox",SetsArray(9)
				DlgText "KpShortKeyBox",SetsArray(10)
				DlgText "PreRepStrBox",SetsArray(11)
				DlgText "AutoWebFlagBox",SetsArray(12)
				DlgText "AccKeyBox",SetsArray(13)
				DlgValue "AddAcckeyBox",StrToLong(SetsArray(14))
				DlgValue "LineSplitModeBox",StrToLong(SetsArray(15))
				DlgText "AppInsertSplitBox",SetsArray(16)
				DlgText "ReplaceSplitBox",SetsArray(17)
				DlgListBoxArray "AppLngList",AppLngList()
				DlgListBoxArray "UseLngList",UseLngList()
				DlgValue "AppLngList",0
				DlgValue "UseLngList",0
			End If

			HeaderID = DlgValue("CheckList")
			'MainArray = ReSplit(CheckDataList(HeaderID),JoinStr)
			If DlgValue("SetType") <> 5 Then
				SetsArray = ReSplit(MainArray(1),SubJoinStr)
				SetsArray(0) = DlgText("ExCrBox")
				SetsArray(1) = DlgText("PreInsertSplitBox")
				SetsArray(2) = DlgText("ChkBktBox")
				SetsArray(3) = DlgText("KpPairBox")
				SetsArray(4) = DlgValue("AsiaKeyBox")
				SetsArray(5) = DlgText("ChkEndBox")
				SetsArray(6) = DlgText("NoTrnEndBox")
				SetsArray(7) = DlgText("AutoTrnEndBox")
				SetsArray(8) = DlgText("ShortBox")
				SetsArray(9) = DlgText("ShortKeyBox")
				SetsArray(10) = DlgText("KpShortKeyBox")
				SetsArray(11) = DlgText("PreRepStrBox")
				SetsArray(12) = DlgText("AutoWebFlagBox")
				SetsArray(13) = DlgText("AccKeyBox")
				SetsArray(14) = DlgValue("AddAcckeyBox")
				SetsArray(15) = DlgValue("LineSplitModeBox")
				SetsArray(16) = DlgText("AppInsertSplitBox")
				SetsArray(17) = DlgText("ReplaceSplitBox")
				MainArray(1) = Join(SetsArray,SubJoinStr)
			Else
				Set Dic = CreateObject("Scripting.Dictionary")
				For i = LBound(UseLngList) To UBound(UseLngList)
					If Not Dic.Exists(UseLngList(i)) Then
						Dic.Add(UseLngList(i),"")
					End If
				Next i
				TempList = ReSplit(MainArray(2),SubLngJoinStr)
				For i = 0 To UBound(TempList)
					LangPairList = ReSplit(TempList(i),LngJoinStr)
					If Dic.Exists(LangPairList(0)) Then
						TempList(i) = LangPairList(0) & LngJoinStr & LangPairList(1) & LngJoinStr & LangPairList(1)
					Else
						TempList(i) = LangPairList(0) & LngJoinStr & LangPairList(1) & LngJoinStr
					End If
				Next i
				Set Dic = Nothing
				MainArray(2) = Join(TempList,SubLngJoinStr)
			End If
			CheckDataList(HeaderID) = Join(MainArray,JoinStr)

			If DlgText("AppLngList") = "" Then
				DlgEnable "AddLangButton",False
				DlgEnable "AddAllLangButton",False
				DlgEnable "EditAppLangButton",False
				DlgEnable "DelAppLangButton",False
				DlgText "AppLngText",Replace$(MsgList(31),"%s","0")
			Else
				DlgEnable "AddLangButton",True
				DlgEnable "AddAllLangButton",True
				DlgEnable "EditAppLangButton",True
				DlgEnable "DelAppLangButton",True
				DlgText "AppLngText",Replace$(MsgList(31),"%s",CStr$(UBound(AppLngList) + 1))
			End If
			If DlgText("UseLngList") = "" Then
				DlgEnable "DelLangButton",False
				DlgEnable "DelAllLangButton",False
				DlgEnable "EditUseLangButton",False
				DlgEnable "DelUseLangButton",False
				DlgText "UseLngText",Replace$(MsgList(32),"%s","0")
			Else
				DlgEnable "DelLangButton",True
				DlgEnable "DelAllLangButton",True
				DlgEnable "EditUseLangButton",True
				DlgEnable "DelUseLangButton",True
				DlgText "UseLngText",Replace$(MsgList(32),"%s",CStr$(UBound(UseLngList) + 1))
			End If
			If DlgValue("AsiaKeyBox") = 1 Then
				DlgEnable "AddAcckeyBox",False
				DlgValue "AddAcckeyBox",0
			Else
				DlgEnable "AddAcckeyBox",True
			End If

			cStemp = False
			For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
				If DefaultCheckList(i) = CheckList(HeaderID) Then
					cStemp = True
					Exit For
				End If
			Next i
			If cStemp = True Then
				DlgEnable "ChangButton",False
				DlgEnable "DelButton",False
			Else
				DlgEnable "ChangButton",True
				DlgEnable "DelButton",IIf(UBound(CheckList) = 0,False,True)
			End If
		Case 2
			Select Case DlgItem$
			Case "ExeBrowseButton"
				If PSL.SelectFile(Path,True,MsgList(34),MsgList(33)) = False Then
					Exit Function
				End If
				DlgText "CmdPathBox",Path
				If InStr(LCase$(Path),"winrar.exe") Then
					DlgText "ArgumentBox","e -ibck ""%1"" %2 ""%3"""
				ElseIf InStr(LCase$(Path),"winzip.exe") Then
					DlgText "ArgumentBox"," ""%1"" %2 ""%3"""
				ElseIf InStr(LCase$(Path),"7z.exe") Then
					DlgText "ArgumentBox","e ""%1"" -o""%3"" %2"
				ElseIf InStr(LCase$(Path),"haozip.exe") Then
					DlgText "ArgumentBox","e ""%1"" -r -o""%3"" %2"
				ElseIf InStr(LCase$(Path),"haozipc.exe") Then
					DlgText "ArgumentBox","e ""%1"" -r -o""%3"" %2"
				End If
			Case "ArgumentButton"
				ReDim TempArray(2) As String
				TempArray(0) = MsgList(36)
				TempArray(1) = MsgList(37)
				TempArray(2) = MsgList(38)
				i = ShowPopupMenu(TempArray)
				If i < 0 Then Exit Function
				If i = 0 Then
					DlgText "ArgumentBox",DlgText("ArgumentBox")  & " " & """%1"""
				ElseIf i = 1 Then
					DlgText "ArgumentBox",DlgText("ArgumentBox")  & " " & """%2"""
				ElseIf i = 2 Then
					DlgText "ArgumentBox",DlgText("ArgumentBox")  & " " & """%3"""
				End If
			Case "ResetButton"
				ReDim TempArray(1) As String
				TempArray(0) = MsgList(1)
				TempArray(1) = MsgList(2)
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i < 0 Then Exit Function
				If i = 0 Then
					Temp = updateMainUrl & vbCrLf & updateMinorUrl
					DlgValue "UpdateSet",1
					DlgText "WebSiteBox",Temp
					Call getCMDPath(".rar",Path,Temp)
					DlgText "CmdPathBox",Path
					DlgText "ArgumentBox",Temp
					DlgText "UpdateCycleBox","7"
				ElseIf i = 1 Then
					DlgValue "UpdateSet",StrToLong(tUpdateSetBak(0))
					DlgText "WebSiteBox",tUpdateSetBak(1)
					DlgText "CmdPathBox",tUpdateSetBak(2)
					DlgText "ArgumentBox",tUpdateSetBak(3)
					DlgText "UpdateCycleBox",tUpdateSetBak(4)
				End If
	    	Case "CleanButton"
	    		DlgText "WebSiteBox",""
	    		DlgText "CmdPathBox",""
				DlgText "ArgumentBox",""
	    	Case "CheckButton"
	    		If DlgText("CmdPathBox") = "" Or DlgText("ArgumentBox") = "" Then
	    			MsgBox(MsgList(35),vbOkOnly+vbInformation,MsgList(0))
					Exit Function
	    		End If
				i = Download(tUpdateSet,DlgText("WebSiteBox"),3)
	    		If i > 0 Then
	    			tStemp = False
	    			If tUpdateSet(5) < Format(Date,"yyyy-MM-dd") Then
	    				tStemp = True
	    			ElseIf i = 3 And ArrayComp(tUpdateSet,tUpdateSetBak) = False Then
	    				tStemp = True
	    			End If
	    			If tStemp = True Then
						DlgText "UpdateDateBox",Format(Date,"yyyy-MM-dd")
						tUpdateSet(5) = DlgText("UpdateDateBox")
						Path = IIf(DlgValue("tWriteType") = 0,EngineFilePath,EngineRegKey)
						If WriteEngineSet(Path,"Update") = False Then
							MsgBox Replace$(MsgList(20),"%s",Path),vbOkOnly+vbInformation,MsgList(0)
							Exit Function
						Else
							tUpdateSetBak(5) = tUpdateSet(5)
						End If
					End If
					If i = 3 Then Call ExitMacro(1)
				End If
			Case "TestButton"
				If DlgText("CmdPathBox") = "" Or DlgText("ArgumentBox") = "" Then
	    			MsgBox(MsgList(35),vbOkOnly+vbInformation,MsgList(0))
	    			Exit Function
	    		End If
	    		Download(tUpdateSet,DlgText("WebSiteBox"),4)
	    		Exit Function
	    	End Select
			tUpdateSet(0) = DlgValue("UpdateSet")
			tUpdateSet(1) = DlgText("WebSiteBox")
			tUpdateSet(2) = DlgText("CmdPathBox")
			tUpdateSet(3) = DlgText("ArgumentBox")
			tUpdateSet(4) = DlgText("UpdateCycleBox")
			tUpdateSet(5) = DlgText("UpdateDateBox")
			If tUpdateSet(5) = MsgList(4) Then tUpdateSet(5) = ""
		Case 3
			Select Case DlgItem$
			Case "UILangList"
				HeaderID = DlgValue("UILangList")
				If HeaderID < 0 Then Exit Function
				If HeaderID > 1 Then
					If LCase$(tSelected(0)) = LCase$(UIFileList(HeaderID - 2).LangID) Then Exit Function
				Else
					If tSelected(0) = CStr$(HeaderID) Then Exit Function
				End If
			Case "EditUILangButton"
				ReDim FileDataList(UBound(UIFileList)) As String
				For i = 0 To UBound(UIFileList)
					FileDataList(i) = UIFileList(i).FilePath & JoinStr & "unicodeFFFE"
				Next i
				If EditFile(LangFile,FileDataList,True) = False Then Exit Function
				If GetUIList(MacroDir & "\Data\",UIFileList) = False Then Exit Function
				HeaderID = DlgValue("UILangList")
			Case "MainFontButton","SrcStrFontButton","TrnStrFontButton"
				HeaderID = CLng(DlgText("SuppValueBox"))
				Select Case DlgItem$
				Case "MainFontButton"
					i = 0
				Case "SrcStrFontButton"
					i = 1
				Case "TrnStrFontButton"
					i = 2
				End Select
				ReDim TempArray(2) As String
				TempArray(0) = MsgList(1)
				TempArray(1) = MsgList(2)
				TempArray(2) = MsgList(61)
				Select Case ShowPopupMenu(TempArray,vbPopupUseRightButton)
				Case Is < 0
					Exit Function
				Case 0	'默认字体
					If CheckFont(LFList(i)) = False Then Exit Function
					ReDim tmpLFList(0) As LOG_FONT
					LFList(i) = tmpLFList(0)
				Case 1	'原值字体
					If FontComp(LFList(i),LFListBak(i)) = False Then Exit Function
					LFList(i) = LFListBak(i)
				Case 2	'自定义字体
					If SelectFont(HeaderID,LFList(i)) = 0 Then Exit Function
				End Select
				DlgText Replace$(DlgItem$,"Button","Box"),GetFontText(HeaderID,LFList(i))
				If DlgItem$ <> "MainFontButton" Then Exit Function
				'创建字体并应用到对话框
				j = CreateFont(HeaderID,LFList(0))
				If j = 0 Then Exit Function
				For i = 0 To DlgCount() - 1
					SendMessageLNG(GetDlgItem(HeaderID,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
				Next i
				DrawWindow(HeaderID,j)
				Exit Function
			Case Else
				Exit Function
			End Select

			If HeaderID > 1 Then
				Path = UIFileList(HeaderID - 2).FilePath
			Else
				If HeaderID = 0 Then
					n = Val("&H" & OSLanguage)
				Else
					n = PSL.Option(pslOptionSystemLanguage)
				End If
				For i = 0 To UBound(UIFileList)
					TempArray = ReSplit(LCase$(UIFileList(i).LangID),";")
					For j = 0 To UBound(TempArray)
						If Val("&H" & TempArray(j)) = n Then
							Path = UIFileList(i).FilePath
							Exit For
						End If
					Next j
					If Path <> "" Then Exit For
				Next i
			End If
			If Path = "" Then Exit Function
			If Dir$(Path) = "" Then Exit Function
			If getINIFile(UIDataList,Path,"unicodeFFFE",0) = False Then Exit Function
			If DlgItem$ = "UILangList" Then
				If HeaderID < 2 Then
					tSelected(0) = CStr$(HeaderID)
				Else
					tSelected(0) = UIFileList(HeaderID - 2).LangID
				End If
			End If
			LangFile = Path

			'重新获取字符代码列表
			On Error Resume Next
			Set objStream = CreateObject("Adodb.Stream")
			On Error GoTo 0
			If objStream Is Nothing Then
				CodeList = getCodePageList(0,0)
			Else
				CodeList = getCodePageList(0,49)
			End If
			'重置设置对话框字串
			If getMsgList(UIDataList,MsgList,"Settings",1) = True Then
				DlgText -1,MsgList(0)
				DlgText "MainText",MsgList(1)
				DlgText "StrHandle",MsgList(2)
				DlgText "AutoUpdate",MsgList(3)
				DlgText "UILangListSet",MsgList(4)

				DlgText "GroupBox1",MsgList(5)
				DlgText "LevelButton",MsgList(6)
				DlgText "AddButton",MsgList(7)
				DlgText "ChangButton",MsgList(8)
				DlgText "DelButton",MsgList(9)

				DlgText "GroupBox2",MsgList(10)
				DlgText "cWriteToFile",MsgList(11)
				DlgText "cWriteToRegistry",MsgList(12)
				DlgText "ImportButton",MsgList(13)
				DlgText "ExportButton",MsgList(14)

				DlgText "GroupBox3",MsgList(15)
				DlgText "SetItemText",MsgList(16)
				DlgText "NextSetType",MsgList(17)

				DlgText "AccKeyBoxTxt",MsgList(18)
				DlgText "ExCrBoxTxt",MsgList(19)
				DlgText "ChkBktBoxTxt",MsgList(20)
				DlgText "KpPairBoxTxt",MsgList(21)
				DlgText "AsiaKeyBox",MsgList(22)
				DlgText "AddAcckeyBox",MsgList(23)

				DlgText "ChkEndBoxTxt",MsgList(24)
				DlgText "NoTrnEndBoxTxt",MsgList(25)
				DlgText "AutoTrnEndBoxTxt",MsgList(26)

				DlgText "ShortBoxTxt",MsgList(27)
				DlgText "ShortKeyBoxTxt",MsgList(28)
				DlgText "KpShortKeyBoxTxt",MsgList(29)

				DlgText "PreRepStrBoxTxt",MsgList(30)
				DlgText "AutoWebFlagBoxTxt",MsgList(31)

				DlgText "LineSplitBoxTxt",MsgList(32)
				DlgText "PreInsertSplitBoxTxt",MsgList(33)
				DlgText "AppInsertSplitBoxTxt",MsgList(34)
				DlgText "ReplaceSplitBoxTxt",MsgList(35)
				DlgText "LineSplitModeBox",MsgList(36)

				DlgText "AppLngText",MsgList(37)
				DlgText "UseLngText",MsgList(38)
				DlgText "AddLangButton",MsgList(39)
				DlgText "AddAllLangButton",MsgList(40)
				DlgText "DelLangButton",MsgList(41)
				DlgText "DelAllLangButton",MsgList(42)
				DlgText "SetAppLangButton",MsgList(43)
				DlgText "EditAppLangButton",MsgList(44)
				DlgText "DelAppLangButton",MsgList(45)
				DlgText "SetUseLangButton",MsgList(46)
				DlgText "EditUseLangButton",MsgList(47)
				DlgText "DelUseLangButton",MsgList(48)

				DlgText "UpdateSetGroup",MsgList(49)
				DlgText "AutoButton",MsgList(50)
				DlgText "ManualButton",MsgList(51)
				DlgText "OffButton",MsgList(52)
				DlgText "CheckGroup",MsgList(53)
				DlgText "UpdateCycleText",MsgList(54)
				DlgText "UpdateDatesText",MsgList(55)
				DlgText "UpdateDateText",MsgList(56)
				DlgText "CheckButton",MsgList(57)
				DlgText "WebSiteGroup",MsgList(58)
				DlgText "CmdGroup",MsgList(59)
				DlgText "CmdPathBoxText",MsgList(60)
				DlgText "ArgumentBoxText",MsgList(61)
				DlgText "ExeBrowseButton",MsgList(6)
				DlgText "ArgumentButton",MsgList(62)

				DlgText "UILangSetGroup",MsgList(63)
				DlgText "UILangSetText1",MsgList(64)
				DlgText "UILangSetText2",MsgList(65)
				DlgText "UILangSetText3",MsgList(66)

				DlgText "UIFontSetGroup",MsgList(67)
				DlgText "MainFontText",MsgList(68)
				DlgText "MainFontButton",MsgList(69)

				DlgText "HelpButton",MsgList(70)
				DlgText "ResetButton",MsgList(71)
				DlgText "TestButton",MsgList(72)
				DlgText "CleanButton",MsgList(73)
				DlgText "EditUILangButton",MsgList(74)

				DlgText "TrnEngine",MsgList(75)
				DlgText "tWriteToFile",MsgList(11)
				DlgText "tWriteToRegistry",MsgList(12)

				DlgText "GroupBox4",MsgList(15)
				DlgText "EngineArgument",MsgList(76)
				DlgText "LangCodePair",MsgList(77)
				DlgText "EngineEnable",MsgList(78)

				DlgText "ObjectNameText",MsgList(79)
				DlgText "AppIDText",MsgList(80)
				DlgText "UrlText",MsgList(81)
				DlgText "UrlTemplateText",MsgList(82)
				DlgText "bstrMethodText",MsgList(83)
				DlgText "varAsyncText",MsgList(84)
				DlgText "bstrUserText",MsgList(85)
				DlgText "bstrPasswordText",MsgList(86)
				DlgText "varBodyText",MsgList(87)
				DlgText "setRequestHeaderText",MsgList(88)
				DlgText "setRequestHeaderText2",MsgList(89)
				DlgText "responseTypeText",MsgList(90)
				DlgText "TranBeforeStrText",MsgList(91)
				DlgText "TranAfterStrText",MsgList(92)
				DlgText "UrlTemplateButton",MsgList(69)
				DlgText "bstrMethodButton",MsgList(69)
				DlgText "varAsyncButton",MsgList(69)
				DlgText "varBodyButton",MsgList(69)
				DlgText "RequestButton",MsgList(69)
				DlgText "responseTypeButton",MsgList(69)
				DlgText "TranBeforeStrButton",MsgList(69)
				DlgText "TranAfterStrButton",MsgList(69)

				DlgText "LngNameText",MsgList(93)
				DlgText "SrcLngText",MsgList(94)
				DlgText "TranLngText",MsgList(95)
				DlgText "AddLngButton",MsgList(7)
				DlgText "DelLngButton",MsgList(9)
				DlgText "DelAllButton",MsgList(96)
				DlgText "EditLngButton",MsgList(97)
				DlgText "ExtEditButton",MsgList(98)
				DlgText "NullLngButton",MsgList(99)
				DlgText "ResetLngButton",MsgList(100)
				DlgText "ShowNoNullLngButton",MsgList(101)
				DlgText "ShowNullLngButton",MsgList(102)
				DlgText "ShowAllLngButton",MsgList(103)

				DlgText "EnableText1",MsgList(104)
				DlgText "EnableText2",MsgList(105)
				DlgText "EnableBox",MsgList(106)
			End If
			'重置设置对话框字串
			If getMsgList(UIDataList,MsgList,"SettingsDlgFunc",1) = True Then
				ReDim TempList(5) As String
				For i = 0 To 5
					TempList(i) = MsgList(i + 53)
				Next i
				i = DlgValue("SetType")
				DlgListBoxArray "SetType",TempList()
				DlgValue "SetType",i

				ReDim TempList(UBound(UIFileList) + 2) As String
				For i = 0 To UBound(UIFileList) + 2
					If i < 2 Then
						TempList(i) = MsgList(i + 59)
					Else
						TempList(i) = UIFileList(i - 2).LangName
					End If
				Next i
				DlgListBoxArray "UILangList",TempList()
				DlgValue "UILangList",HeaderID
			End If
			'重置主对话框字串
			If getMsgList(UIDataList,MsgList,"Main",1) = True Then
				'初始化编辑工具菜单名称
				Tools(0).sName = MsgList(96)
				Tools(1).sName = MsgList(97)
				Tools(2).sName = MsgList(98)
				Tools(3).sName = MsgList(99)
				'替换配置名称
				For i = LBound(CheckList) To UBound(CheckList)
					n = 0
					Select Case CheckList(i)
					Case DefaultCheckList(0)
						n = 70
					Case DefaultCheckList(1)
						n = 71
					End Select
					If n <> 0 Then
						CheckList(i) = MsgList(n)
						TempArray = ReSplit(CheckDataList(i),JoinStr)
						TempArray(0) = MsgList(n)
						CheckDataList(i) = Join(TempArray,JoinStr)
					End If
				Next i
				n = DlgValue("CheckList")
				DlgListBoxArray "CheckList",CheckList()
				DlgValue "CheckList",n
				'替换备份配置名称
				For i = LBound(CheckListBak) To UBound(CheckListBak)
					n = 0
					Select Case CheckListBak(i)
					Case DefaultCheckList(0)
						n = 70
					Case DefaultCheckList(1)
						n = 71
					End Select
					If n <> 0 Then
						CheckListBak(i) = MsgList(n)
						TempArray = ReSplit(CheckDataListBak(i),JoinStr)
						TempArray(0) = MsgList(n)
						CheckDataListBak(i) = Join(TempArray,JoinStr)
					End If
				Next i
				'替换方案名称
				For i = LBound(ProjectList) To UBound(ProjectList)
					n = 0
					Select Case ProjectList(i)
					Case DefaultProjectList(0)
						n = 89
					Case DefaultProjectList(1)
						n = 90
					Case DefaultProjectList(2)
						n = 91
					Case DefaultProjectList(3)
						n = 92
					Case DefaultProjectList(4)
						n = 93
					End Select
					If n <> 0 Then
						ProjectList(i) = MsgList(n)
						TempArray = ReSplit(ProjectDataList(i),JoinStr)
						TempArray(0) = MsgList(n)
						ProjectDataList(i) = Join(TempArray,JoinStr)
					End If
				Next i
				'替换选项配置名称
				Select Case tSelected(2)
				Case DefaultCheckList(0)
					tSelected(2) = MsgList(70)
				Case DefaultCheckList(1)
					tSelected(2) = MsgList(71)
				End Select
				'重置默认检查列表
				DefaultCheckList(0) = MsgList(70)
				DefaultCheckList(1) = MsgList(71)
				'重置默认方案列表
				DefaultProjectList(0) = MsgList(89)
				DefaultProjectList(1) = MsgList(90)
				DefaultProjectList(2) = MsgList(91)
				DefaultProjectList(3) = MsgList(92)
				DefaultProjectList(4) = MsgList(93)
			End If
		End Select
	Case 3 ' 文本框或者组合框文本被更改
		Select Case DlgValue("Options")
		Case 0
			If getMsgList(UIDataList,MsgList,"SettingsDlgFunc",1) = False Then Exit Function
			HeaderID = DlgValue("EngineList")
			MainArray = ReSplit(EngineDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),SubJoinStr)
			SetsArray(0) = DlgText("ObjectNameBox")
			SetsArray(1) = DlgText("AppIDBox")
			SetsArray(2) = DlgText("UrlBox")
			SetsArray(3) = DlgText("UrlTemplateBox")
			SetsArray(4) = DlgText("bstrMethodBox")
			SetsArray(5) = DlgText("varAsyncBox")
			SetsArray(6) = DlgText("bstrUserBox")
			SetsArray(7) = DlgText("bstrPasswordBox")
			SetsArray(8) = DlgText("varBodyBox")
			SetsArray(9) = DlgText("setRequestHeaderBox")
			SetsArray(10) = DlgText("responseTypeBox")
			Select Case DlgText("responseTypeBox")
			Case "responseText"
				SetsArray(11) = DlgText("TranBeforeStrBox")
				SetsArray(12) = DlgText("TranAfterStrBox")
			Case "responseBody"
				SetsArray(13) = DlgText("TranBeforeStrBox")
				SetsArray(14) = DlgText("TranAfterStrBox")
			Case "responseStream"
				SetsArray(15) = DlgText("TranBeforeStrBox")
				SetsArray(16) = DlgText("TranAfterStrBox")
			Case "responseXML"
				SetsArray(17) = DlgText("TranBeforeStrBox")
				SetsArray(18) = DlgText("TranAfterStrBox")
			End Select
			SetsArray(19) = DlgValue("EnableBox")
			MainArray(1) = Join(SetsArray,SubJoinStr)
			EngineDataList(HeaderID) = Join(MainArray,JoinStr)
			If DlgText("responseTypeBox") = "responseXML" Then
				DlgText "TranBeforeStrText",MsgList(29)
				DlgText "TranAfterStrText",MsgList(30)
			Else
				DlgText "TranBeforeStrText",MsgList(27)
				DlgText "TranAfterStrText",MsgList(28)
			End If
		Case 1
			HeaderID = DlgValue("CheckList")
			MainArray = ReSplit(CheckDataList(HeaderID),JoinStr)
			SetsArray = ReSplit(MainArray(1),SubJoinStr)
			SetsArray(0) = DlgText("ExCrBox")
			SetsArray(1) = DlgText("PreInsertSplitBox")
			SetsArray(2) = DlgText("ChkBktBox")
			SetsArray(3) = DlgText("KpPairBox")
			SetsArray(4) = DlgValue("AsiaKeyBox")
			SetsArray(5) = DlgText("ChkEndBox")
			SetsArray(6) = DlgText("NoTrnEndBox")
			SetsArray(7) = DlgText("AutoTrnEndBox")
			SetsArray(8) = DlgText("ShortBox")
			SetsArray(9) = DlgText("ShortKeyBox")
			SetsArray(10) = DlgText("KpShortKeyBox")
			SetsArray(11) = DlgText("PreRepStrBox")
			SetsArray(12) = DlgText("AutoWebFlagBox")
			SetsArray(13) = DlgText("AccKeyBox")
			SetsArray(14) = DlgValue("AddAcckeyBox")
			SetsArray(15) = DlgValue("LineSplitModeBox")
			SetsArray(16) = DlgText("AppInsertSplitBox")
			SetsArray(17) = DlgText("ReplaceSplitBox")
			MainArray(1) = Join(SetsArray,SubJoinStr)
			CheckDataList(HeaderID) = Join(MainArray,JoinStr)
		Case 2
			If getMsgList(UIDataList,MsgList,"SettingsDlgFunc",1) = False Then Exit Function
			tUpdateSet(0) = DlgValue("UpdateSet")
			tUpdateSet(1) = DlgText("WebSiteBox")
			tUpdateSet(2) = DlgText("CmdPathBox")
			tUpdateSet(3) = DlgText("ArgumentBox")
			tUpdateSet(4) = DlgText("UpdateCycleBox")
			tUpdateSet(5) = DlgText("UpdateDateBox")
			If tUpdateSet(5) = MsgList(4) Then tUpdateSet(5) = ""
		End Select
	End Select
End Function


'添加设置名称
Function AddSet(DataArr() As String,Optional ByVal sName As String) As String
	Dim Header As String,NewHeader As String
	Dim i As Long,MsgList() As String

	If getMsgList(UIDataList,MsgList,"AddSet",1) = False Then Exit Function
	Begin Dialog UserDialog 380,77,MsgList(0),.CommonDlgFunc ' %GRID:10,7,1,1
		Text 10,7,360,14,MsgList(1),.MainText
		TextBox 10,21,360,21,.TextBox
		OKButton 100,49,80,21,.OKButton
		CancelButton 200,49,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = sName
	DataInPutDlg:
	If Dialog(dlg) = 0 Then
		AddSet = ""
		Exit Function
	End If

	NewHeader = Trim(dlg.TextBox)
	AddSet = ""
	If NewHeader <> "" Then
		For i = LBound(DataArr) To UBound(DataArr)
			If LCase$(NewHeader) = LCase$(DataArr(i)) Then
				AddSet = DataArr(i)
				Exit For
			End If
		Next i
	End If

	If NewHeader = "" Then
		MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(2)
		GoTo DataInPutDlg
	ElseIf LCase$(NewHeader) = LCase$(AddSet) Then
		MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(2)
		GoTo DataInPutDlg
	End If
	AddSet = NewHeader
End Function


'编辑设置名称
Function EditSet(DataArr() As String,ByVal HeaderID As Long) As String
	Dim Header As String,NewHeader As String
	Dim i As Long,MsgList() As String

	If getMsgList(UIDataList,MsgList,"EditSet",1) = False Then Exit Function
	Header = DataArr(HeaderID)
	If InStr(Header,"&") Then
		Header = Replace$(Header,"&","&&")
	End If

	Begin Dialog UserDialog 380,126,MsgList(0),.CommonDlgFunc ' %GRID:10,7,1,1
		Text 10,7,350,14,MsgList(1),.NameText
		GroupBox 10,17,360,28,"",.GroupBox1
		Text 20,28,340,14,Header,.oldNameText
		Text 10,56,360,14,MsgList(2),.newNameText
		TextBox 10,70,360,21,.TextBox
		OKButton 100,98,80,21,.OKButton
		CancelButton 200,98,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	DataInPutDlg:
	dlg.TextBox = DataArr(HeaderID)
	If Dialog(dlg) = 0 Then
		EditSet = ""
		Exit Function
	End If

	NewHeader = Trim$(dlg.TextBox)
	EditSet = ""
	If NewHeader <> "" Then
		For i = LBound(DataArr) To UBound(DataArr)
			If LCase$(NewHeader) = LCase$(DataArr(i)) Then
				EditSet = DataArr(i)
				Exit For
			End If
		Next i
	End If

	If NewHeader = "" Then
		MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(3)
		GoTo DataInPutDlg
	ElseIf LCase$(NewHeader) = LCase$(EditSet) Then
		MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(3)
		GoTo DataInPutDlg
	End If
	EditSet = NewHeader
End Function


'编辑设置名称和添加设置名称对话框函数
Private Function CommonDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	End Select
End Function


'添加或编辑语言对
Function SetLang(DataArr() As String,ByVal LangName As String,ByVal SrcCode As String) As String
	Dim i As Long,MsgList() As String,LangPairList() As String
	If getMsgList(UIDataList,MsgList,"SetLang",1) = False Then Exit Function
	Begin Dialog UserDialog 390,168,IIf(LangName = "",MsgList(0),MsgList(1)),.LangDlgFunc ' %GRID:10,7,1,1
		Text 10,14,370,14,MsgList(2),.LangText
		TextBox 10,35,370,21,.LangName
		Text 10,77,370,14,MsgList(3),.CodeText
		TextBox 10,98,370,21,.SrcCode
		OKButton 90,140,80,21,.OKButton
		CancelButton 220,140,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    dlg.LangName = LangName
    dlg.SrcCode = SrcCode
    StartDlg:
    If Dialog(dlg) = 0 Then Exit Function
	If dlg.LangName & dlg.SrcCode = "" Then
		MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(4)
		GoTo StartDlg
	End If
	If dlg.LangName = "" Or dlg.SrcCode = "" Then
		MsgBox MsgList(6),vbOkOnly+vbInformation,MsgList(4)
		GoTo StartDlg
	End If
	If dlg.LangName = LangName And dlg.SrcCode = SrcCode Then Exit Function
	For i = LBound(DataArr) To UBound(DataArr)
		LangPairList = ReSplit(LCase$(DataArr(i)),LngJoinStr)
		If LCase$(dlg.LangName) = LangPairList(0) Then
			MsgBox MsgList(7),vbOkOnly+vbInformation,MsgList(4)
			GoTo StartDlg
		ElseIf LCase$(dlg.SrcCode) = LangPairList(1) Then
			MsgBox MsgList(8),vbOkOnly+vbInformation,MsgList(4)
			GoTo StartDlg
		End If
	Next i
	SetLang = dlg.LangName & LngJoinStr & dlg.SrcCode & LngJoinStr
End Function


'语言编辑对话框函数
Private Function LangDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,LangName As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		LangName = LCase$(DlgText("LangName"))
		If LangName <> "" Then
			For i = LBound(PslLangDataList) To UBound(PslLangDataList)
				If LCase$(ReSplit(PslLangDataList(i),LngJoinStr)(0)) = LangName Then
					DlgEnable "LangName",False
					Exit For
				End If
			Next i
		End If
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 3 ' 文本框或者组合框文本被更改
		Select Case DlgItem$
		Case "LangName"
			DlgText "LangName",Trim$(DlgText("LangName"))
		Case "SrcCode"
			DlgText "SrcCode",Trim$(DlgText("SrcCode"))
			If DlgText("SrcCode") = "" Then Exit Function
			LangName = PSL.GetLangCode(PSL.GetLangID(DlgText("SrcCode"),pslCode639_1),pslCodeText)
			If LangName = "3FF3F" Then
				LangName = PSL.GetLangCode(PSL.GetLangID(DlgText("SrcCode"),pslCodeLangRgn),pslCodeText)
			End If
			If LangName <> "3FF3F" Then DlgText "LangName",LangName
		Case "TarnCode"
			DlgText "TarnCode",Trim$(DlgText("TarnCode"))
		End Select
	End Select
End Function


'更改配置优先级
Function SetLevel(LevelList() As String,ByVal ID As Long,ByVal Msg As String) As Boolean
	Dim MsgList() As String
	If getMsgList(UIDataList,MsgList,"SetLevel",1) = False Then Exit Function
	UseStrList = LevelList
	AllStrList = LevelList
	Begin Dialog UserDialog 480,322,MsgList(0),.SetLevelDlgFunc ' %GRID:10,7,1,1
		Text 10,7,460,105,Msg,.MainText
		ListBox 10,119,360,161,LevelList(),.LevelList
		PushButton 380,119,90,21,MsgList(1),.UpButton
		PushButton 380,147,90,21,MsgList(2),.DownButton
		PushButton 380,175,90,21,MsgList(3),.ResetButton
		OKButton 130,294,80,21,.OKButton
		CancelButton 280,294,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.LevelList = ID
	If Dialog(dlg) = 0 Then Exit Function
	If ArrayComp(UseStrList,AllStrList) = True Then
		SetLevel = True
		LevelList = UseStrList
	End If
	Erase UseStrList,AllStrList
End Function


'更改配置优先级对话框函数
Private Function SetLevelDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,Temp As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		If UBound(UseStrList) = 0 Then
			DlgEnable "UpButton",False
			DlgEnable "DownButton",False
		Else
			Select Case DlgValue("LevelList")
			Case Is < 0
				DlgEnable "UpButton",False
				DlgEnable "DownButton",False
			Case 0
				DlgEnable "UpButton",False
			Case UBound(UseStrList)
				DlgEnable "DownButton",False
			Case Else
				DlgEnable "DownButton",True
			End Select
		End If
		DlgEnable "ResetButton",False
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		SetLevelDlgFunc = True '防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton","OKButton"
			SetLevelDlgFunc = False
			Exit Function
		Case "UpButton"
			i = DlgValue("LevelList")
			If i = 0 Then Exit Function
			Temp = UseStrList(i)
			UseStrList(i) = UseStrList(i - 1)
			UseStrList(i - 1) = Temp
			DlgListBoxArray "LevelList",UseStrList()
			DlgValue "LevelList",i - 1
		Case "DownButton"
			i = DlgValue("LevelList")
			If i = UBound(UseStrList) Then Exit Function
			Temp = UseStrList(i)
			UseStrList(i) = UseStrList(i + 1)
			UseStrList(i + 1) = Temp
			DlgListBoxArray "LevelList",UseStrList()
			DlgValue "LevelList",i + 1
		Case "ResetButton"
			UseStrList = AllStrList
			DlgListBoxArray "LevelList",UseStrList()
			DlgValue "LevelList",0
		End Select
		If UBound(UseStrList) = 0 Then
			DlgEnable "UpButton",False
			DlgEnable "DownButton",False
		Else
			Select Case DlgValue("LevelList")
			Case Is < 0
				DlgEnable "UpButton",False
				DlgEnable "DownButton",False
			Case 0
				DlgEnable "UpButton",False
				DlgEnable "DownButton",True
			Case UBound(UseStrList)
				DlgEnable "UpButton",True
				DlgEnable "DownButton",False
			Case Else
				DlgEnable "UpButton",True
				DlgEnable "DownButton",True
			End Select
		End If
		DlgEnable "ResetButton",ArrayComp(UseStrList,AllStrList)
	End Select
End Function


'添加或编辑语言对
Function EditLang(DataArr() As String,ByVal LangName As String,ByVal SrcCode As String,ByVal TarnCode As String) As String
	Dim i As Long,MsgList() As String,LangPairList() As String
	If getMsgList(UIDataList,MsgList,"EditLang",1) = False Then Exit Function
	Begin Dialog UserDialog 390,189,IIf(LangName = "",MsgList(0),MsgList(1)),.LangDlgFunc ' %GRID:10,7,1,1
		Text 10,7,370,14,MsgList(2),.LangText
		TextBox 10,28,370,21,.LangName
		Text 10,56,370,14,MsgList(3),.SrcCodeText
		TextBox 10,77,370,21,.SrcCode
		Text 10,105,370,14,MsgList(4),.TranCodeText
		TextBox 10,126,370,21,.TarnCode
		OKButton 90,161,80,21,.OKButton
		CancelButton 220,161,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    dlg.LangName = LangName
    dlg.SrcCode = SrcCode
    dlg.TarnCode = TarnCode
    StartDlg:
    If Dialog(dlg) = 0 Then Exit Function
	If dlg.LangName & dlg.SrcCode = "" Then
		MsgBox MsgList(6),vbOkOnly+vbInformation,MsgList(5)
		GoTo StartDlg
	End If
	If dlg.LangName = "" Or dlg.SrcCode = "" Then
		MsgBox MsgList(7),vbOkOnly+vbInformation,MsgList(5)
		GoTo StartDlg
	End If
	If dlg.LangName <> LangName Or dlg.SrcCode <> SrcCode Then
		For i = LBound(DataArr) To UBound(DataArr)
			LangPairList = ReSplit(LCase$(DataArr(i)),LngJoinStr)
			If LCase$(dlg.LangName) = LangPairList(0) Then
				MsgBox MsgList(8),vbOkOnly+vbInformation,MsgList(5)
				GoTo StartDlg
			ElseIf LCase$(dlg.SrcCode) = LangPairList(1) Then
				MsgBox MsgList(9),vbOkOnly+vbInformation,MsgList(5)
				GoTo StartDlg
			End If
		Next i
	ElseIf dlg.TarnCode = TarnCode Then
		Exit Function
	End If
	If dlg.TarnCode = "" Then
		If MsgBox(MsgList(10),vbYesNo+vbInformation,MsgList(5)) = vbYes Then GoTo StartDlg
		dlg.TarnCode = NullValue
	End If
	EditLang = dlg.LangName & LngJoinStr & dlg.SrcCode & LngJoinStr & dlg.TarnCode
End Function


'获取最近历史记录
Function GetHistory(DataList() As String,ByVal ItemName As String,ByVal ValueName As String,Optional ByVal Mode As Boolean) As Boolean
	Dim i As Long,n As Long,TempArray() As String
	On Error GoTo ExitFunction
	TempArray = GetAllSettings(AppName,ItemName)
	If Mode = True Then
		n = UBound(DataList) + 1
		ReDim Preserve DataList(n + UBound(TempArray)) As String
	Else
		ReDim DataList(UBound(TempArray)) As String
	End If
	For i = LBound(TempArray) To UBound(TempArray)
		If InStr(TempArray(i,0),ValueName) Then
			If TempArray(i,1) <> "" Then
				DataList(n) = TempArray(i,1)
				n = n + 1
			End If
		End If
	Next i
	If n > 0 Then
		n = n - 1
		GetHistory = True
	End If
	ReDim Preserve DataList(n) As String
	Exit Function
	ExitFunction:
	If Mode = False Then
		ReDim DataList(0) As String
	Else
		ReDim Preserve DataList(UBound(DataList)) As String
	End If
End Function


'保存最近历史记录
'Mode = False 写入最多 10 个记录，否则写入所有记录，并删除 DataList 中没有的记录
Function WriteHistory(DataList() As String,ByVal ItemName As String,ByVal ValueName As String,Optional ByVal Mode As Boolean) As Boolean
	Dim i As Long,n As Long,TempArray() As String
	On Error Resume Next
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i) <> "" Then
			SaveSetting(AppName,ItemName,ValueName & CStr$(n),DataList(i))
			n = n + 1
			If Mode = False Then
				If n = 10 Then
					WriteHistory = True
					Exit Function
				End If
			End If
		End If
	Next i
	If n > 0 Then WriteHistory = True
	TempArray = GetAllSettings(AppName,ItemName)
	For i = 0 To UBound(TempArray)
		If InStr(TempArray(i,0),ValueName & CStr$(n)) Then
			DeleteSetting(AppName,ItemName,TempArray(i,0))
			n = n + 1
		End If
	Next i
	If n < i Then Exit Function
	If WriteHistory = True Then Exit Function
	Dim WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.RegDelete "HKCU\Software\VB and VBA Program Settings\" & AppName & "\" & ItemName & "\"
	Set WshShell = Nothing
End Function


'获取设置
Function GetEngineSet(ByVal SelSet As String,ByVal Path As String) As Long
	Dim i As Long,j As Long,k As Long,m As Long,n As Long,x As Long,y As Long
	Dim Header As String,Temp As String,OldVersion As String
	Dim TempArray() As String,SetsArray() As String,DataList() As INIFILE_DATA

	ReDim SetsArray(19)
	If Path = EngineRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = EngineFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	If Path <> EngineFilePath Then
		ReDim FileDataList(0)
		FileDataList(0) = Path & JoinStr
		If EditFile(Path,FileDataList,False) = False Then
			GetEngineSet = 5
			Exit Function
		End If
		TempArray = ReSplit(FileDataList(0),JoinStr)
		Temp = TempArray(1)
	Else
		Temp = "_autodetect_all"
	End If
	k = 4
	ReDim Preserve Tools(k) As TOOLS_PROPERTIE
	If getINIFile(DataList,Path,Temp,1) = False Then Exit Function
	For i = 0 To UBound(DataList)
		With DataList(i)
			Select Case .Title
			Case "Option" 		'获取 Option 项和值
				For j = 0 To UBound(.Item)
					If .Item(j) = "Version" Then OldVersion = .Value(j)
					If SelSet = "" Or SelSet = "Option" Then
						Select Case .Item(j)
						Case "UILanguageID"
							tSelected(0) = .Value(j)
						Case "TranEngineSet"
							tSelected(1) = .Value(j)
						Case "CheckSet"
							tSelected(2) = .Value(j)
						Case "TranAllType"
							tSelected(3) = .Value(j)
						Case "TranMenu"
							tSelected(4) = .Value(j)
						Case "TranDialog"
							tSelected(5) = .Value(j)
						Case "TranString"
							tSelected(6) = .Value(j)
						Case "TranAcceleratorTable"
							tSelected(7) = .Value(j)
						Case "TranVersion"
							tSelected(8) = .Value(j)
						Case "TranOther"
							tSelected(9) = .Value(j)
						Case "TranSeletedOnly"
							tSelected(10) = .Value(j)
						Case "SkipForReview"
							tSelected(11) = .Value(j)
						Case "SkipValidated"
							tSelected(12) = .Value(j)
						Case "SkipNotTran"
							tSelected(13) = .Value(j)
						Case "SkipAllNumAndSymbol"
							tSelected(14) = .Value(j)
						Case "SkipAllUCase"
							tSelected(15) = .Value(j)
						Case "SkipAllLCase"
							tSelected(16) = .Value(j)
						Case "AutoSelection"
							tSelected(17) = .Value(j)
						Case "CheckSrcProject"
							tSelected(18) = .Value(j)
						Case "CheckSrcString"
							tSelected(19) = .Value(j)
						Case "ReplaceSrcString"
							tSelected(20) = .Value(j)
						Case "SplitTranslate"
							tSelected(21) = .Value(j)
						Case "CheckTrnProject"
							tSelected(22) = .Value(j)
						Case "CheckTrnString"
							tSelected(23) = .Value(j)
						Case "ReplaceTrnString"
							tSelected(24) = .Value(j)
						Case "KeepSetting"
							tSelected(25) = .Value(j)
						Case "ShowMassage"
							tSelected(26) = .Value(j)
						Case "AddTranComment"
							tSelected(27) = .Value(j)
						End Select
					End If
				Next j
				If SelSet = "" Or SelSet = "Option" Then
					If SelSet = "Option" Then
						If CheckArray(tSelected) = True Then GetEngineSet = 1
						Exit For
					End If
				End If
			Case "Update" 		'获取 Update 项和值
				If SelSet = "" Or SelSet = "Update" Then
					For j = 0 To UBound(.Item)
						Select Case .Item(j)
						Case "UpdateMode"
							tUpdateSet(0) = .Value(j)
						Case "Path"
							tUpdateSet(2) = .Value(j)
						Case "Argument"
							tUpdateSet(3) = .Value(j)
						Case "UpdateCycle"
							tUpdateSet(4) = .Value(j)
						Case "UpdateDate"
							tUpdateSet(5) = .Value(j)
						Case Else
							If InStr(.Item(j),"Site_") And .Value(j) <> "" Then
								If tUpdateSet(1) <> "" Then
									tUpdateSet(1) = tUpdateSet(1) & vbCrLf & .Value(j)
								Else
									tUpdateSet(1) = .Value(j)
								End If
							End If
						End Select
					Next j
					If SelSet = "Update" Then
						If CheckArray(tUpdateSet) = True Then GetEngineSet = 2
						Exit For
					End If
				End If
			Case "Tools" 		'获取 Tools 项和值
				If SelSet = "" Or SelSet = "Tools" Then
					For j = 0 To UBound(.Item)
						If .Item(j) Like "Tools[0-9]_Name" Then
							Tools(k).sName = .Value(j)
							y = y + 1
						ElseIf .Item(j) Like "Tools[0-9]_Path" Then
							Tools(k).FilePath = .Value(j)
							y = y + 1
						ElseIf .Item(j) Like "Tools[0-9]_Argument" Then
							Tools(k).Argument = .Value(j)
							y = y + 1
						End If
						If y = 3 Then
							If Tools(k).sName <> "" And Tools(k).FilePath <> "" Then
								k = k + 1
								ReDim Preserve Tools(k) As TOOLS_PROPERTIE
							End If
							y = 0
						End If
					Next j
					ReDim Preserve Tools(k - 1) As TOOLS_PROPERTIE
					If SelSet = "Tools" Then
						If k > 4 Then GetEngineSet = 3
						Exit For
					End If
				End If
			Case Else
				If SelSet = "" Or SelSet = "Sets" Or SelSet = .Title Then '获取 Option 项外的全部项和值
					For j = 0 To UBound(.Item)
						Select Case .Item(j)
						Case "ObjectName"
							SetsArray(0) = .Value(j)
						Case "AppId"
							SetsArray(1) = .Value(j)
						Case "EngineUrl"
							SetsArray(2) = .Value(j)
						Case "UrlTemplate"
							SetsArray(3) = .Value(j)
						Case "Method"
							SetsArray(4) = .Value(j)
						Case "Async"
							SetsArray(5) = .Value(j)
						Case "User"
							SetsArray(6) = .Value(j)
						Case "Password"
							SetsArray(7) = .Value(j)
						Case "SendBody"
							SetsArray(8) = .Value(j)
						Case "RequestHeader"
							SetsArray(9) = Convert(.Value(j))
						Case "ResponseType"
							SetsArray(10) = .Value(j)
						Case "TranBeforeStrByText"
							SetsArray(11) = .Value(j)
						Case "TranAfterStrByText"
							SetsArray(12) = .Value(j)
						Case "TranBeforeStrByBody"
							SetsArray(13) = .Value(j)
						Case "TranAfterStrByBody"
							SetsArray(14) = .Value(j)
						Case "TranBeforeStrByStream"
							SetsArray(15) = .Value(j)
						Case "TranAfterStrByStream"
							SetsArray(16) = .Value(j)
						Case "TranXMLIdName"
							SetsArray(17) = .Value(j)
						Case "TranXMLTagName"
							SetsArray(18) = .Value(j)
						Case "Enable"
							SetsArray(19) = .Value(j)
						Case "LangCodePair"
							LngPair = .Value(j)
						Case "TranBeforeStr"
							bStr = .Value(j)
						Case "TranAfterStr"
							aStr = .Value(j)
						End Select
					Next j
					If SetsArray(10) = "responseXML" Then
						If SetsArray(17) = "" Then SetsArray(17) = bStr
						If SetsArray(18) = "" Then SetsArray(18) = aStr
					Else
						If SetsArray(11) = "" Then SetsArray(11) = bStr
						If SetsArray(12) = "" Then SetsArray(12) = aStr
						If SetsArray(13) = "" Then SetsArray(13) = bStr
						If SetsArray(14) = "" Then SetsArray(14) = aStr
						If SetsArray(15) = "" Then SetsArray(15) = bStr
						If SetsArray(16) = "" Then SetsArray(16) = aStr
					End If
					If CheckNullData("",SetsArray,"1,6-9,15-19",6) = False Then
						If LngPair <> "" Then
							LngPair = Join(MergeLngList(LangCodeList(.Title,0,-1), _
										ReSplit(LngPair,SubLngJoinStr),"engine"),SubLngJoinStr)
						Else
							LngPair = Join(LangCodeList(.Title,0,-1),SubLngJoinStr)
						End If
						Temp = .Title & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
						'更新旧版的默认配置值
						If StrComp(ToUpdateEngineVersion,OldVersion) = 1 Then
							Temp = EngineDataUpdate(.Title,Temp)
						End If
						'保存数据到数组中
						CreateArray(.Title,Temp,EngineList,EngineDataList)
						x = x + 1
					End If
					'数据初始化
					ReDim SetsArray(19)
					LngPair = ""
					bStr = ""
					aStr = ""
					n = n + 1
				End If
			End Select
		End With
	Next i
	If n > 0 And x = n Then GetEngineSet = 4
	If Path = EngineFilePath Then
		If GetEngineSet = 0 Then GoTo GetFromRegistry
		'保存更新和导入后的数据到文件
		If GetEngineSet = 4 Then
			If OldVersion <> "" Then
				If StrComp(ToUpdateEngineVersion,OldVersion) = 1 Then
					WriteEngineSet(EngineFilePath,"All")
				End If
			End If
		End If
		If tWriteLoc = "" Then tWriteLoc = EngineFilePath
		GoTo GetFontSetFromRegistry
	End If
	Exit Function

	GetFromRegistry:
	If tWriteLoc = "" Then tWriteLoc = EngineRegKey
	'获取 Option 项和值
	OldVersion = GetSetting("WebTranslate","Option","Version","")
	If SelSet = "" Or SelSet = "Option" Then
		tSelected(0) = GetSetting("WebTranslate","Option","UILanguageID",0)
		tSelected(1) = GetSetting("WebTranslate","Option","TranEngineSet","")
		tSelected(2) = GetSetting("WebTranslate","Option","CheckSet","")
		tSelected(3) = GetSetting("WebTranslate","Option","TranAllType",0)
		tSelected(4) = GetSetting("WebTranslate","Option","TranMenu",0)
		tSelected(5) = GetSetting("WebTranslate","Option","TranDialog",0)
		tSelected(6) = GetSetting("WebTranslate","Option","TranString",0)
		tSelected(7) = GetSetting("WebTranslate","Option","TranAcceleratorTable",0)
		tSelected(8) = GetSetting("WebTranslate","Option","TranVersion",0)
		tSelected(9) = GetSetting("WebTranslate","Option","TranOther",0)
		tSelected(10) = GetSetting("WebTranslate","Option","TranSeletedOnly",1)
		tSelected(11) = GetSetting("WebTranslate","Option","SkipForReview",1)
		tSelected(12) = GetSetting("WebTranslate","Option","SkipValidated",1)
		tSelected(13) = GetSetting("WebTranslate","Option","SkipNotTran",0)
		tSelected(14) = GetSetting("WebTranslate","Option","SkipAllNumAndSymbol",1)
		tSelected(15) = GetSetting("WebTranslate","Option","SkipAllUCase",1)
		tSelected(16) = GetSetting("WebTranslate","Option","SkipAllLCase",0)
		tSelected(17) = GetSetting("WebTranslate","Option","AutoSelection",1)
		tSelected(18) = GetSetting("WebTranslate","Option","CheckSrcProject",4)
		tSelected(19) = GetSetting("WebTranslate","Option","CheckSrcString",1)
		tSelected(20) = GetSetting("WebTranslate","Option","ReplaceSrcString",1)
		tSelected(21) = GetSetting("WebTranslate","Option","SplitTranslate",1)
		tSelected(22) = GetSetting("WebTranslate","Option","CheckTrnProject",1)
		tSelected(23) = GetSetting("WebTranslate","Option","CheckTrnString",1)
		tSelected(24) = GetSetting("WebTranslate","Option","ReplaceTrnString",1)
		tSelected(25) = GetSetting("WebTranslate","Option","KeepSetting",1)
		tSelected(26) = GetSetting("WebTranslate","Option","ShowMassage",1)
		tSelected(27) = GetSetting("WebTranslate","Option","AddTranComment",0)
		If SelSet = "Option" Then
			If CheckArray(tSelected) = True Then GetEngineSet = 1
			Exit Function
		End If
	End If
	'获取 Update 项和值
	If SelSet = "" Or SelSet = "Update" Then
		tUpdateSet(0) = GetSetting("WebTranslate","Update","UpdateMode",1)
		Temp = GetSetting("WebTranslate","Update","Count","")
		If Temp <> "" Then
			For i = 0 To StrToLong(Temp)
				Temp = GetSetting("WebTranslate","Update",CStr(i),"")
				If Temp <> "" Then
					If tUpdateSet(1) <> "" Then
						tUpdateSet(1) = tUpdateSet(1) & vbCrLf & Temp
					Else
						tUpdateSet(1) = Temp
					End If
				End If
			Next i
		End If
		tUpdateSet(2) = GetSetting("WebTranslate","Update","Path","")
		tUpdateSet(3) = GetSetting("WebTranslate","Update","Argument","")
		tUpdateSet(4) = GetSetting("WebTranslate","Update","UpdateCycle",7)
		tUpdateSet(5) = GetSetting("WebTranslate","Update","UpdateDate","")
		If SelSet = "Update" Then
			If CheckArray(tUpdateSet) = True Then GetEngineSet = 2
			Exit Function
		End If
	End If
	'获取自定义工具
	If SelSet = "" Or SelSet = "Tools" Then
		Temp = GetSetting("WebTranslate","Tools","Count","")
		If Temp <> "" Then
			j = StrToLong(Temp): x = 4
			ReDim Tools(j + x) As TOOLS_PROPERTIE
			For i = 0 To j
				Tools(x).sName = GetSetting("WebTranslate","Tools",i & "_Name","")
				Tools(x).FilePath = GetSetting("WebTranslate","Tools",i & "_Path","")
				Tools(x).Argument = GetSetting("WebTranslate","Tools",i & "_Argument","")
				If Tools(x).sName <> "" And Tools(x).FilePath <> "" Then
					x = x + 1
				End If
			Next i
			If x > 4 Then x = x - 1
			ReDim Preserve Tools(x) As TOOLS_PROPERTIE
			If x - 4 = j Then GetEngineSet = 3
		End If
		If SelSet = "Tools" Then Exit Function
	End If
	'获取 Option 外的项和值
	If SelSet = "" Or SelSet = "Sets" Then
		Header = GetSetting("WebTranslate","Option","Headers","")
		If Header <> "" Then
			ReDim SetsArray(19)
			TempArray = ReSplit(Header,";",-1)
			n = UBound(TempArray): x = 0
			For i = 0 To n
				If TempArray(i) <> "" Then
					'转存旧版的每个项和值
					Header = GetSetting("WebTranslate",TempArray(i),"Name","")
				End If
				If Header = "" Then Header = TempArray(i)
				SetsArray(0) = GetSetting("WebTranslate",TempArray(i),"ObjectName","")
				SetsArray(1) = GetSetting("WebTranslate",TempArray(i),"AppId","")
				SetsArray(2) = GetSetting("WebTranslate",TempArray(i),"EngineUrl","")
				SetsArray(3) = GetSetting("WebTranslate",TempArray(i),"UrlTemplate","")
				SetsArray(4) = GetSetting("WebTranslate",TempArray(i),"Method","")
				SetsArray(5) = GetSetting("WebTranslate",TempArray(i),"Async","")
				SetsArray(6) = GetSetting("WebTranslate",TempArray(i),"User","")
				SetsArray(7) = GetSetting("WebTranslate",TempArray(i),"Password","")
				SetsArray(8) = GetSetting("WebTranslate",TempArray(i),"SendBody","")
				SetsArray(9) = Convert(GetSetting("WebTranslate",TempArray(i),"RequestHeader",""))
				SetsArray(10) = GetSetting("WebTranslate",TempArray(i),"ResponseType","")
				SetsArray(11) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStrByText","")
				SetsArray(12) = GetSetting("WebTranslate",TempArray(i),"TranAfterStrByText","")
				SetsArray(13) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStrByBody","")
				SetsArray(14) = GetSetting("WebTranslate",TempArray(i),"TranAfterStrByBody","")
				SetsArray(15) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStrByStream","")
				SetsArray(16) = GetSetting("WebTranslate",TempArray(i),"TranAfterStrByStream","")
				SetsArray(17) = GetSetting("WebTranslate",TempArray(i),"TranXMLIdName","")
				SetsArray(18) = GetSetting("WebTranslate",TempArray(i),"TranXMLTagName","")
				SetsArray(19) = GetSetting("WebTranslate",TempArray(i),"Enable","1")
				If SetsArray(10) = "responseXML" Then
					If SetsArray(17) = "" Then SetsArray(17) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStr","")
					If SetsArray(18) = "" Then SetsArray(18) = GetSetting("WebTranslate",TempArray(i),"TranAfterStr","")
				Else
					If SetsArray(11) = "" Then SetsArray(11) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStr","")
					If SetsArray(12) = "" Then SetsArray(12) = GetSetting("WebTranslate",TempArray(i),"TranAfterStr","")
					If SetsArray(13) = "" Then SetsArray(13) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStr","")
					If SetsArray(14) = "" Then SetsArray(14) = GetSetting("WebTranslate",TempArray(i),"TranAfterStr","")
					If SetsArray(15) = "" Then SetsArray(15) = GetSetting("WebTranslate",TempArray(i),"TranBeforeStr","")
					If SetsArray(16) = "" Then SetsArray(16) = GetSetting("WebTranslate",TempArray(i),"TranAfterStr","")
				End If
				LngPair = GetSetting("WebTranslate",TempArray(i),"LangCodePair","")
				If CheckNullData("",SetsArray,"1,6-9,15-19",6) = False Then
					If LngPair <> "" Then
						LngPair = Join(MergeLngList(LangCodeList(Header,0,-1), _
									ReSplit(LngPair,SubLngJoinStr),"engine"),SubLngJoinStr)
					Else
						LngPair = Join(LangCodeList(Header,0,-1),SubLngJoinStr)
					End If
					Temp = Header & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
					'更新旧版的默认配置值
					If StrComp(ToUpdateEngineVersion,OldVersion) = 1 Then
						Temp = EngineDataUpdate(Header,Temp)
					End If
					'保存数据到数组中
					CreateArray(Header,Temp,EngineList,EngineDataList)
					x = x + 1
				End If
				'删除旧版配置值
				On Error Resume Next
				If Header = TempArray(i) Then DeleteSetting("WebTranslate",Header)
				On Error GoTo 0
			Next i
			If x = n + 1 Then GetEngineSet = 4
			'保存更新后的数据到注册表
			If GetEngineSet = 4 Then
				If StrComp(ToUpdateEngineVersion,OldVersion) = 1 Then
					WriteEngineSet(EngineRegKey,"Sets")
				End If
			End If
		End If
	End If

	'获取对话框字体设置
	GetFontSetFromRegistry:
	If SelSet = "" Or InStr(";DlgFont;MainFont;SrcStrFont;TrnStrFont;",";" & SelSet & ";") Then
		x = 0
		On Error GoTo ExitFunction
		Dim TempByte() As Byte
		Temp = IIf(SelSet = "DlgFont","_",SelSet & "_")
		TempArray = GetAllSettings(AppName,"DlgFonts")
		For i = LBound(TempArray) To UBound(TempArray)
			Select Case ReSplit(TempArray(i,0),"_",2)(0)
			Case "MainFont"
				j = 0
			Case "SrcStrFont"
				j = 1
			Case "TrnStrFont"
				j = 2
			End Select
			If InStr(TempArray(i,0),Temp) Then
				With LFList(j)
					Select Case ReSplit(TempArray(i,0),"_",2)(1)
					Case "lfCharSet"
						.lfCharSet = StrToLong(TempArray(i,1))
					Case "lfClipPrecision"
						.lfClipPrecision = StrToLong(TempArray(i,1))
					Case "lfEscapement"
						.lfEscapement = StrToLong(TempArray(i,1))
					Case "lfFaceName"
						TempByte = StrConv$(TempArray(i,1),vbFromUnicode)
						ReDim Preserve TempByte(31) As Byte
						CopyMemory .lfFaceName(0),TempByte(0),UBound(.lfFaceName) + 1
					Case "lfHeight"
						.lfHeight = StrToLong(TempArray(i,1))
					Case "lfItalic"
						.lfItalic = StrToLong(TempArray(i,1))
					Case "lfOrientation"
						.lfOrientation = StrToLong(TempArray(i,1))
					Case "lfOutPrecision"
						.lfOutPrecision = StrToLong(TempArray(i,1))
					Case "lfPitchAndFamily"
						.lfPitchAndFamily = StrToLong(TempArray(i,1))
					Case "lfQuality"
						.lfQuality = StrToLong(TempArray(i,1))
					Case "lfStrikeOut"
						.lfStrikeOut = StrToLong(TempArray(i,1))
					Case "lfUnderline"
						.lfUnderline = StrToLong(TempArray(i,1))
					Case "lfWeight"
						.lfWeight = StrToLong(TempArray(i,1))
					Case "lfWidth"
						.lfWidth = StrToLong(TempArray(i,1))
					Case "lfColor"
						.lfColor = StrToLong(TempArray(i,1))
					End Select
					x = x + 1
				End With
			End If
		Next i
		On Error GoTo 0
		ExitFunction:
		If SelSet <> "" Then
			If x > 0 Then GetEngineSet = 5
		End If
	End If
End Function


'写入翻译引擎设置
Function WriteEngineSet(ByVal Path As String,ByVal WriteType As String) As Boolean
	Dim i As Long,j As Long,n As Long,Temp As String,KeepSet As Long
	Dim SetsArray() As String,LineArray() As String
	Dim Selected() As String,TempArray() As String

	KeepSet = StrToLong(tSelected(25))
	If KeepSet = 0 Then
		If WriteType = "Main" Or WriteType = "All" Then KeepSet = 1
		If Path = EngineFilePath Then KeepSet = 1
	End If
	Selected = tSelected
	Select Case Selected(2)
	Case DefaultCheckList(0)
		Selected(2) = "en2zh"
	Case DefaultCheckList(1)
		Selected(2) = "zh2en"
	End Select

	'写入文件
	If Path <> "" And Path <> EngineRegKey Then
   		Temp = Left(Path,InStrRev(Path,"\"))
   		On Error Resume Next
   		If Dir(Temp & "*.*") = "" Then MkDir Temp
		If Dir(Path) <> "" Then SetAttr Path,vbNormal
		On Error GoTo ExitFunction
		ReDim LineArray(6)
			LineArray(0) = ";------------------------------------------------------------"
			LineArray(1) = ";Settings for PSLWebTrans.bas"
			LineArray(2) = ";------------------------------------------------------------"
			LineArray(4) = "[Option]"
			LineArray(5) = "AppName = " & AppName
			LineArray(6) = "Version = " & Version
			n = UBound(LineArray) + 1
			ReDim Preserve LineArray(n + 0)
			LineArray(n + 0) = "UILanguageID = " & Selected(0)
			If KeepSet = 1 Then
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 26)
				LineArray(n + 0) = "TranEngineSet = " & Selected(1)
				LineArray(n + 1) = "CheckSet = " & Selected(2)
				LineArray(n + 2) = "TranAllType = " & Selected(3)
				LineArray(n + 3) = "TranMenu = " & Selected(4)
				LineArray(n + 4) = "TranDialog = " & Selected(5)
				LineArray(n + 5) = "TranString = " & Selected(6)
				LineArray(n + 6) = "TranAcceleratorTable = " & Selected(7)
				LineArray(n + 7) = "TranVersion = " & Selected(8)
				LineArray(n + 8) = "TranOther = " & Selected(9)
				LineArray(n + 9) = "TranSeletedOnly = " & Selected(10)
				LineArray(n + 10) = "SkipForReview = " & Selected(11)
				LineArray(n + 11) = "SkipValidated = " & Selected(12)
				LineArray(n + 12) = "SkipNotTran = " & Selected(13)
				LineArray(n + 13) = "SkipAllNumAndSymbol = " & Selected(14)
				LineArray(n + 14) = "SkipAllUCase = " & Selected(15)
				LineArray(n + 15) = "SkipAllLCase = " & Selected(16)
				LineArray(n + 16) = "AutoSelection = " & Selected(17)
				LineArray(n + 17) = "CheckSrcProject = " & Selected(18)
				LineArray(n + 18) = "CheckSrcString = " & Selected(19)
				LineArray(n + 19) = "ReplaceSrcString = " & Selected(20)
				LineArray(n + 20) = "SplitTranslate = " & Selected(21)
				LineArray(n + 21) = "CheckTrnProject = " & Selected(22)
				LineArray(n + 22) = "CheckTrnString = " & Selected(23)
				LineArray(n + 23) = "ReplaceTrnString = " & Selected(24)
				LineArray(n + 24) = "KeepSetting = " & Selected(25)
				LineArray(n + 25) = "ShowMassage = " & Selected(26)
				LineArray(n + 26) = "AddTranComment = " & Selected(27)
			End If
			n = UBound(LineArray) + 1
			ReDim Preserve LineArray(n + 0)
			If CheckArray(tUpdateSet) = True Then
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 1)
				LineArray(n + 0) = "[Update]"
				LineArray(n + 1) = "UpdateMode = " & tUpdateSet(0)
				TempArray = ReSplit(tUpdateSet(1),vbCrLf,-1)
				j = 0
				For i = 0 To UBound(TempArray)
					If Trim$(TempArray(i)) <> "" Then
						n = UBound(LineArray) + 1
						ReDim Preserve LineArray(n)
						LineArray(n) = "Site_" & j & " = " & TempArray(i)
						j = j + 1
					End If
				Next i
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 4)
				LineArray(n + 0) = "Path = " & tUpdateSet(2)
				LineArray(n + 1) = "Argument = " & tUpdateSet(3)
				LineArray(n + 2) = "UpdateCycle = " & tUpdateSet(4)
				LineArray(n + 3) = "UpdateDate = " & tUpdateSet(5)
			End If
			If UBound(Tools) > 3 Then
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n)
				LineArray(n) = "[Tools]"
				For i = 4 To UBound(Tools)
					If Tools(i).sName <> "" And Tools(i).FilePath <> "" Then
						n = UBound(LineArray) + 1
						ReDim Preserve LineArray(n + 2)
						LineArray(n + 0) = "Tools" & i - 4 & "_Name = " & Tools(i).sName
						LineArray(n + 1) = "Tools" & i - 4 & "_Path = " & Tools(i).FilePath
						LineArray(n + 2) = "Tools" & i - 4 & "_Argument = " & Tools(i).Argument
					End If
				Next i
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n)
			End If
			For i = LBound(EngineDataList) To UBound(EngineDataList)
				TempArray = ReSplit(EngineDataList(i),JoinStr)
				SetsArray = ReSplit(TempArray(1),SubJoinStr)
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 21)
				LineArray(n + 0) = "[" & TempArray(0) & "]"
				LineArray(n + 1) = "ObjectName = " & SetsArray(0)
				LineArray(n + 2) = "AppId = " & SetsArray(1)
				LineArray(n + 3) = "EngineUrl = " & SetsArray(2)
				LineArray(n + 4) = "UrlTemplate = " & SetsArray(3)
				LineArray(n + 5) = "Method = " & SetsArray(4)
				LineArray(n + 6) = "Async = " & SetsArray(5)
				LineArray(n + 7) = "User = " & SetsArray(6)
				LineArray(n + 8) = "Password = " & SetsArray(7)
				LineArray(n + 9) = "SendBody = " & SetsArray(8)
				LineArray(n + 10) = "RequestHeader = " & ReConvert(SetsArray(9))
				LineArray(n + 11) = "ResponseType = " & SetsArray(10)
				LineArray(n + 12) = "TranBeforeStrByText = " & SetsArray(11)
				LineArray(n + 13) = "TranAfterStrByText = " & SetsArray(12)
				LineArray(n + 14) = "TranBeforeStrByBody = " & SetsArray(13)
				LineArray(n + 15) = "TranAfterStrByBody = " & SetsArray(14)
				LineArray(n + 16) = "TranBeforeStrByStream = " & SetsArray(15)
				LineArray(n + 17) = "TranAfterStrByStream = " & SetsArray(16)
				LineArray(n + 18) = "TranXMLIdName = " & SetsArray(17)
				LineArray(n + 19) = "TranXMLTagName = " & SetsArray(18)
				LineArray(n + 20) = "Enable = " & SetsArray(19)
				LineArray(n + 21) = "LangCodePair = " & getLngPair(TempArray(2),"engine")
				If i <> UBound(EngineDataList) Then
					n = UBound(LineArray) + 1
					ReDim Preserve LineArray(n)
				End If
			Next i
		WriteToFile(Path,Join(LineArray,vbCrLf),"unicodeFFFE")
		On Error GoTo 0
		WriteEngineSet = True
		If Path = EngineFilePath Then
			tWriteLoc = EngineFilePath
			GoTo RemoveRegKey
		End If
	'写入注册表
	ElseIf Path = EngineRegKey Then
		On Error GoTo ExitFunction
		SaveSetting("WebTranslate","Option","Version",Version)
		If WriteType = "Main" Or WriteType = "Sets" Or WriteType = "All" Then
			If KeepSet = 1 Then
				SaveSetting("WebTranslate","Option","TranEngineSet",Selected(1))
				SaveSetting("WebTranslate","Option","CheckSet",Selected(2))
				SaveSetting("WebTranslate","Option","TranAllType",Selected(3))
				SaveSetting("WebTranslate","Option","TranMenu",Selected(4))
				SaveSetting("WebTranslate","Option","TranDialog",Selected(5))
				SaveSetting("WebTranslate","Option","TranString",Selected(6))
				SaveSetting("WebTranslate","Option","TranAcceleratorTable",Selected(7))
				SaveSetting("WebTranslate","Option","TranVersion",Selected(8))
				SaveSetting("WebTranslate","Option","TranOther",Selected(9))
				SaveSetting("WebTranslate","Option","TranSeletedOnly",Selected(10))
				SaveSetting("WebTranslate","Option","SkipForReview",Selected(11))
				SaveSetting("WebTranslate","Option","SkipValidated",Selected(12))
				SaveSetting("WebTranslate","Option","SkipNotTran",Selected(13))
				SaveSetting("WebTranslate","Option","SkipAllNumAndSymbol",Selected(14))
				SaveSetting("WebTranslate","Option","SkipAllUCase",Selected(15))
				SaveSetting("WebTranslate","Option","SkipAllLCase",Selected(16))
				SaveSetting("WebTranslate","Option","AutoSelection",Selected(17))
				SaveSetting("WebTranslate","Option","CheckSrcProject",Selected(18))
				SaveSetting("WebTranslate","Option","CheckSrcString",Selected(19))
				SaveSetting("WebTranslate","Option","ReplaceSrcString",Selected(20))
				SaveSetting("WebTranslate","Option","SplitTranslate",Selected(21))
				SaveSetting("WebTranslate","Option","CheckTrnProject",Selected(22))
				SaveSetting("WebTranslate","Option","CheckTrnString",Selected(23))
				SaveSetting("WebTranslate","Option","ReplaceTrnString",Selected(24))
				SaveSetting("WebTranslate","Option","KeepSetting",Selected(25))
				SaveSetting("WebTranslate","Option","ShowMassage",Selected(26))
				SaveSetting("WebTranslate","Option","AddTranComment",Selected(27))
			End If
		End If
		If WriteType = "Sets" Or WriteType = "All" Then
			'删除原配置项
			Temp = GetSetting("WebTranslate","Option","Headers")
			If Temp <> "" Then
				TempArray = ReSplit(Temp,";",-1)
				On Error Resume Next
				For i = 0 To UBound(TempArray)
					DeleteSetting("WebTranslate",TempArray(i))
				Next i
				On Error GoTo 0
			End If
			'写入新配置项
			LineArray = EngineDataList
			For i = LBound(EngineDataList) To UBound(EngineDataList)
				LineArray(i) = CStr(i)
				TempArray = ReSplit(EngineDataList(i),JoinStr)
				SetsArray = ReSplit(TempArray(1),SubJoinStr)
				SaveSetting("WebTranslate",LineArray(i),"Name",TempArray(0))
				SaveSetting("WebTranslate",LineArray(i),"ObjectName",SetsArray(0))
				SaveSetting("WebTranslate",LineArray(i),"AppId",SetsArray(1))
				SaveSetting("WebTranslate",LineArray(i),"EngineUrl",SetsArray(2))
				SaveSetting("WebTranslate",LineArray(i),"UrlTemplate",SetsArray(3))
				SaveSetting("WebTranslate",LineArray(i),"Method",SetsArray(4))
				SaveSetting("WebTranslate",LineArray(i),"Async",SetsArray(5))
				SaveSetting("WebTranslate",LineArray(i),"User",SetsArray(6))
				SaveSetting("WebTranslate",LineArray(i),"Password",SetsArray(7))
				SaveSetting("WebTranslate",LineArray(i),"SendBody",SetsArray(8))
				SaveSetting("WebTranslate",LineArray(i),"RequestHeader",ReConvert(SetsArray(9)))
				SaveSetting("WebTranslate",LineArray(i),"ResponseType",SetsArray(10))
				SaveSetting("WebTranslate",LineArray(i),"TranBeforeStrByText",SetsArray(11))
				SaveSetting("WebTranslate",LineArray(i),"TranAfterStrByText",SetsArray(12))
				SaveSetting("WebTranslate",LineArray(i),"TranBeforeStrByBody",SetsArray(13))
				SaveSetting("WebTranslate",LineArray(i),"TranAfterStrByBody",SetsArray(14))
				SaveSetting("WebTranslate",LineArray(i),"TranBeforeStrByStream",SetsArray(15))
				SaveSetting("WebTranslate",LineArray(i),"TranAfterStrByStream",SetsArray(16))
				SaveSetting("WebTranslate",LineArray(i),"TranXMLIdName",SetsArray(17))
				SaveSetting("WebTranslate",LineArray(i),"TranXMLTagName",SetsArray(18))
				SaveSetting("WebTranslate",LineArray(i),"Enable",SetsArray(19))
				SaveSetting("WebTranslate",LineArray(i),"LangCodePair",getLngPair(TempArray(2),"engine"))
			Next i
			SaveSetting("WebTranslate","Option","Headers",Join(LineArray,";"))
			SaveSetting("WebTranslate","Option","UILanguageID",Selected(0))
		End If
		If WriteType = "Update" Or WriteType = "Sets" Or WriteType = "All" Then
			If CheckArray(tUpdateSet) = True Then
				On Error Resume Next
				DeleteSetting("WebTranslate","Update")
				On Error GoTo 0
				TempArray = ReSplit(tUpdateSet(1),vbCrLf,-1)
				SaveSetting("WebTranslate","Update","UpdateMode",tUpdateSet(0))
				n = 0
				For i = 0 To UBound(TempArray)
					If Trim$(TempArray(i)) <> "" Then
						SaveSetting("WebTranslate","Update",CStr(n),TempArray(i))
						n = n + 1
					End If
				Next i
				If n > 0 Then SaveSetting("WebTranslate","Update","Count",n - 1)
				SaveSetting("WebTranslate","Update","Path",tUpdateSet(2))
				SaveSetting("WebTranslate","Update","Argument",tUpdateSet(3))
				SaveSetting("WebTranslate","Update","UpdateCycle",tUpdateSet(4))
				SaveSetting("WebTranslate","Update","UpdateDate",tUpdateSet(5))
			End If
		End If
		If WriteType = "Tools" Or WriteType = "Sets" Or WriteType = "All" Then
			On Error Resume Next
			DeleteSetting("WebTranslate","Tools")
			On Error GoTo 0
			If UBound(Tools) > 3 Then
				n = 0
				For i = 4 To UBound(Tools)
					If Tools(i).sName <> "" And Tools(i).FilePath <> "" Then
						SaveSetting("WebTranslate","Tools",n & "_Name",Tools(i).sName)
						SaveSetting("WebTranslate","Tools",n & "_Path",Tools(i).FilePath)
						SaveSetting("WebTranslate","Tools",n & "_Argument",Tools(i).Argument)
						n = n + 1
					End If
				Next i
				If n > 0 Then SaveSetting("WebTranslate","Tools","Count",n - 1)
			End If
		End If
		WriteEngineSet = True
		tWriteLoc = EngineRegKey
		GoTo RemoveFilePath
	'删除所有保存的设置
	ElseIf Path = "" Then
		'删除文件配置项
 		RemoveFilePath:
		On Error Resume Next
		If Dir(EngineFilePath) <> "" Then
			SetAttr EngineFilePath,vbNormal
			Kill EngineFilePath
		End If
		Temp = Left(EngineFilePath,InStrRev(EngineFilePath,"\"))
		If Dir(Temp & "*.*") = "" Then RmDir Temp
		On Error GoTo 0
		If Path = EngineRegKey Then GoTo ExitFunction
		'删除注册表配置项
		RemoveRegKey:
		If GetSetting("WebTranslate","Option","Version") <> "" Then
			Temp = GetSetting("WebTranslate","Option","Headers")
			On Error Resume Next
			If Temp <> "" Then
				TempArray = ReSplit(Temp,";",-1)
				For i = 0 To UBound(TempArray)
					DeleteSetting("WebTranslate",TempArray(i))
				Next i
			End If
			DeleteSetting("WebTranslate","Option")
			DeleteSetting("WebTranslate","Update")
			DeleteSetting("WebTranslate","Tools")
			Dim WshShell As Object
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.RegDelete EngineRegKey
			Set WshShell = Nothing
			On Error GoTo 0
		End If
		If Path = EngineFilePath Then GoTo ExitFunction
		'设置写入位置设置为空
		WriteEngineSet = True
		tWriteLoc = ""
	End If
	ExitFunction:
	'保存对话框字体设置
	If WriteType = "DlgFont" Or WriteType = "Sets" Then
		On Error Resume Next
		DeleteSetting(AppName,"DlgFonts")
		On Error GoTo 0
		For i = 0 To UBound(LFList)
			Select Case i
			Case 0
				Temp = "MainFont"
			Case 1
				Temp = "SrcStrFont"
			Case 2
				Temp = "TrnStrFont"
			End Select
			If CheckFont(LFList(i)) = True Then
				With LFList(i)
					SaveSetting(AppName,"DlgFonts",Temp & "_lfCharSet",CStr$(.lfCharSet))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfClipPrecision",CStr$(.lfClipPrecision))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfEscapement",CStr$(.lfEscapement))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfFaceName",ReSplit(StrConv$(.lfFaceName,vbUnicode),vbNullChar,2)(0))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfHeight",CStr$(.lfHeight))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfItalic",CStr$(.lfItalic))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfOrientation",CStr$(.lfOrientation))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfOutPrecision",CStr$(.lfOutPrecision))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfPitchAndFamily",CStr$(.lfPitchAndFamily))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfQuality",CStr$(.lfQuality))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfStrikeOut",CStr$(.lfStrikeOut))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfUnderline",CStr$(.lfUnderline))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfWeight",CStr$(.lfWeight))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfWidth",CStr$(.lfWidth))
					SaveSetting(AppName,"DlgFonts",Temp & "_lfColor",CStr$(.lfColor))
				End With
			Else
				ReDim tmpLFList(0) As LOG_FONT
				LFList(i) = tmpLFList(0)
			End If
		Next i
		WriteEngineSet = True
	End If
End Function


'获取字串检查设置
Function GetCheckSet(ByVal SelSet As String,ByVal Path As String) As Long
	Dim i As Long,j As Long,k As Long,m As Long,n As Long,x As Long
	Dim Header As String,Temp As String
	Dim TempArray() As String,SetsArray() As String,DataList() As INIFILE_DATA

	ReDim SetsArray(17)
	If SelSet = DefaultCheckList(0) Then
		SelSet = "en2zh"
	ElseIf SelSet = DefaultCheckList(1) Then
		SelSet = "zh2en"
	End If

	If Path = CheckRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = CheckFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	If Path = CheckFilePath Then On Error GoTo GetFromRegistry
	If Path <> CheckFilePath Then
		ReDim FileDataList(0)
		FileDataList(0) = Path & JoinStr
		If EditFile(Path,FileDataList,False) = False Then
			GetCheckSet = 5
			Exit Function
		End If
		TempArray = ReSplit(FileDataList(0),JoinStr)
		Temp = TempArray(1)
	Else
		Temp = "_autodetect_all"
	End If
	If getINIFile(DataList,Path,Temp,1) = False Then Exit Function
	For i = 0 To UBound(DataList)
		With DataList(i)
			Select Case .Title
			Case "Option" 		'获取 Option 项和值
				For j = 0 To UBound(.Item)
					If .Item(j) = "Version" Then CheckVersion = .Value(j)
					If SelSet = "" Or SelSet = "Option" Then
						Select Case .Item(j)
						Case "UILanguageID"
							cSelected(0) = .Value(j)
						Case "AutoMacroSet"
							cSelected(1) = .Value(j)
						Case "CheckMacroSet"
							cSelected(2) = .Value(j)
						Case "AutoMacroCheck"
							cSelected(3) = .Value(j)
						Case "AutoSelection"
							cSelected(4) = .Value(j)
						Case "SelectedCheck"
							cSelected(5) = .Value(j)
						Case "CheckAllType"
							cSelected(6) = .Value(j)
						Case "CheckMenu"
							cSelected(7) = .Value(j)
						Case "CheckDialog"
							cSelected(8) = .Value(j)
						Case "CheckString"
							cSelected(9) = .Value(j)
						Case "CheckAcceleratorTable"
							cSelected(10) = .Value(j)
						Case "CheckVersion"
							cSelected(11) = .Value(j)
						Case "CheckOther"
							cSelected(12) = .Value(j)
						Case "CheckSeletedOnly"
							cSelected(13) = .Value(j)
						Case "CheckAllCont"
							cSelected(14) = .Value(j)
						Case "CheckAccKey"
							cSelected(15) = .Value(j)
						Case "CheckEndChar"
							cSelected(16) = .Value(j)
						Case "CheckAcceler"
							cSelected(17) = .Value(j)
						Case "IgnoreVerTag"
							cSelected(18) = .Value(j)
						Case "IgnoreSetTag"
							cSelected(19) = .Value(j)
						Case "IgnoreDateTag"
							cSelected(20) = .Value(j)
						Case "IgnoreStateTag"
							cSelected(21) = .Value(j)
						Case "TgnoreAllTag"
							cSelected(22) = .Value(j)
						Case "NoCheckTag"
							cSelected(23) = .Value(j)
						Case "NoChangeTrnState"
							cSelected(24) = .Value(j)
						Case "AutoRepString"
							cSelected(25) = .Value(j)
						Case "KeepSettings"
							cSelected(26) = .Value(j)
						End Select
					End If
				Next j
				If SelSet = "" Or SelSet = "Option" Then
					Select Case cSelected(1)
					Case "Default", "en2zh"
						cSelected(1) = DefaultCheckList(0)
					Case "zh2en"
						cSelected(1) = DefaultCheckList(1)
					End Select
					Select Case cSelected(2)
					Case "Default", "en2zh"
						cSelected(2) = DefaultCheckList(0)
					Case "zh2en"
						cSelected(2) = DefaultCheckList(1)
					End Select
					If SelSet = "Option" Then
						If CheckArray(cSelected) = True Then GetCheckSet = 1
						Exit For
					End If
				End If
			Case "Update" 		'获取 Update 项和值
				If SelSet = "" Or SelSet = "Update" Then
					For j = 0 To UBound(.Item)
						Select Case .Item(j)
						Case "UpdateMode"
							cUpdateSet(0) = .Value(j)
						Case "Path"
							cUpdateSet(2) = .Value(j)
						Case "Argument"
							cUpdateSet(3) = .Value(j)
						Case "UpdateCycle"
							cUpdateSet(4) = .Value(j)
						Case "UpdateDate"
							cUpdateSet(5) = .Value(j)
						Case Else
							If InStr(.Item(j),"Site_") And .Value(j) <> "" Then
								If cUpdateSet(1) <> "" Then
									cUpdateSet(1) = cUpdateSet(1) & vbCrLf & .Value(j)
								Else
									cUpdateSet(1) = .Value(j)
								End If
							End If
						End Select
					Next j
					If SelSet = "Update" Then
						If CheckArray(cUpdateSet) = True Then GetCheckSet = 2
						Exit For
					End If
				End If
			Case "Projects" 		'获取 Project 项和值
				If SelSet = "" Or SelSet = "Project" Then
					For j = 0 To UBound(.Item)
						Select Case .Item(j)
						Case "CheckOnly"
							.Item(j) = DefaultProjectList(0)
						Case "CheckAndCorrect"
							.Item(j) = DefaultProjectList(1)
						Case "DelAccessKey"
							.Item(j) = DefaultProjectList(2)
						Case "DelAccelerator"
							.Item(j) = DefaultProjectList(3)
						Case "DelAccessKeyAndAccelerator"
							.Item(j) = DefaultProjectList(4)
						End Select
						If .Item(j) <> "" Then
							CreateArray(.Item(j),.Item(j) & JoinStr & .Value(j),ProjectList,ProjectDataList)
							k = k + 1
						End If
					Next j
					If SelSet = "Project" Then
						If k > 0 Then GetCheckSet = 3
						Exit For
					End If
				End If
			Case Else
				If SelSet = "" Or SelSet = "Sets" Or SelSet = .Title Then '获取 Option 项外的全部项和值
					For j = 0 To UBound(.Item)
						Select Case .Item(j)
						Case "ExcludeChar"
							SetsArray(0) = .Value(j)
						Case "LineSplitChar"
							SetsArray(1) = .Value(j)
						Case "CheckBracket"
							SetsArray(2) = .Value(j)
						Case "KeepCharPair"
							SetsArray(3) = .Value(j)
						Case "ShowAsiaKey"
							SetsArray(4) = .Value(j)
						Case "CheckEndChar"
							SetsArray(5) = .Value(j)
						Case "NoTrnEndChar"
							SetsArray(6) = .Value(j)
						Case "AutoTrnEndChar"
							SetsArray(7) = .Value(j)
						Case "CheckShortChar"
							SetsArray(8) = .Value(j)
						Case "CheckShortKey"
							SetsArray(9) = .Value(j)
						Case "KeepShortKey"
							SetsArray(10) = .Value(j)
						Case "PreRepString"
							SetsArray(11) = .Value(j)
						Case "AutoRepString"
							SetsArray(12) = .Value(j)
						Case "AccessKeyChar"
							SetsArray(13) = .Value(j)
						Case "AddAccessKeyWithFirstChar"
							SetsArray(14) = .Value(j)
						Case "LineSplitMode"
							SetsArray(15) = .Value(j)
						Case "AppInsertSplitChar"
							SetsArray(16) = .Value(j)
						Case "ReplaceSplitChar"
							SetsArray(17) = .Value(j)
						Case "ApplyLangList"
							LngPair = .Value(j)
						End Select
					Next j
					Temp = Join(SetsArray,"")
					If Temp <> "" And Temp <> "0" And Temp <> "1" Then
						Select Case .Title
						Case "en2zh"
							.Title = DefaultCheckList(0)
						Case "zh2en"
							.Title = DefaultCheckList(1)
						End Select
						If LngPair <> "" Then
							LngPair = Join(MergeLngList(LangCodeList(.Title,1,-1), _
										ReSplit(LngPair,SubLngJoinStr),"check"),SubLngJoinStr)
						Else
							LngPair = Join(LangCodeList(.Title,1,-1),SubLngJoinStr)
						End If
						Temp = .Title & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
						'更新旧版的默认配置值
						If StrComp(ToUpdateCheckVersion,CheckVersion) = 1 Then
							Temp = CheckDataUpdate(.Title,Temp)
						End If
						'保存数据到数组中
						CreateArray(.Title,Temp,CheckList,CheckDataList)
						x = x + 1
					End If
					'数据初始化
					ReDim SetsArray(17)
					LngPair = ""
					n = n + 1
				End If
			End Select
		End With
	Next i
	If n > 0 And x = n Then GetCheckSet = 4
	If Path = CheckFilePath Then
		If GetCheckSet = 0 Then GoTo GetFromRegistry
		'保存更新和导入后的数据到文件
		If GetCheckSet = 4 Then
			If CheckVersion <> "" Then
				If StrComp(ToUpdateCheckVersion,CheckVersion) = 1 Then
					CheckVersion = ToUpdateCheckVersion
					WriteCheckSet(CheckFilePath,"All")
				End If
			End If
		End If
		If cWriteLoc = "" Then cWriteLoc = CheckFilePath
	End If
	Exit Function

	GetFromRegistry:
	If cWriteLoc = "" Then cWriteLoc = CheckRegKey
	'获取 Option 项和值
	CheckVersion = GetSetting("AccessKey","Option","Version","")
	If SelSet = "" Or SelSet = "Option" Then
		cSelected(0) = GetSetting("AccessKey","Option","UILanguageID",0)
		cSelected(1) = GetSetting("AccessKey","Option","AutoMacroSet","")
		cSelected(2) = GetSetting("AccessKey","Option","CheckMacroSet","")
		cSelected(3) = GetSetting("AccessKey","Option","AutoMacroCheck",1)
		cSelected(4) = GetSetting("AccessKey","Option","AutoSelection",1)
		cSelected(5) = GetSetting("AccessKey","Option","SelectedCheck",0)
		cSelected(6) = GetSetting("AccessKey","Option","CheckAllType",0)
		cSelected(7) = GetSetting("AccessKey","Option","CheckMenu",0)
		cSelected(8) = GetSetting("AccessKey","Option","CheckDialog",0)
		cSelected(9) = GetSetting("AccessKey","Option","CheckString",0)
		cSelected(10) = GetSetting("AccessKey","Option","CheckAcceleratorTable",0)
		cSelected(11) = GetSetting("AccessKey","Option","CheckVersion",0)
		cSelected(12) = GetSetting("AccessKey","Option","CheckOther",0)
		cSelected(13) = GetSetting("AccessKey","Option","CheckSeletedOnly",1)
		cSelected(14) = GetSetting("AccessKey","Option","CheckAllCont",0)
		cSelected(15) = GetSetting("AccessKey","Option","CheckAccKey",0)
		cSelected(16) = GetSetting("AccessKey","Option","CheckEndChar",0)
		cSelected(17) = GetSetting("AccessKey","Option","CheckAcceler",0)
		cSelected(18) = GetSetting("AccessKey","Option","IgnoreVerTag",0)
		cSelected(19) = GetSetting("AccessKey","Option","IgnoreSetTag",0)
		cSelected(20) = GetSetting("AccessKey","Option","IgnoreDateTag",0)
		cSelected(21) = GetSetting("AccessKey","Option","IgnoreStateTag",0)
		cSelected(22) = GetSetting("AccessKey","Option","TgnoreAllTag",1)
		cSelected(23) = GetSetting("AccessKey","Option","NoCheckTag",1)
		cSelected(24) = GetSetting("AccessKey","Option","NoChangeTrnState",0)
		cSelected(25) = GetSetting("AccessKey","Option","AutoRepString",0)
		cSelected(26) = GetSetting("AccessKey","Option","KeepSetting",1)
		Select Case cSelected(1)
		Case "Default", "en2zh"
			cSelected(1) = DefaultCheckList(0)
		Case "zh2en"
			cSelected(1) = DefaultCheckList(1)
		End Select
		Select Case cSelected(2)
		Case "Default", "en2zh"
			cSelected(2) = DefaultCheckList(0)
		Case "zh2en"
			cSelected(2) = DefaultCheckList(1)
		End Select
		If SelSet = "Option" Then
			If CheckArray(cSelected) = True Then GetCheckSet = 1
			Exit Function
		End If
	End If
	'获取 Update 项和值
	If SelSet = "" Or SelSet = "Update" Then
		cUpdateSet(0) = GetSetting("AccessKey","Update","UpdateMode",1)
		Temp = GetSetting("AccessKey","Update","Count","")
		If Temp <> "" Then
			For i = 0 To StrToLong(Temp)
				Temp = GetSetting("AccessKey","Update",CStr(i),"")
				If Temp <> "" Then
					If cUpdateSet(1) <> "" Then
						cUpdateSet(1) = cUpdateSet(1) & vbCrLf & Temp
					Else
						cUpdateSet(1) = Temp
					End If
				End If
			Next i
		End If
		cUpdateSet(2) = GetSetting("AccessKey","Update","Path","")
		cUpdateSet(3) = GetSetting("AccessKey","Update","Argument","")
		cUpdateSet(4) = GetSetting("AccessKey","Update","UpdateCycle",7)
		cUpdateSet(5) = GetSetting("AccessKey","Update","UpdateDate","")
		If SelSet = "Update" Then
			If CheckArray(cUpdateSet) = True Then GetCheckSet = 2
			Exit Function
		End If
	End If
	'获取 Project 项和值
	If SelSet = "" Or SelSet = "Project" Then
		k = 0
		On Error GoTo NextItem
		TempArray = GetAllSettings("AccessKey","Projects")
		For i = LBound(TempArray) To UBound(TempArray)
			If TempArray(i,0) <> "" And TempArray(i,1) <> "" Then
				Select Case TempArray(i,0)
				Case "CheckOnly"
					TempArray(i,0) = DefaultProjectList(0)
				Case "CheckAndCorrect"
					TempArray(i,0) = DefaultProjectList(1)
				Case "DelAccessKey"
					TempArray(i,0) = DefaultProjectList(2)
				Case "DelAccelerator"
					TempArray(i,0) = DefaultProjectList(3)
				Case "DelAccessKeyAndAccelerator"
					TempArray(i,0) = DefaultProjectList(4)
				End Select
				CreateArray(TempArray(i,0),TempArray(i,0) & JoinStr & TempArray(i,1),ProjectList,ProjectDataList)
				k = k + 1
			End If
		Next i
		On Error GoTo 0
		NextItem:
		If SelSet = "Project" Then
			If k > 0 Then GetCheckSet = 3
			Exit Function
		End If
	End If
	'获取 Option 外的项和值
	If SelSet = "" Or SelSet = "Sets" Then
		Header = GetSetting("AccessKey","Option","Headers","")
		If Header <> "" Then
			ReDim SetsArray(17)
			TempArray = ReSplit(Header,";",-1)
			n = UBound(TempArray): x = 0
			For i = 0 To n
				If TempArray(i) <> "" Then
					'转存旧版的每个项和值
					Header = GetSetting("AccessKey",TempArray(i),"Name","")
				End If
				If Header = "" Then Header = TempArray(i)
				SetsArray(0) = GetSetting("AccessKey",TempArray(i),"ExcludeChar","")
				SetsArray(1) = GetSetting("AccessKey",TempArray(i),"LineSplitChar","")
				SetsArray(2) = GetSetting("AccessKey",TempArray(i),"CheckBracket","")
				SetsArray(3) = GetSetting("AccessKey",TempArray(i),"KeepCharPair","")
				SetsArray(4) = GetSetting("AccessKey",TempArray(i),"ShowAsiaKey","")
				SetsArray(5) = GetSetting("AccessKey",TempArray(i),"CheckEndChar","")
				SetsArray(6) = GetSetting("AccessKey",TempArray(i),"NoTrnEndChar","")
				SetsArray(7) = GetSetting("AccessKey",TempArray(i),"AutoTrnEndChar","")
				SetsArray(8) = GetSetting("AccessKey",TempArray(i),"CheckShortChar","")
				SetsArray(9) = GetSetting("AccessKey",TempArray(i),"CheckShortKey","")
				SetsArray(10) = GetSetting("AccessKey",TempArray(i),"KeepShortKey","")
				SetsArray(11) = GetSetting("AccessKey",TempArray(i),"PreRepString","")
				SetsArray(12) = GetSetting("AccessKey",TempArray(i),"AutoRepString","")
				SetsArray(13) = GetSetting("AccessKey",TempArray(i),"AccessKeyChar","")
				SetsArray(14) = GetSetting("AccessKey",TempArray(i),"AddAccessKeyWithFirstChar","")
				SetsArray(15) = GetSetting("AccessKey",TempArray(i),"LineSplitMode","")
				SetsArray(16) = GetSetting("AccessKey",TempArray(i),"AppInsertSplitChar","")
				SetsArray(17) = GetSetting("AccessKey",TempArray(i),"ReplaceSplitChar","")
				LngPair = GetSetting("AccessKey",TempArray(i),"ApplyLangList","")
				Temp = Join(SetsArray,"")
				If Temp <> "" And Temp <> "0" And Temp <> "1" Then
					Select Case Header
					Case "en2zh"
						Header = DefaultCheckList(0)
					Case "zh2en"
						Header = DefaultCheckList(1)
					End Select
					If LngPair <> "" Then
						LngPair = Join(MergeLngList(LangCodeList(Header,1,-1), _
									ReSplit(LngPair,SubLngJoinStr),"check"),SubLngJoinStr)
					Else
						LngPair = Join(LangCodeList(Header,1,-1),SubLngJoinStr)
					End If
					Temp = Header & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
					'更新旧版的默认配置值
					If StrComp(ToUpdateCheckVersion,CheckVersion) = 1 Then
						Temp = CheckDataUpdate(Header,Temp)
					End If
					'保存数据到数组中
					CreateArray(Header,Temp,CheckList,CheckDataList)
					x = x + 1
				End If
				'删除旧版配置值
				On Error Resume Next
				If Header = TempArray(i) Then DeleteSetting("AccessKey",Header)
				On Error GoTo 0
			Next i
			If x = n + 1 Then GetCheckSet = 4
			'保存更新后的数据到注册表
			If GetCheckSet = 4 Then
				If StrComp(ToUpdateCheckVersion,CheckVersion) = 1 Then
					CheckVersion = ToUpdateCheckVersion
					WriteCheckSet(CheckRegKey,"Sets")
				End If
			End If
		End If
	End If
End Function


'写入字串检查设置
Function WriteCheckSet(ByVal Path As String,ByVal WriteType As String) As Boolean
	Dim i As Long,j As Long,n As Long,Temp As String,KeepSet As Long
	Dim SetsArray() As String,LineArray() As String
	Dim Selected() As String,TempArray() As String

	KeepSet = StrToLong(cSelected(26))
	If KeepSet = 0 Then
		If WriteType = "Main" Or WriteType = "All" Then KeepSet = 1
		If Path = CheckFilePath Then KeepSet = 1
	End If
	Selected = cSelected
	Select Case Selected(1)
	Case DefaultCheckList(0)
		Selected(1) = "en2zh"
	Case DefaultCheckList(1)
		Selected(1) = "zh2en"
	End Select
	Select Case Selected(2)
	Case DefaultCheckList(0)
		Selected(2) = "en2zh"
	Case DefaultCheckList(1)
		Selected(2) = "zh2en"
	End Select

	'写入文件
	If Path <> "" And Path <> CheckRegKey Then
   		On Error Resume Next
   		Temp = Left(Path,InStrRev(Path,"\"))
   		If Dir(Temp & "*.*") = "" Then MkDir Temp
		If Dir(Path) <> "" Then SetAttr Path,vbNormal
		On Error GoTo ExitFunction
		ReDim LineArray(6)
			LineArray(0) = ";------------------------------------------------------------"
			LineArray(1) = ";Settings for PSLCheckAccessKeys.bas and PslAutoAccessKey.bas"
			LineArray(2) = ";------------------------------------------------------------"
			LineArray(4) = "[Option]"
			LineArray(5) = "AppName = PSLCheckAccessKeys"
			LineArray(6) = "Version = " & CheckVersion
			n = UBound(LineArray) + 1
			ReDim Preserve LineArray(n + 0)
			LineArray(n + 0) = "UILanguageID = " & Selected(0)
			If KeepSet = 1 Then
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 25)
				LineArray(n + 0) = "AutoMacroSet = " & Selected(1)
				LineArray(n + 1) = "CheckMacroSet = " & Selected(2)
				LineArray(n + 2) = "AutoMacroCheck = " & Selected(3)
				LineArray(n + 3) = "AutoSelection = " & Selected(4)
				LineArray(n + 4) = "SelectedCheck = " & Selected(5)
				LineArray(n + 5) = "CheckAllType = " & Selected(6)
				LineArray(n + 6) = "CheckMenu = " & Selected(7)
				LineArray(n + 7) = "CheckDialog = " & Selected(8)
				LineArray(n + 8) = "CheckString = " & Selected(9)
				LineArray(n + 9) = "CheckAcceleratorTable = " & Selected(10)
				LineArray(n + 10) = "CheckVersion = " & Selected(11)
				LineArray(n + 11) = "CheckOther = " & Selected(12)
				LineArray(n + 12) = "CheckSeletedOnly = " & Selected(13)
				LineArray(n + 13) = "CheckAllCont = " & Selected(14)
				LineArray(n + 14) = "CheckAccKey = " & Selected(15)
				LineArray(n + 15) = "CheckEndChar = " & Selected(16)
				LineArray(n + 16) = "CheckAcceler = " & Selected(17)
				LineArray(n + 17) = "IgnoreVerTag = " & Selected(18)
				LineArray(n + 18) = "IgnoreSetTag = " & Selected(19)
				LineArray(n + 19) = "IgnoreDateTag = " & Selected(20)
				LineArray(n + 20) = "IgnoreStateTag = " & Selected(21)
				LineArray(n + 21) = "TgnoreAllTag = " & Selected(22)
				LineArray(n + 22) = "NoCheckTag = " & Selected(23)
				LineArray(n + 23) = "NoChangeTrnState = " & Selected(24)
				LineArray(n + 24) = "AutoRepString = " & Selected(25)
				LineArray(n + 25) = "KeepSettings = " & Selected(26)
			End If
			n = UBound(LineArray) + 1
			ReDim Preserve LineArray(n + 0)
			If CheckArray(cUpdateSet) = True Then
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 1)
				LineArray(n + 0) = "[Update]"
				LineArray(n + 1) = "UpdateMode = " & cUpdateSet(0)
				TempArray = ReSplit(cUpdateSet(1),vbCrLf,-1)
				j = 0
				For i = 0 To UBound(TempArray)
					If Trim$(TempArray(i)) <> "" Then
						n = UBound(LineArray) + 1
						ReDim Preserve LineArray(n)
						LineArray(n) = "Site_" & j & " = " & TempArray(i)
						j = j + 1
					End If
				Next i
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 4)
				LineArray(n + 0) = "Path = " & cUpdateSet(2)
				LineArray(n + 1) = "Argument = " & cUpdateSet(3)
				LineArray(n + 2) = "UpdateCycle = " & cUpdateSet(4)
				LineArray(n + 3) = "UpdateDate = " & cUpdateSet(5)
			End If
			If CheckArray(ProjectList) = True Then
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n)
				LineArray(n) = "[Projects]"
				For i = LBound(ProjectDataList) To UBound(ProjectDataList)
					TempArray = ReSplit(ProjectDataList(i),JoinStr)
					Select Case TempArray(0)
					Case DefaultProjectList(0)
						TempArray(0) = "CheckOnly"
					Case DefaultProjectList(1)
						TempArray(0) = "CheckAndCorrect"
					Case DefaultProjectList(2)
						TempArray(0) = "DelAccessKey"
					Case DefaultProjectList(3)
						TempArray(0) = "DelAccelerator"
					Case DefaultProjectList(4)
						TempArray(0) = "DelAccessKeyAndAccelerator"
					End Select
					If TempArray(0) <> "" Then
						n = UBound(LineArray) + 1
						ReDim Preserve LineArray(n)
						LineArray(n) = TempArray(0) & " = " & TempArray(1)
					End If
				Next i
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n)
			End If
			For i = LBound(CheckDataList) To UBound(CheckDataList)
				TempArray = ReSplit(CheckDataList(i),JoinStr)
				SetsArray = ReSplit(TempArray(1),SubJoinStr)
				Select Case TempArray(0)
				Case DefaultCheckList(0)
					TempArray(0) = "en2zh"
				Case DefaultCheckList(1)
					TempArray(0) = "zh2en"
				End Select
				n = UBound(LineArray) + 1
				ReDim Preserve LineArray(n + 19)
				LineArray(n + 0) = "[" & TempArray(0) & "]"
				LineArray(n + 1) = "ExcludeChar = " & SetsArray(0)
				LineArray(n + 2) = "LineSplitChar = " & SetsArray(1)
				LineArray(n + 3) = "CheckBracket = " & SetsArray(2)
				LineArray(n + 4) = "KeepCharPair = " & SetsArray(3)
				LineArray(n + 5) = "ShowAsiaKey = " & SetsArray(4)
				LineArray(n + 6) = "CheckEndChar = " & SetsArray(5)
				LineArray(n + 7) = "NoTrnEndChar = " & SetsArray(6)
				LineArray(n + 8) = "AutoTrnEndChar = " & SetsArray(7)
				LineArray(n + 9) = "CheckShortChar = " & SetsArray(8)
				LineArray(n + 10) = "CheckShortKey = " & SetsArray(9)
				LineArray(n + 11) = "KeepShortKey = " & SetsArray(10)
				LineArray(n + 12) = "PreRepString = " & SetsArray(11)
				LineArray(n + 13) = "AutoRepString = " & SetsArray(12)
				LineArray(n + 14) = "AccessKeyChar = " & SetsArray(13)
				LineArray(n + 15) = "AddAccessKeyWithFirstChar = " & SetsArray(14)
				LineArray(n + 16) = "LineSplitMode = " & SetsArray(15)
				LineArray(n + 17) = "AppInsertSplitChar = " & SetsArray(16)
				LineArray(n + 18) = "ReplaceSplitChar = " & SetsArray(17)
				LineArray(n + 19) = "ApplyLangList = " & getLngPair(TempArray(2),"check")
				If i <> UBound(CheckDataList) Then
					n = UBound(LineArray) + 1
					ReDim Preserve LineArray(n)
				End If
			Next i
		WriteToFile(Path,Join(LineArray,vbCrLf),"unicodeFFFE")
		On Error GoTo 0
		WriteCheckSet = True
		If Path = CheckFilePath Then
			cWriteLoc = CheckFilePath
			GoTo RemoveRegKey
		End If
	'写入注册表
	ElseIf Path = CheckRegKey Then
		On Error GoTo ExitFunction
		SaveSetting("AccessKey","Option","Version",CheckVersion)
		If WriteType = "Main" Or WriteType = "Sets" Or WriteType = "All" Then
			If KeepSet = 1 Then
				SaveSetting("AccessKey","Option","AutoMacroSet",Selected(1))
				SaveSetting("AccessKey","Option","CheckMacroSet",Selected(2))
				SaveSetting("AccessKey","Option","AutoMacroCheck",Selected(3))
				SaveSetting("AccessKey","Option","AutoSelection",Selected(4))
				SaveSetting("AccessKey","Option","SelectedCheck",Selected(5))
				SaveSetting("AccessKey","Option","CheckAllType",Selected(6))
				SaveSetting("AccessKey","Option","CheckMenu",Selected(7))
				SaveSetting("AccessKey","Option","CheckDialog",Selected(8))
				SaveSetting("AccessKey","Option","CheckString",Selected(9))
				SaveSetting("AccessKey","Option","CheckAcceleratorTable",Selected(10))
				SaveSetting("AccessKey","Option","CheckVersion",Selected(11))
				SaveSetting("AccessKey","Option","CheckOther",Selected(12))
				SaveSetting("AccessKey","Option","CheckSeletedOnly",Selected(13))
				SaveSetting("AccessKey","Option","CheckAllCont",Selected(14))
				SaveSetting("AccessKey","Option","CheckAccKey",Selected(15))
				SaveSetting("AccessKey","Option","CheckEndChar",Selected(16))
				SaveSetting("AccessKey","Option","CheckAcceler",Selected(17))
				SaveSetting("AccessKey","Option","IgnoreVerTag",Selected(18))
				SaveSetting("AccessKey","Option","IgnoreSetTag",Selected(19))
				SaveSetting("AccessKey","Option","IgnoreDateTag",Selected(20))
				SaveSetting("AccessKey","Option","IgnoreStateTag",Selected(21))
				SaveSetting("AccessKey","Option","TgnoreAllTag",Selected(22))
				SaveSetting("AccessKey","Option","NoCheckTag",Selected(23))
				SaveSetting("AccessKey","Option","NoChangeTrnState",Selected(24))
				SaveSetting("AccessKey","Option","AutoRepString",Selected(25))
				SaveSetting("AccessKey","Option","KeepSetting",Selected(26))
			End If
		End If
		If WriteType = "Sets" Or WriteType = "All" Then
			'删除原配置项
			Temp = GetSetting("AccessKey","Option","Headers")
			If Temp <> "" Then
				TempArray = ReSplit(Temp,";",-1)
				On Error Resume Next
				For i = 0 To UBound(TempArray)
					DeleteSetting("AccessKey",TempArray(i))
				Next i
				On Error GoTo 0
			End If
			'写入新配置项
			LineArray = CheckDataList
			For i = LBound(CheckDataList) To UBound(CheckDataList)
				LineArray(i) = CStr(i)
				TempArray = ReSplit(CheckDataList(i),JoinStr)
				SetsArray = ReSplit(TempArray(1),SubJoinStr)
				Select Case TempArray(0)
				Case DefaultCheckList(0)
					TempArray(0) = "en2zh"
				Case DefaultCheckList(1)
					TempArray(0) = "zh2en"
				End Select
				SaveSetting("AccessKey",LineArray(i),"Name",TempArray(0))
				SaveSetting("AccessKey",LineArray(i),"ExcludeChar",SetsArray(0))
				SaveSetting("AccessKey",LineArray(i),"LineSplitChar",SetsArray(1))
				SaveSetting("AccessKey",LineArray(i),"CheckBracket",SetsArray(2))
				SaveSetting("AccessKey",LineArray(i),"KeepCharPair",SetsArray(3))
				SaveSetting("AccessKey",LineArray(i),"ShowAsiaKey",SetsArray(4))
				SaveSetting("AccessKey",LineArray(i),"CheckEndChar",SetsArray(5))
				SaveSetting("AccessKey",LineArray(i),"NoTrnEndChar",SetsArray(6))
				SaveSetting("AccessKey",LineArray(i),"AutoTrnEndChar",SetsArray(7))
				SaveSetting("AccessKey",LineArray(i),"CheckShortChar",SetsArray(8))
				SaveSetting("AccessKey",LineArray(i),"CheckShortKey",SetsArray(9))
				SaveSetting("AccessKey",LineArray(i),"KeepShortKey",SetsArray(10))
				SaveSetting("AccessKey",LineArray(i),"PreRepString",SetsArray(11))
				SaveSetting("AccessKey",LineArray(i),"AutoRepString",SetsArray(12))
				SaveSetting("AccessKey",LineArray(i),"AccessKeyChar",SetsArray(13))
				SaveSetting("AccessKey",LineArray(i),"AddAccessKeyWithFirstChar",SetsArray(14))
				SaveSetting("AccessKey",LineArray(i),"LineSplitMode",SetsArray(15))
				SaveSetting("AccessKey",LineArray(i),"AppInsertSplitChar",SetsArray(16))
				SaveSetting("AccessKey",LineArray(i),"ReplaceSplitChar",SetsArray(17))
				SaveSetting("AccessKey",LineArray(i),"ApplyLangList",getLngPair(TempArray(2),"check"))
			Next i
			SaveSetting("AccessKey","Option","Headers",Join(LineArray,";"))
			SaveSetting("AccessKey","Option","UILanguageID",Selected(0))
		End If
		If WriteType = "Update" Or WriteType = "Sets" Or WriteType = "All" Then
			If CheckArray(cUpdateSet) = True Then
				On Error Resume Next
				DeleteSetting("AccessKey","Update")
				On Error GoTo 0
				TempArray = ReSplit(cUpdateSet(1),vbCrLf,-1)
				SaveSetting("AccessKey","Update","UpdateMode",cUpdateSet(0))
				n = 0
				For i = 0 To UBound(TempArray)
					If Trim$(TempArray(i)) <> "" Then
						SaveSetting("AccessKey","Update",CStr(n),TempArray(i))
						n = n + 1
					End If
				Next i
				If n > 0 Then SaveSetting("AccessKey","Update","Count",n - 1)
				SaveSetting("AccessKey","Update","Path",cUpdateSet(2))
				SaveSetting("AccessKey","Update","Argument",cUpdateSet(3))
				SaveSetting("AccessKey","Update","UpdateCycle",cUpdateSet(4))
				SaveSetting("AccessKey","Update","UpdateDate",cUpdateSet(5))
			End If
		End If
		If WriteType = "Project" Or WriteType = "Sets" Or WriteType = "All" Then
			If CheckArray(ProjectList) = True Then
				On Error Resume Next
				DeleteSetting("AccessKey","Projects")
				On Error GoTo 0
				For i = LBound(ProjectDataList) To UBound(ProjectDataList)
					TempArray = ReSplit(ProjectDataList(i),JoinStr)
					Select Case TempArray(0)
					Case DefaultProjectList(0)
						TempArray(0) = "CheckOnly"
					Case DefaultProjectList(1)
						TempArray(0) = "CheckAndCorrect"
					Case DefaultProjectList(2)
						TempArray(0) = "DelAccessKey"
					Case DefaultProjectList(3)
						TempArray(0) = "DelAccelerator"
					Case DefaultProjectList(4)
						TempArray(0) = "DelAccessKeyAndAccelerator"
					End Select
					If TempArray(0) <> "" Then
						SaveSetting("AccessKey","Projects",TempArray(0),TempArray(1))
					End If
				Next i
			End If
		End If
		WriteCheckSet = True
		cWriteLoc = CheckRegKey
		GoTo RemoveFilePath
	'删除所有保存的设置
	ElseIf Path = "" Then
		'删除文件配置项
 		RemoveFilePath:
		On Error Resume Next
		If Dir(CheckFilePath) <> "" Then
			SetAttr CheckFilePath,vbNormal
			Kill CheckFilePath
		End If
		Temp = Left(CheckFilePath,InStrRev(CheckFilePath,"\"))
		If Dir(Temp & "*.*") = "" Then RmDir Temp
		On Error GoTo 0
		If Path = CheckRegKey Then GoTo ExitFunction
		'删除注册表配置项
		RemoveRegKey:
		If GetSetting("AccessKey","Option","Version") <> "" Then
			Temp = GetSetting("AccessKey","Option","Headers")
			On Error Resume Next
			If Temp <> "" Then
				TempArray = ReSplit(Temp,";",-1)
				For i = 0 To UBound(TempArray)
					DeleteSetting("AccessKey",TempArray(i))
				Next i
			End If
			DeleteSetting("AccessKey","Option")
			DeleteSetting("AccessKey","Update")
			DeleteSetting("AccessKey","Projects")
			Dim WshShell As Object
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.RegDelete CheckRegKey
			Set WshShell = Nothing
			On Error GoTo 0
		End If
		If Path = CheckFilePath Then GoTo ExitFunction
		'设置写入位置设置为空
		WriteCheckSet = True
		cWriteLoc = ""
	End If
	ExitFunction:
End Function


'更新引擎旧版本配置值
Function EngineDataUpdate(ByVal Header As String,ByVal Data As String) As String
	Dim i As Long,TempArray() As String,dSets() As String,uSets() As String
	EngineDataUpdate = Data
	For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
		If DefaultEngineList(i) = Header Then
			Data = "1"
			Exit For
		End If
	Next i
	If Data <> "1" Then Exit Function
	TempArray = ReSplit(EngineDataUpdate,JoinStr)
	uSets = ReSplit(TempArray(1),SubJoinStr)
	dSets = ReSplit(EngineSettings(Header),SubJoinStr)
	For i = 0 To UBound(uSets)
		If Trim$(uSets(i)) = "" Then
			uSets(i) = dSets(i)
		ElseIf uSets(i) <> dSets(i) Then
			uSets(i) = dSets(i)
		End If
	Next i
	TempArray(1) = Join$(uSets,SubJoinStr)
	EngineDataUpdate = Join$(TempArray,JoinStr)
End Function


'更新检查旧版本配置值
Function CheckDataUpdate(ByVal Header As String,ByVal Data As String) As String
	Dim i As Long,j As Long,dSets() As String,uSets() As String
	Dim TempArray() As String,TempList() As String
	CheckDataUpdate = Data
	For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
		If DefaultCheckList(i) = Header Then
			j = 1
			Exit For
		End If
	Next i
	If j = 0 Then Exit Function
	Data = CheckSettings(Header,0)
	If CheckDataUpdate = Data Then Exit Function
	dSets = ReSplit(Data,SubJoinStr)
	If CheckArray(dSets) = False Then Exit Function
	TempArray = ReSplit(CheckDataUpdate,JoinStr)
	uSets = ReSplit(TempArray(1),SubJoinStr)
	For i = 0 To UBound(uSets)
		If Trim$(uSets(i)) = "" Or i = 1 Or i = 6 Then
			uSets(i) = dSets(i)
		ElseIf uSets(i) <> dSets(i) Then
			If i <> 4 And i <> 14 And i <> 15 And i < 18 Then
				If i = 5 Or i = 7 Then Data = " " Else Data = ","
				If i = 7 And InStr(uSets(i),"|") = 0 Then
					TempList = ReSplit(uSets(i),Data)
					For j = 0 To UBound(TempList)
						TempList(j) = Left$(Trim$(TempList(j)),1) & "|" & Right$(Trim$(TempList(j)),1)
					Next j
					uSets(i) = Join$(TempList,Data)
				End If
				uSets(i) = Join$(ClearArray(ReSplit(uSets(i) & Data & dSets(i),Data,-1)),Data)
			End If
		End If
	Next i
	TempArray(1) = Join$(uSets,SubJoinStr)
	CheckDataUpdate = Join$(TempArray,JoinStr)
End Function


'增加或更改数组项目
Function CreateArray(ByVal Header As String,ByVal Data As String,HeaderList() As String,DataList() As String) As Boolean
	Dim i As Long,n As Long,FindHeader As String,Stemp As Boolean
	If HeaderList(0) = "" Then
		HeaderList(0) = Header
		DataList(0) = Data
		CreateArray = True
	Else
		n = UBound(HeaderList)
		FindHeader = LCase$(Header)
		For i = LBound(HeaderList) To n
			If LCase$(HeaderList(i)) = FindHeader Then
				If DataList(i) <> Data Then DataList(i) = Data
				Stemp = True
				Exit For
			End If
		Next i
		If Stemp = False Then
			n = n + 1
			ReDim Preserve HeaderList(n),DataList(n)
			HeaderList(n) = Header
			If DataList(n) <> Data Then DataList(n) = Data
			CreateArray = True
		End If
	End If
End Function


'删除数组项目
Sub DelArray(List() As String,ByVal IDList As Variant,Optional ByVal Separator As String)
	Dim i As Long,n As Long
	If IsArray(IDList) Then
		If UBound(List) = UBound(IDList) Then
			ReDim List(0) As String
			Exit Sub
		End If
		ReDim Stemp(UBound(List)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(List)
			If Stemp(i) = 0 Then
				List(n) = List(i)
				n = n + 1
			End If
		Next i
	ElseIf IsNumeric(IDList) And Separator = "" Then
		n = IDList
		For i = IDList + 1 To UBound(List)
			List(n) = List(i)
			n = n + 1
		Next i
	ElseIf IDList <> "" Then
		For i = LBound(List) To UBound(List)
			If Separator <> "" Then
				If ReSplit(List(i),Separator)(0) <> IDList Then
					List(n) = List(i)
					n = n + 1
				End If
			ElseIf List(i) <> IDList Then
				List(n) = List(i)
				n = n + 1
			End If
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve List(n - 1) As String
	Else
		ReDim List(0) As String
	End If
End Sub


'删除数组项目
Sub DelArrays(List() As String,DataList() As String,IDList As Variant)
	Dim i As Long,n As Long
	If IsArray(IDList) Then
		If UBound(List) = UBound(IDList) Then
			ReDim List(0) As String,DataList(0) As String
			Exit Sub
		End If
		ReDim Stemp(UBound(List)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(List)
			If Stemp(i) = 0 Then
				List(n) = List(i)
				DataList(n) = DataList(i)
				n = n + 1
			End If
		Next i
	ElseIf IsNumeric(IDList) Then
		n = IDList
		For i = IDList + 1 To UBound(List)
			List(n) = List(i)
			DataList(n) = DataList(i)
			n = n + 1
		Next i
	ElseIf IDList <> "" Then
		For i = LBound(List) To UBound(List)
			If List(i) <> IDList Then
				List(n) = List(i)
				DataList(n) = DataList(i)
				n = n + 1
			End If
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve List(n - 1) As String,DataList(n - 1) As String
	Else
		ReDim List(0) As String,DataList(0) As String
	End If
End Sub


'删除自定义工具数组项目
Sub DelToolsArrays(List() As String,DataList() As TOOLS_PROPERTIE,ByVal IDList As Variant)
	Dim i As Long,n As Long
	If IsArray(IDList) Then
		If UBound(DataList) = UBound(IDList) Then
			ReDim List(0) As String,DataList(0) As TOOLS_PROPERTIE
			Exit Sub
		End If
		ReDim Stemp(UBound(DataList)) As Long
		For i = LBound(IDList) To UBound(IDList)
			If IDList(i) > -1 Then Stemp(IDList(i)) = 1
		Next i
		n = IDList(LBound(IDList))
		For i = IDList(LBound(IDList)) To UBound(DataList)
			If Stemp(i) = 0 Then
				List(n) = List(i)
				DataList(n) = DataList(i)
				n = n + 1
			End If
		Next i
	ElseIf IsNumeric(IDList) Then
		n = IDList
		For i = IDList + 1 To UBound(DataList)
			List(n) = List(i)
			DataList(n) = DataList(i)
			n = n + 1
		Next i
	ElseIf IDList <> "" Then
		For i = LBound(List) To UBound(List)
			If List(i) <> IDList Then
				List(n) = List(i)
				DataList(n) = DataList(i)
				n = n + 1
			End If
		Next i
	Else
		Exit Sub
	End If
	If n > 0 Then
		ReDim Preserve List(n - 1) As String,DataList(n - 1) As TOOLS_PROPERTIE
	Else
		ReDim List(0) As String,DataList(0) As TOOLS_PROPERTIE
	End If
End Sub


'拆分数组
Function SplitData(ByVal Data As String,NameList() As String,SrcList() As String,TranList() As String) As Boolean
	Dim i As Long,TempList() As String,TempArray() As String
	TempArray = ReSplit(Data,SubLngJoinStr)
	i = UBound(TempArray)
	ReDim NameList(i),SrcList(i),TranList(i)
	For i = 0 To UBound(TempArray)
		TempList = ReSplit(TempArray(i),LngJoinStr)
		If TempList(0) <> "" Then
			If TempList(1) = "" Then TempList(1) = NullValue
			If TempList(2) = "" Then TempList(2) = NullValue
			NameList(i) = TempList(0)
			SrcList(i) = TempList(1)
			TranList(i) = TempList(2)
			SplitData = True
		End If
	Next i
End Function


'创建二个互补的名称数组
Function getLngNameList(ByVal Data As String,NameList() As String,SrcList() As String) As Boolean
	Dim i As Long,j As Long,n As Long,TempList() As String,TempArray() As String
	TempArray = ReSplit(Data,SubLngJoinStr)
	i = UBound(TempArray)
	ReDim NameList(i),SrcList(i)
	For i = 0 To UBound(TempArray)
		If TempArray(i) <> "" Then
			TempList = ReSplit(TempArray(i),LngJoinStr)
			If TempList(2) = "" Then
				NameList(j) = TempList(0)
				j = j + 1
			Else
				SrcList(n) = TempList(0)
				n = n + 1
			End If
		End If
	Next i
	If j > 0 Then ReDim Preserve NameList(j - 1) Else ReDim NameList(0)
	If n > 0 Then ReDim Preserve SrcList(n - 1) Else ReDim SrcList(0)
	If j + n > 0 Then getLngNameList = True
End Function


'合并标准和自定义语言列表
Function MergeLngList(LangArray() As String,DataArray() As String,ByVal fType As String) As String()
	Dim i As Long,j As Long,n As Long,TempList() As String,TempArray() As String,Dic As Object
	Set Dic = CreateObject("Scripting.Dictionary")
	n = UBound(LangArray)
	For i = LBound(LangArray) To n
		LangCode = LCase$(ReSplit(LangArray(i),LngJoinStr)(1))
		If Not Dic.Exists(LangCode) Then
			Dic.Add(LangCode,i)
		End If
	Next i
	TempList = LangArray
	ReDim Preserve TempList(n + UBound(DataArray))
	For i = LBound(DataArray) To UBound(DataArray)
		LangDataList = ReSplit(DataArray(i),LngJoinStr)
		LangCode = LCase$(LangDataList(1))
		If Dic.Exists(LangCode) Then
			j = Dic.Item(LangCode)
			TempArray = ReSplit(LangArray(j),LngJoinStr)
			If fType = "check" Then
				TempList(j) = TempArray(0) & LngJoinStr & TempArray(1) & LngJoinStr & LangDataList(1)
			ElseIf fType = "engine" Then
				TempList(j) = TempArray(0) & LngJoinStr & TempArray(1) & LngJoinStr & LangDataList(2)
			End If
		ElseIf LangDataList(0) <> "" Then
			n = n + 1
			If fType = "check" Then
				TempList(n) = DataArray(i) & LngJoinStr & LangDataList(1)
			ElseIf fType = "engine" Then
				TempList(n) = DataArray(i)
			End If
		End If
	Next i
	ReDim Preserve TempList(n)
	Set Dic = Nothing
	MergeLngList = TempList
End Function


'生成去除数据列表中空项后的语言对
Function getLngPair(ByVal Data As String,ByVal fType As String) As String
	Dim i As Long,n As Long,TempList() As String,Dic As Object
	Set Dic = CreateObject("Scripting.Dictionary")
	For i = LBound(PslLangDataList) To UBound(PslLangDataList)
		LangCode = LCase$(ReSplit(PslLangDataList(i),LngJoinStr)(1))
		If Not Dic.Exists(LangCode) Then
			Dic.Add(LangCode,"")
		End If
	Next i
	LangArray = ReSplit(Data,SubLngJoinStr)
	ReDim TempList(UBound(LangArray))
	For i = 0 To UBound(LangArray)
		LangPairList = ReSplit(LangArray(i),LngJoinStr)
		If LangPairList(0) <> "" Then
			SrcCode = IIf(LangPairList(1) <> NullValue,LangPairList(1),"")
			TranCode = IIf(LangPairList(2) <> NullValue,LangPairList(2),"")
			LangCode = LCase$(SrcCode)
			If TranCode <> "" Then
				If fType = "check" Then
					If Dic.Exists(LangCode) Then
						TempList(n) = LngJoinStr & TranCode
					Else
						TempList(n) = LangPairList(0) & LngJoinStr & TranCode
					End If
				ElseIf fType = "engine" Then
					If Dic.Exists(LangCode) Then
						TempList(n) = LngJoinStr & SrcCode & LngJoinStr & TranCode
					Else
						TempList(n) = LangArray(i)
					End If
				Else
					TempList(n) = LangArray(i)
				End If
				n = n + 1
			End If
		End If
	Next i
	If n > 0 Then
		ReDim Preserve TempList(n - 1)
		getLngPair = Join(TempList,SubLngJoinStr)
	End If
End Function


'互换数组项目
Function ChangeList(ByVal Data As String,UseList() As String) As String()
	Dim i As Long,n As Long,TempList() As String,Dic As Object
	Set Dic = CreateObject("Scripting.Dictionary")
	For i = LBound(UseList) To UBound(UseList)
		If Not Dic.Exists(UseList(i)) Then
			Dic.Add(UseList(i),"")
		End If
	Next i
	LangArray = ReSplit(Data,SubLngJoinStr)
	ReDim TempList(UBound(LangArray))
	For i = 0 To UBound(LangArray)
		LangPairList = ReSplit(LangArray(i),LngJoinStr)
		If Not Dic.Exists(LangPairList(0)) Then
			TempList(n) = LangPairList(0)
			n = n + 1
		End If
	Next i
	Set Dic = Nothing
	If n > 0 Then n = n - 1
	ReDim Preserve TempList(n)
	ChangeList = TempList
End Function


'查找语言对是否存在
Function getEngineLngPair(LangArray() As String,ByVal srcLang As String,ByVal trnLang As String) As String
	Dim i As Long,j As Long,k As Long,TempList() As String
	srcLang = LngJoinStr & LCase$(srcLang) & LngJoinStr
	trnLang = LngJoinStr & LCase$(trnLang) & LngJoinStr
	For i = 0 To UBound(LangArray)
		TempList = ReSplit(LCase$(LangArray(i)),LngJoinStr)
		If TempList(2) <> "" Then
			If j = 0 Then
				If InStr(srcLang,LngJoinStr & TempList(1) & LngJoinStr) Then
					srcLang = TempList(2)
					j = 1
				End If
			End If
			If k = 0 Then
				If InStr(trnLang,LngJoinStr & TempList(1) & LngJoinStr) Then
					trnLang = TempList(2)
					k = 1
				End If
			End If
			If j + k = 2 Then
				getEngineLngPair = srcLang & LngJoinStr & trnLang
				Exit For
			End If
		End If
	Next i
End Function


'查找指定值是否在数组中
Function getCheckID(DataList() As String,ByVal LngCode As String,ByVal OldLngCode As String) As Long
	Dim i As Long,j As Long,TempList() As String
	getCheckID = -1
	LngCode = LngJoinStr & LCase$(LngCode) & LngJoinStr
	For i = LBound(DataList) To UBound(DataList)
		TempList = ReSplit(DataList(i),JoinStr)
		If TempList(2) <> "" Then
			TempList = ReSplit(LCase$(TempList(2)),SubLngJoinStr)
			For j = 0 To UBound(TempList)
				If InStr(LngCode,LngJoinStr & ReSplit(TempList(j),LngJoinStr)(2) & LngJoinStr) Then
					getCheckID = i
					Exit Function
				End If
			Next j
		End If
	Next i
	If OldLngCode = "Asia" Then getCheckID = 0 Else getCheckID = 1
End Function


'检查数组中是否有空值
'Mode = 0     检查多项数组项内是否全为空值
'Mode = 1     检查多项数组项内是否有空值
'Mode = 2     仅检查多项数组的参数项内是否全为空值
'Mode = 3     检查多项数组的参数项内是否有空值
'Mode = 4     仅检查多项数组的语言项内是否全为空值
'Mode = 5     检查多项数组的语言项内是否有空值
'Mode = 6     检查单项数组项内是否有空值
'Header = ""   检查整个数组
'Header <> ""  检查指定数组项
Function CheckNullData(ByVal Header As String,DataList() As String,ByVal SkipNum As String,ByVal Mode As Long) As Boolean
	Dim i As Long,j As Long,m As Long,n As Long,dMax As Long,nMax As Long,sMax As Long
	Dim SkipArray() As String,TempArray() As String,SetsArray() As String
	Dim Stemp As Boolean,Dic As Object

	Set Dic = CreateObject("Scripting.Dictionary")
	SkipArray = ReSplit(SkipNum,",",-1)
	For i = 0 To UBound(SkipArray)
		TempArray = ReSplit(SkipArray(i),"-")
		For j = CLng(TempArray(0)) To CLng(TempArray(UBound(TempArray)))
			If Not Dic.Exists(j) Then
				Dic.Add(j,"")
			End If
		Next j
	Next i

	dMax = UBound(DataList)
	nMax = Dic.Count - 1
	If Header = "" Then Stemp = True
	For i = LBound(DataList) To dMax
		If Mode = 6 Then
			If Not Dic.Exists(i) Then
				If Trim(DataList(i)) = "" Then
					CheckNullData = True
					Exit For
				End If
			End If
		Else
			n = 0
			TempArray = ReSplit(DataList(i),JoinStr)
			SetsArray = ReSplit(TempArray(1),SubJoinStr)
			sMax = UBound(SetsArray)
			If Stemp = False And TempArray(0) = Header Then Stemp = True
			If Stemp = True Then
				If Mode < 4 Then
					For j = 0 To sMax
						If Not Dic.Exists(j) Then
							If Trim(SetsArray(j)) = "" Then
								If Mode = 0 Or Mode = 2 Then
									n = n + 1
								ElseIf Mode = 1 Or Mode = 3 Then
									CheckNullData = True
									Exit For
								End If
							End If
						End If
					Next j
				End If
				If Mode <> 2 And Mode <> 3 Then
					If getLngPair(TempArray(2),"") = "" Then
						If Mode = 0 Or Mode = 4 Then
							n = n + 1
						ElseIf Mode = 1 Or Mode = 5 Then
							CheckNullData = True
						End If
					End If
				End If
			End If
			If Mode = 0 Then
				If Header <> "" Then
					If n = sMax - nMax + 1 Then CheckNullData = True
				Else
					If n = sMax - nMax + 1 Then m = m + 1
					If m = dMax + 1 Then CheckNullData = True
				End If
			ElseIf Mode = 2 Then
				If Header <> "" Then
					If n = sMax - nMax Then CheckNullData = True
				Else
					If n = sMax - nMax Then m = m + 1
					If m = dMax + 1 Then CheckNullData = True
				End If
			ElseIf Mode = 4 Then
				If Header <> "" Then
					If n = 1 Then CheckNullData = True
				Else
					If n = dMax + 1 Then CheckNullData = True
				End If
			End If
			If CheckNullData = True Then Exit For
		End If
	Next i
	If Mode <> 6 And Header <> "" And Stemp = False Then CheckNullData = True
	Set Dic = Nothing
End Function


'检查提取翻译用标志字符值是否为空
Function CheckTargetValue(DataList() As String,ByVal ID As Long) As Boolean
	Dim i As Long,BeforeStr As String,AfterStr As String
	CheckTargetValue = False
	For i = LBound(DataList) To UBound(DataList)
		TranBeforeStr = ""
		TranAfterStr = ""
		If i = ID Or ID < 0 Then
			TempArray = ReSplit(DataList(i),JoinStr)
			SetsArray = ReSplit(TempArray(1),SubJoinStr)
			Select Case SetsArray(10)
			Case "responseText"
				BeforeStr = SetsArray(11)
				AfterStr = SetsArray(12)
			Case "responseBody"
				BeforeStr = SetsArray(13)
				AfterStr = SetsArray(14)
			Case "responseStream"
				BeforeStr = SetsArray(15)
				AfterStr = SetsArray(16)
			Case "responseXML"
				BeforeStr = SetsArray(17)
				AfterStr = SetsArray(18)
			End Select
			If BeforeStr = "" Or AfterStr = "" Then
				CheckTargetValue = True
				Exit For
			End If
		End If
	Next i
End Function


'按字串的字符数递归法排序字串数组
'Mode = False 从小到大排序，否则从大到小排序，l = 数组的左边界，r = 数组的右边界.
Sub SortArrayByLength(ByRef MyArray As Variant, ByVal l As Long, ByVal r As Long, ByVal Mode As Boolean)
	Dim i As Long, j As Long, TmpX As Variant, TmpA As Variant
	i = l: j = r: TmpX = Len(MyArray((l + r) \ 2))
	While (i <= j)
		If Mode = False Then
			While (Len(MyArray(i)) < TmpX And i < r)
				i = i + 1
			Wend
			While (TmpX < Len(MyArray(j)) And j > l)
				j = j - 1
			Wend
		Else
			While (Len(MyArray(i)) > TmpX And i < r)
				i = i + 1
			Wend
			While (TmpX > Len(MyArray(j)) And j > l)
				j = j - 1
			Wend
		End If
		If (i <= j) Then
			TmpA = MyArray(i)
			MyArray(i) = MyArray(j)
			MyArray(j) = TmpA
			i = i + 1: j = j - 1
		End If
	Wend
	If (l < j) Then Call SortArrayByLength(MyArray, l, j, Mode)
	If (i < r) Then Call SortArrayByLength(MyArray, i, r, Mode)
End Sub


'清理数组中重复的数据
'Mode = 0 不清除空置项，否则清除空置项
Function ClearArray(xArray() As String,Optional ByVal Mode As Long) As String()
	Dim i As Long,n As Long,yArray() As String,Dic As Object
	ClearArray = xArray
	n = UBound(xArray)
	If n = 0 Then Exit Function
	Set Dic = CreateObject("Scripting.Dictionary")
	ReDim yArray(n)
	n = 0
	For i = LBound(xArray) To UBound(xArray)
		If Mode > 0 Then
			If xArray(i) <> "" Then
				If Not Dic.Exists(xArray(i)) Then
					Dic.Add(xArray(i),"")
					yArray(n) = xArray(i)
					n = n + 1
				End If
			End If
		ElseIf Not Dic.Exists(xArray(i)) Then
			Dic.Add(xArray(i),"")
			yArray(n) = xArray(i)
			n = n + 1
		End If
	Next i
	If n > 0 Then n = n - 1
	ReDim Preserve yArray(n)
	ClearArray = yArray
End Function


'检查字串数组是否为空，非空返回 True
Function CheckArray(DataList() As String) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i) <> "" Then
			CheckArray = True
			Exit For
		End If
	Next i
	errHandle:
End Function


'检查 INI 数据数组是否为空，非空返回 True
Function CheckINIArray(DataList() As INIFILE_DATA) As Boolean
	Dim i As Long
	On Error GoTo errHandle
	For i = LBound(DataList) To UBound(DataList)
		If DataList(i).Title <> "" Then
			CheckINIArray = True
			Exit For
		End If
	Next i
	errHandle:
End Function


'比较二个字串数组是否相同，不相同返回 True
Function ArrayComp(uArray() As String,oArray() As String,Optional ByVal Index As String) As Boolean
	Dim i As Long,j As Long,SkipArray() As String,TempArray() As String
	If UBound(uArray) <> UBound(oArray) Then
		ArrayComp = True
		Exit Function
	End If
	If Index = "" Then
		For i = LBound(uArray) To UBound(uArray)
			If oArray(i) <> uArray(i) Then
				ArrayComp = True
				Exit For
			End If
		Next i
	Else
		SkipArray = ReSplit(Index,",",-1)
		For i = 0 To UBound(SkipArray)
			TempArray = ReSplit(SkipArray(i),"-")
			For j = CLng(TempArray(0)) To CLng(TempArray(UBound(TempArray)))
				If oArray(j) <> uArray(j) Then
					ArrayComp = True
					Exit Function
				End If
			Next j
		Next i
	End If
End Function


'插入数据到数组
'Mode = True 不允许重复项并插入或移位 Data 到最前面
Function InsertArray(List() As String,ByVal Data As String,ByVal insPos As Long,Optional ByVal Mode As Boolean) As Boolean
	Dim i As Long,j As Long
	i = LBound(List)
	j = UBound(List)
	If j = i And List(i) = "" Then
		List(i) = Data
	Else
		If Mode = True Then
			If insPos = 0 Then
				If List(0) = Data Then Exit Function
			ElseIf InStr(vbNullChar & Join$(List,vbNullChar) & vbNullChar,vbNullChar & Data & vbNullChar) Then
				Exit Function
			End If
		End If
		ReDim Preserve List(j + 1) As String
		If insPos <= j Then
			For i = j + 1 To insPos + 1 Step -1
				List(i) = List(i - 1)
			Next i
			List(insPos) = Data
		Else
			List(j + 1) = Data
		End If
		If Mode = True Then List = ClearArray(List,1)
	End If
	InsertArray = True
End Function


'通配符查找指定值
Function CheckKeyCode(ByVal FindKey As String,ByVal CheckKey As String) As Long
	Dim FindStr As String,Key As String,Pos As Long
	Key = Trim(FindKey)
	If InStr(Key,"%") Then Key = Replace(Key,"%","x")
	If CheckKey <> "" And Key <> "" Then
		FindStrArr = ReSplit(Convert(CheckKey),",",-1)
		For i = 0 To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If InStr(FindStr,"%") Then FindStr = Replace(FindStr,"%","x")
			If InStr(FindStr,"-") Then
				If Left(FindStr,1) <> "[" And Right(FindStr,1) <> "]" Then
					FindStr = "[" & FindStr & "]"
				End If
			End If
			Pos = InStr(FindStr,"[")
			If Pos > 0 Then
				If Left(FindStr,Pos-1) <> "[" And Right(FindStr,Pos+1) <> "]" Then
					FindStr = Replace(FindStr,"[","[[]")
				End If
			End If
			'PSL.Output Key & " : " &  FindStr  '调试用
			If UCase(Key) Like UCase(FindStr) = True Then
				CheckKeyCode = 1
				Exit For
			End If
		Next i
	ElseIf CheckKey = "" And Key <> "" Then
		CheckKeyCode = 1
	End If
End Function


'测试在线翻译程序
Sub TranTest(ByVal EngineID As Long,ByVal fType As Long)
	Dim MsgList() As String
	If getMsgList(UIDataList,MsgList,"TranTest",1) = False Then Exit Sub
	Begin Dialog UserDialog 660,518,MsgList(0),.TranTestFunc ' %GRID:10,7,1,1
		GroupBox 10,42,640,70,"",.GroupBox
		Text 10,7,640,28,MsgList(1),.MainText
		Text 30,59,90,14,MsgList(2),.SetNameText
		DropListBox 130,56,200,21,EngineList(),.SelSetBox
		Text 340,59,90,14,MsgList(3),.LngNameText
		DropListBox 430,56,200,21,MsgList(),.LngNameBox
		Text 30,87,90,14,MsgList(4),.TrnListText
		DropListBox 130,84,300,21,MsgList(),.TrnListBox
		Text 450,87,120,14,MsgList(5),.LineNumText
		TextBox 580,84,50,21,.LineNumBox
		Text 10,126,360,14,MsgList(6),.InText
		TextBox 10,147,640,147,.InTextBox,1
		OptionGroup .StrType
			OptionButton 380,126,130,14,MsgList(8),.SrcString
			OptionButton 520,126,130,14,MsgList(9),.TranString
		Text 10,301,360,14,MsgList(7),.OutText
		TextBox 10,322,640,154,.OutTextBox,1
		OptionGroup .TranType
			OptionButton 380,301,130,14,MsgList(10),.TranOnly
			OptionButton 520,301,130,14,MsgList(11),.AllTran
		PushButton 10,490,100,21,MsgList(12),.HelpButton
		PushButton 120,490,100,21,MsgList(13),.TestButton
		PushButton 230,490,100,21,MsgList(14),.HeaderButton
		PushButton 340,490,100,21,MsgList(15),.ClearButton
		CancelButton 550,490,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.SelSetBox = EngineID
	dlg.TranType = fType
	If Dialog(dlg) = 0 Then Exit Sub
End Sub


'测试对话框函数
Private Function TranTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,n As Long,m As Long,LineNum As Long,EngineID As Long,CheckID As Long
	Dim inText As String,PslLng As String,srcLng As String,trnLng As String,Temp As String
	Dim TempList() As String,MsgList() As String
	Dim TrnList As PslTransList,xmlHttp As Object
	Dim mCheckSrc As Long,ProjectIDSrc As Long,ProjectIDTrn As Long,mCheckTrn As Long
	Select Case Action%
	Case 1
		If getMsgList(UIDataList,MsgList,"TranTestFunc",1) = False Then Exit Function
		If trn.SourceList.LastChange > trn.LastUpdate Then trn.Update
		m = 10
		If trn.StringCount < m Then m = trn.StringCount
		DlgText "LineNumBox",CStr$(m)
		ReDim TempList(m - 1) As String
		For i = 1 To trn.StringCount
			Set TransString = trn.String(i)
			If TransString.Text <> "" Then
				If DlgValue("StrType") = 0 Then
					TempList(n) = TransString.SourceText
				Else
					TempList(n) = TransString.Text
				End If
				n = n + 1
				If n = m Then Exit For
			End If
		Next i
		ReDim Preserve TempList(IIf(n = 0,0,n - 1)) As String
		DlgText "InTextBox",Join$(TempList,vbCrLf)
		If DlgText("InTextBox") <> "" Then
			DlgText "ClearButton",MsgList(0)
		Else
			DlgText "ClearButton",MsgList(1)
			DlgEnable "ClearButton",False
			DlgEnable "HeaderButton",False
		End If
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
		If trnLng = "" Or trnLng = "zh" Then
			trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
			If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then
				trnLng = "zh-CN"
			ElseIf trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then
				trnLng = "zh-TW"
			End If
		Else
			trnLng = trnLng & LngJoinStr & PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
		End If
		trnLng = LngJoinStr & LCase$(trnLng) & LngJoinStr
		n = 0
		LangArray = ReSplit(ReSplit(EngineDataList(DlgValue("SelSetBox")),JoinStr)(2),SubLngJoinStr)
		ReDim TempList(UBound(LangArray)) As String
		For i = 0 To UBound(LangArray)
			LangPairList = ReSplit(LangArray(i),LngJoinStr)
			If LangPairList(2) <> "" Then
				TempList(n) = LangPairList(0)
				If InStr(trnLng,LngJoinStr & LCase$(LangPairList(1)) & LngJoinStr) Then m = n
				n = n + 1
			End If
		Next i
		ReDim Preserve TempList(IIf(n = 0,0,n - 1)) As String
		DlgListBoxArray "LngNameBox",TempList()
		DlgValue "LngNameBox",m
		ReDim TempList(trn.Project.TransLists.Count - 1)
		For i = 1 To trn.Project.TransLists.Count
			Set TrnList = trn.Project.TransLists(i)
			TempList(i - 1) = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
		Next i
		DlgListBoxArray "TrnListBox",TempList()
		DlgText "TrnListBox",trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		TranTestFunc = True '防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			TranTestFunc = False '按下按钮关闭对话框窗口
			Exit Function
		Case "HelpButton"
			Call Help("EngineTestHelp")
			Exit Function
		Case "SelSetBox"
			TempArray = ReSplit(EngineDataList(DlgValue("SelSetBox")),JoinStr)
			LangArray = ReSplit(TempArray(2),SubLngJoinStr)
			ReDim TempList(UBound(LangArray)) As String
			For i = 0 To UBound(LangArray)
				LangPairList = ReSplit(LangArray(i),LngJoinStr)
				If LangPairList(2) <> "" Then
					TempList(n) = LangPairList(0)
					n = n + 1
				End If
			Next i
			ReDim Preserve TempList(IIf(n = 0,0,n - 1)) As String
			Temp = DlgText("LngNameBox")
			DlgListBoxArray "LngNameBox",TempList()
			DlgText "LngNameBox",Temp
			If DlgText("LngNameBox") = "" Then DlgValue "LngNameBox",0
		Case "TrnListBox", "StrType", "ClearButton"
			If getMsgList(UIDataList,MsgList,"TranTestFunc",1) = False Then Exit Function
			If DlgItem$ <> "ClearButton" Or DlgText("ClearButton") = MsgList(1) Then
				Temp = DlgText("TrnListBox")
				Set TrnList = trn
				For i = 1 To trn.Project.TransLists.Count
					Set TrnList = trn.Project.TransLists(i)
					If Temp = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) Then
						Exit For
					End If
				Next i
				If TrnList.SourceList.LastChange > TrnList.LastUpdate Then TrnList.Update
				m = StrToLong(DlgText("LineNumBox"))
				If TrnList.StringCount < m Then m = TrnList.StringCount
				n = 0
				ReDim TempList(m - 1) As String
				For i = 1 To TrnList.StringCount
					Set TransString = TrnList.String(i)
					If TransString.Text <> "" Then
						If DlgValue("StrType") = 0 Then
							TempList(n) = TransString.SourceText
						Else
							TempList(n) = TransString.Text
						End If
						n = n + 1
						If n = m Then Exit For
					End If
				Next i
				ReDim Preserve TempList(IIf(n = 0,0,n - 1)) As String
				DlgText "InTextBox",Join$(TempList,vbCrLf)
				If DlgText("InTextBox") <> "" Then
					DlgText "ClearButton",MsgList(0)
					DlgEnable "TestButton",True
					DlgEnable "HeaderButton",True
				Else
					DlgText "ClearButton",MsgList(1)
   					DlgEnable "TestButton",False
					DlgEnable "HeaderButton",False
    			End If
			ElseIf DlgText("ClearButton") = MsgList(0) Then
				DlgText "InTextBox",""
				DlgText "OutTextBox",""
   				DlgText "ClearButton",MsgList(1)
   				DlgEnable "TestButton",False
				DlgEnable "HeaderButton",False
			End If
		Case "TestButton", "TranType", "HeaderButton"
			If getMsgList(UIDataList,MsgList,"TranTestFunc",1) = False Then Exit Function
			EngineID = DlgValue("SelSetBox")
			inText = DlgText("InTextBox")
			DlgText "OutTextBox",""
			DlgText "OutTextBox",MsgList(2)

			'获取翻译列表的来源和目标语言
			Temp = DlgText("TrnListBox")
			Set TrnList = trn
			For i = 1 To trn.Project.TransLists.Count
				Set TrnList = trn.Project.TransLists(i)
				If Temp = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) Then
					Exit For
				End If
			Next i
			If DlgValue("StrType") = 0 Then
				srcLng = PSL.GetLangCode(TrnList.SourceList.LangID,pslCode639_1)
			Else
				srcLng = PSL.GetLangCode(TrnList.Language.LangID,pslCode639_1)
			End If
			If srcLng = "" Or srcLng = "zh" Then
				If DlgValue("StrType") = 0 Then
					srcLng = PSL.GetLangCode(TrnList.SourceList.LangID,pslCodeLangRgn)
				Else
					srcLng = PSL.GetLangCode(TrnList.Language.LangID,pslCodeLangRgn)
				End If
				If srcLng = "zh-CHS" Or srcLng = "zh-SG" Then
					srcLng = "zh-CN"
				ElseIf srcLng = "zh-CHT" Or srcLng = "zh-HK" Or srcLng = "zh-MO" Then
					srcLng = "zh-TW"
				End If
			Else
				srcLng = srcLng & LngJoinStr & PSL.GetLangCode(TrnList.SourceList.LangID,pslCodeLangRgn)
			End If
			'检测 Microsoft.XMLHTTP 是否存在
			TempArray = ReSplit(EngineDataList(EngineID),JoinStr)
			TempDataList = ReSplit(TempArray(1),SubJoinStr)
			TempList = ReSplit(TempDataList(0),IIf(InStr(TempDataList(0),";"),";",","))
			On Error Resume Next
			For i = 0 To UBound(TempList)
				Temp = Trim(TempList(i))
				If Temp <> "" Then
					Set xmlHttp = CreateObject(Temp)
					If Not xmlHttp Is Nothing Then Exit For
				End If
			Next i
			If xmlHttp Is Nothing Then
				Err.Source = Join(TempList,"; ")
				Call sysErrorMassage(Err,2)
				TranTestFunc = True '防止按下按钮关闭对话框窗口
				Exit Function
			End If
			On Error GoTo 0
			'获取翻译引擎的来源和目标语言
			n = 0: m = 0
			Temp = LCase$(DlgText("LngNameBox"))
			srcLng = LngJoinStr & LCase$(srcLng) & LngJoinStr
			'TempArray = ReSplit(EngineDataList(EngineID),JoinStr)
			LangArray = ReSplit(TempArray(2),SubLngJoinStr)
			For i = 0 To UBound(LangArray)
				LangPairList = ReSplit(LangArray(i),LngJoinStr)
				If LangPairList(2) <> "" Then
					If n = 0 Then
						If InStr(srcLng,LngJoinStr & LCase$(LangPairList(1)) & LngJoinStr) Then
							srcLng = LangPairList(2)
							n = 1
						End If
					End If
					If m = 0 Then
						If Temp = LCase$(LangPairList(0)) Then
							PslLng = LCase$(LangPairList(1))
							trnLng = LangPairList(2)
							m = 1
						End If
					End If
					If n + m = 2 Then Exit For
				End If
			Next i
			If n + m < 2 Then
				MsgBox MsgList(10),vbOkOnly+vbInformation,MsgList(9)
				TranTestFunc = True '防止按下按钮关闭对话框窗口
				Exit Function
			End If
			LangPair = srcLng & LngJoinStr & trnLng
			Temp = IIf(InStr("|zh-chs|zh-sg|zh-cht|zh-hk|zh-mo|zh-cn|zh-tw|ja|ko|","|" & PslLng & "|"),"Asia","")
			'获取翻译设置
			If CheckArray(tSelected) = True Then
				ProjectIDSrc = StrToLong(tSelected(18))
				mCheckSrc = StrToLong(tSelected(19))
				mPreStrRep = StrToLong(tSelected(20))
				mSplitTrn = StrToLong(tSelected(21))
				ProjectIDTrn = StrToLong(tSelected(22))
				mCheckTrn = StrToLong(tSelected(23))
				mAppStrRep = StrToLong(tSelected(24))
				If tSelected(17) = "1" Then
					CheckID = getCheckID(CheckDataList,PslLng,Temp)
				Else
					For i = LBound(CheckDataList) To UBound(CheckDataList)
						TempArray = ReSplit(CheckDataList(i),JoinStr)
						If TempArray(0) = tSelected(2) Then
							CheckID = i
							Exit For
						End If
					Next i
				End If
			Else
				ProjectIDSrc = 1
				mCheckSrc = 1
				mPreStrRep = 1
				mSplitTrn = 1
				ProjectIDTrn = 1
				mCheckTrn = 1
				mAppStrRep = 1
				CheckID = getCheckID(CheckDataList,PslLng,Temp)
			End If
			TempArray = ReSplit(CheckDataList(CheckID),JoinStr)
			CheckName = TempArray(0)
			AllCont = 1
			AccKey = 0
			EndChar = 0
			Acceler = 0
			'获取检查方案设置
			If mCheckSrc = 1 Then
				TempArray = ReSplit(ProjectDataList(ProjectIDSrc),JoinStr)
				SrcProjectName = TempArray(0)
			Else
				SrcProjectName = MsgList(5)
			End If
			If mCheckTrn = 1 Then
				TempArray = ReSplit(ProjectDataList(ProjectIDTrn),JoinStr)
				TempDataList = ReSplit(TempArray(1),LngJoinStr)
				ShowOriginalTran = StrToLong(TempDataList(19))
				ApplyCheckResult = StrToLong(TempDataList(20))
				TrnProjectName = TempArray(0)
			Else
				TrnProjectName = MsgList(5)
			End If
			'输出开始翻译消息
			If DlgValue("TranType") = 1 Or DlgItem$ = "HeaderButton" Then
				Temp = Replace(Replace(MsgList(2) & vbCrLf & MsgList(3),"%s",srcLng & " > " & trnLng),"%d",CheckName)
			Else
				Temp = Replace(Replace(MsgList(2) & vbCrLf & MsgList(3) & vbCrLf & MsgList(4),"%s",srcLng & " > " & trnLng),"%d",CheckName)
				Temp = Replace(Replace(Temp,"%p",SrcProjectName),"%n",TrnProjectName)
			End If
			DlgText "OutTextBox",Temp
			'获取翻译和字串处理
			If DlgItem$ = "HeaderButton" Then
				inText = getTranslate(xmlHttp,EngineDataList,EngineID,inText,LangPair,2)
			Else
				If mPreStrRep = 1 Then inText = ReplaceStr(CheckID,inText,0,0)
				If mSplitTrn = 0 Then
					If mCheckSrc = 1 Then inText = CheckHanding(CheckID,DlgText("InTextBox"),inText,ProjectIDSrc)
					If InStr(inText,"&") Then inText = Replace(inText,"&","")
					inText = getTranslate(xmlHttp,EngineDataList,EngineID,inText,LangPair,DlgValue("TranType"))
				Else
					inText = SplitTran(xmlHttp,EngineDataList,inText,LangPair,EngineID,CheckID,ProjectIDSrc,mCheckSrc,0,DlgValue("TranType"))
				End If
				If Trim(inText) <> "" And inText <> DlgText("InTextBox") And DlgValue("TranType") = 0 Then
					If mCheckTrn = 1 Then
						If ApplyCheckResult = 1 Then inText = CheckHanding(CheckID,DlgText("InTextBox"),inText,ProjectIDTrn)
					End If
					If mAppStrRep = 1 Then inText = ReplaceStr(CheckID,inText,2,0)
				End If
			End If
			Set xmlHttp = Nothing
			'输出消息和翻译
			If Trim(inText) = "" Then
				DlgText "OutTextBox",Temp & vbCrLf & MsgList(8) & vbCrLf & MsgList(6)
			ElseIf Trim(Replace(inText,vbCrLf,"")) = "" Then
				DlgText "OutTextBox",Temp & vbCrLf & MsgList(8) & vbCrLf & MsgList(7)
			Else
				DlgText "OutTextBox",Temp & vbCrLf & MsgList(8) & vbCrLf & inText
			End If
		End Select
	Case 3 ' 文本框或者组合框文本被更改
		If getMsgList(UIDataList,MsgList,"TranTestFunc",1) = False Then Exit Function
		Select Case DlgItem$
		Case "LineNumBox"
			Temp = DlgText("TrnListBox")
			Set TrnList = trn
			For i = 1 To trn.Project.TransLists.Count
				Set TrnList = trn.Project.TransLists(i)
				If Temp = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) Then
					Exit For
				End If
			Next i
			m = StrToLong(DlgText("LineNumBox"))
			If TrnList.StringCount < m Then m = TrnList.StringCount
			n = 0
			ReDim TempList(m - 1) As String
			For i = 1 To TrnList.StringCount
				Set TransString = TrnList.String(i)
				If TransString.Text <> "" Then
					If DlgValue("StrType") = 0 Then
						TempList(n) = TransString.SourceText
					Else
						TempList(n) = TransString.Text
					End If
					n = n + 1
					If n = m Then Exit For
				End If
			Next i
			ReDim Preserve TempList(IIf(n = 0,0,n - 1)) As String
			DlgText "InTextBox",Join$(TempList,vbCrLf)
			If DlgText("InTextBox") <> "" Then
				DlgText "ClearButton",MsgList(0)
				DlgEnable "TestButton",True
				DlgEnable "HeaderButton",True
			Else
    			DlgText "ClearButton",MsgList(1)
    			DlgEnable "TestButton",False
    			DlgEnable "HeaderButton",False
    		End If
		Case "InTextBox"
			If DlgText("InTextBox") <> "" Then
				DlgText "ClearButton",MsgList(0)
				DlgEnable "TestButton",True
				DlgEnable "HeaderButton",True
			Else
    			DlgText "ClearButton",MsgList(1)
    			DlgEnable "TestButton",False
    			DlgEnable "HeaderButton",False
    		End If
    	End Select
	End Select
End Function


'测试检查程序
Sub CheckTest(ByVal CheckID As Long,ByVal ProjectID As Long)
	Dim MsgList() As String
	If getMsgList(UIDataList,MsgList,"CheckTest",1) = False Then Exit Sub
	Begin Dialog UserDialog 660,518,MsgList(0),.CheckTestFunc ' %GRID:10,7,1,1
		GroupBox 10,42,640,126,"",.GroupBox
		Text 10,7,640,28,MsgList(1),.MainText
		Text 30,59,90,14,MsgList(2),.SetNameText
		DropListBox 130,56,300,21,CheckList(),.CheckListBox
		CheckBox 450,59,180,14,MsgList(16),.RepStrBox
		Text 30,80,90,14,MsgList(3),.ProjectText
		DropListBox 130,77,300,21,ProjectList(),.ProjectListBox
		Text 30,101,90,14,MsgList(4),.TrnListText
		DropListBox 130,98,300,21,MsgList(),.TrnListBox
		Text 450,87,120,14,MsgList(5),.LineNumText
		TextBox 580,84,50,18,.LineNumBox
		Text 30,122,90,14,MsgList(6),.SpecifyText
		TextBox 130,119,300,21,.SpecifyTextBox
		Text 450,117,180,24,MsgList(7),.TypeTipText
		Text 30,147,90,14,MsgList(8),.TypeText
		CheckBox 130,147,110,14,MsgList(9),.AllCheckBox
		CheckBox 250,147,120,14,MsgList(10),.AcckeyCheckBox
		CheckBox 380,147,120,14,MsgList(11),.EndSharCheckBox
		CheckBox 510,147,120,14,MsgList(12),.ShortCheckBox
		TextBox 10,182,640,294,.InTextBox,1
		PushButton 10,490,100,21,MsgList(15),.HelpButton
		PushButton 120,490,100,21,MsgList(13),.TestButton
		PushButton 230,490,100,21,MsgList(14),.ClearButton
		CancelButton 550,490,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CheckListBox = CheckID
	dlg.ProjectListBox = ProjectID
	If Dialog(dlg) = 0 Then Exit Sub
End Sub


'测试对话框函数
Private Function CheckTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim MsgList() As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		Dim i As Long,j As Long
		ReDim MsgList(trn.Project.TransLists.Count - 1)
		For i = 1 To trn.Project.TransLists.Count
			MsgList(i - 1) = trn.Project.TransLists(i).Title & " - " & _
							PSL.GetLangCode(trn.Project.TransLists(i).Language.LangID,pslCodeText)
		Next i
		DlgListBoxArray "TrnListBox",MsgList()
		DlgText "TrnListBox",trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
		DlgText "LineNumBox","10"
		DlgValue "AllCheckBox",1
    	If DlgText("InTextBox") = "" Then DlgEnable "ClearButton",False
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		CheckTestFunc = True '防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			CheckTestFunc = False '按下按钮关闭对话框窗口
		Case "AllCheckBox"
			If DlgValue("AllCheckBox") = 1 Then
				DlgValue "AcckeyCheckBox",0
				DlgValue "EndSharCheckBox",0
				DlgValue "ShortCheckBox",0
			ElseIf DlgValue("AcckeyCheckBox") + DlgValue("EndSharCheckBox") + DlgValue("ShortCheckBox") = 0 Then
				DlgValue "AllCheckBox",1
			End If
		Case "AcckeyCheckBox", "EndSharCheckBox", "ShortCheckBox"
			If DlgValue("AcckeyCheckBox") = 1 Or DlgValue("EndSharCheckBox") = 1 Or DlgValue("ShortCheckBox") = 1 Then
				DlgValue "AllCheckBox",0
			Else
				DlgValue "AllCheckBox",1
			End If
		Case "TestButton"
			If getMsgList(UIDataList,MsgList,"CheckTest",1) = False Then Exit Function
			AllCont = DlgValue("AllCheckBox")
			AccKey = DlgValue("AcckeyCheckBox")
			EndChar = DlgValue("EndSharCheckBox")
			Acceler = DlgValue("ShortCheckBox")
			DlgText "InTextBox",MsgList(17)
			DlgText "InTextBox",CheckStrings(DlgValue("CheckListBox"),DlgValue("ProjectListBox"),DlgText("TrnListBox"), _
					StrToLong(DlgText("LineNumBox")),DlgValue("RepStrBox"),DlgText("SpecifyTextBox"))
			DlgEnable "ClearButton",IIf(DlgText("InTextBox") = "",False,True)
		Case "ClearButton"
			DlgText "InTextBox",""
			DlgEnable "ClearButton",False
		Case "HelpButton"
			Call Help("CheckTestHelp")
		End Select
	Case 3 ' 文本框或者组合框文本被更改
		If DlgItem$ = "InTextBox" Then
			DlgEnable "ClearButton",IIf(DlgText("InTextBox") = "",False,True)
    	End If
	End Select
End Function


'处理符合条件的字串列表中的字串
Function CheckStrings(ByVal cID As Long,ByVal pID As Long,ByVal ListDec As String,ByVal LineNum As Long,ByVal rStr As Long,ByVal sText As String) As String
	Dim i As Long,j As Long,n As Long,srcString As String,trnString As String
	Dim TrnList As PslTransList,Stemp As Boolean
	Dim srcFindNum As Long,trnFindNum As Long,MsgList() As String

	If getMsgList(UIDataList,MsgList,"CheckStrings",1) = False Then Exit Function

	'获取选定的翻译列表
	Set TrnList = trn
	For i = 1 To trn.Project.TransLists.Count
		Set TrnList = trn.Project.TransLists(i)
		If ListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) Then Exit For
	Next i

	If TrnList.SourceList.LastChange > TrnList.LastUpdate Then TrnList.Update
	If TrnList.StringCount < LineNum Then LineNum = TrnList.StringCount
	ReDim TempList(LineNum) As String
	For i = 1 To TrnList.StringCount
		'获取原文和翻译字串
		Set TransString = TrnList.String(i)
		If TransString.Text <> "" Then
			srcString = TransString.SourceText
			trnString = TransString.Text

			'转换转义符
			Stemp = False
			srcString = Convert(srcString)
			If srcString <> TransString.SourceText Then
				trnString = Convert(trnString)
				Stemp = True
			End If

			'开始处理字串
			If sText <> "" Then
				FindStrArr = ReSplit(sText,";",-1)
				For j = 0 To UBound(FindStrArr)
					FindStr = FindStrArr(j)
					If Left(FindStr,1) <> "*" And Right(FindStr,1) <> "*" Then
						FindStr = "*" & FindStr & "*"
					End If
					If CheckKeyCode(srcString,FindStr) <> 0 Then
						trnString = CheckHanding(cID,srcString,trnString,pID)
						If trnString <> "" Then Exit For
					ElseIf CheckKeyCode(trnString,FindStr) <> 0 Then
						trnString = CheckHanding(cID,srcString,trnString,pID)
						If trnString <> "" Then Exit For
					End If
				Next j
			Else
				trnString = CheckHanding(cID,srcString,trnString,pID)
			End If
			If rStr = 1 Then trnString = ReplaceStr(cID,trnString,2,1)
			If Stemp = True Then trnString = ReConvert(trnString)

			'调用消息输出
			If trnString <> "" And trnString <> TransString.Text Then
				TempList(n) = MsgList(0) & srcString & vbCrLf & _
							MsgList(1) & TransString.Text & vbCrLf & _
							MsgList(2) & trnString & vbCrLf & _
							MsgList(3) & ReplaceMassage(cID,pID) & vbCrLf & _
							MsgList(4) & MsgList(4) & MsgList(4)
				n = n + 1
				If n = LineNum Then Exit For
			End If
		End If
	Next i
	If n > 0 Then
		CheckStrings = MsgList(5) & vbCrLf & MsgList(4) & MsgList(4) & MsgList(4) & vbCrLf & Join$(TempList,vbCrLf)
	Else
		If sText = "" Then
			CheckStrings = MsgList(6)
		ElseIf Find = True Then
			CheckStrings = MsgList(7)
		Else
			CheckStrings = MsgList(8)
		End If
	End If
End Function


'添加工具数据
Sub AddTools(ToolData() As TOOLS_PROPERTIE,ByVal CmdName As String,ByVal CmdPath As String,ByVal Argument As String)
	Dim i As Long,n As Long,FindName As String,Stemp As Boolean
	If CmdName = "" Or CmdPath = "" Then Exit Sub
	n = UBound(ToolData)
	FindName = LCase$(CmdName)
	For i = LBound(ToolData) To n
		If LCase$(ToolData(i).sName) = FindName Then
			Stemp = True
			Exit For
		End If
	Next i
	If Stemp = False Then
		ReDim Preserve ToolData(n + 1) As TOOLS_PROPERTIE
		ToolData(n + 1).sName = CmdName
		ToolData(n + 1).FilePath = CmdPath
		ToolData(n + 1).Argument = Argument
	End If
End Sub


'打开文本文件
Function OpenFile(ByVal File As String,FileDataList() As String,ByVal x As Long,RunStemp As Boolean) As Boolean
	Dim i As Long,ExePath As String,ExeName As String,Argument As String,ExtName As String
	Dim TempArray() As String,MsgList() As String,WshShell As Object

	OpenFile = False
	If getMsgList(UIDataList,MsgList,"OpenFile",1) = False Then Exit Function

	If x > 0 Then
		On Error Resume Next
		Set WshShell = CreateObject("WScript.Shell")
		If WshShell Is Nothing Then
			Err.Source = "WScript.Shell"
			Call sysErrorMassage(Err,2)
			Exit Function
		End If
		On Error GoTo 0
	End If

	Select Case x
	Case 0
		If EditFile(File,FileDataList,RunStemp) = True Then OpenFile = True
	Case 1
		ExePath = Environ("SystemRoot") & "\system32\notepad.exe"
		If Dir$(ExePath) = "" Then
			ExePath = Environ("SystemRoot") & "\notepad.exe"
		End If
		If Dir(ExePath) = "" Then
			MsgBox MsgList(1),vbOkOnly+vbInformation,MsgList(0)
		Else
			If WshShell.Run("""" & ExePath & """ " & """" & File & """",1,RunStemp) <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				OpenFile = True
			End If
		End If
	Case 2
		i = InStrRev(File,".")
		If i > 0 Then ExtName = Mid$(File,i)
		On Error Resume Next
		ExtName = WshShell.RegRead("HKCR\" & ExtName & "\")
		If ExtName <> "" Then
			ExePath = WshShell.RegRead("HKCR\" & ExtName & "\shell\edit\command\")
			If ExePath = "" Then
				ExePath = WshShell.RegRead("HKCR\" & ExtName & "\shell\open\command\")
			End If
			If ExePath = "" Then
				ExePath = WshShell.RegRead("HKCR\" & ExtName & "\shell\preview\command\")
			End If
		End If
		On Error GoTo 0
		If ExePath <> "" Then
			i = InStr(ExePath,".")
			If i > 0 Then Argument = Trim$(Mid$(ExePath,InStr(i,ExePath," ")))
			ExePath = Left$(ExePath,Len(ExePath) - Len(Argument))
			TempArray = ReSplit(ExePath,"%")
			If UBound(TempArray) >= 2 Then
				ExePath = Replace$(ExePath,"%" & TempArray(1) & "%",Environ(TempArray(1)),,1)
			End If
			ExePath = RemoveBackslash(ExePath,"""","""",1)
			ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)

			If ExePath <> "" Then
				If InStr(ExePath,"\") = 0 Then
					If Dir$(Environ("SystemRoot") & "\system32\" & ExePath) <> "" Then
						ExePath = Environ("SystemRoot") & "\system32\" & ExePath
					ElseIf Dir$(Environ("SystemRoot") & "\" & ExePath) <> "" Then
						ExePath = Environ("SystemRoot") & "\" & ExePath
					End If
				End If
			End If

			If Argument <> "" Then
				If InStr(Argument,"%1") Then
					File = Replace$(Argument,"%1",File)
				ElseIf InStr(Argument,"%L") Then
					File = Replace$(Argument,"%L",File)
				Else
					File = Argument & " " & """" & File & """"
				End If
			Else
				File = """" & File & """"
			End If
		End If
		If ExePath = "" Then
			MsgBox MsgList(2),vbOkOnly+vbInformation,MsgList(0)
		ElseIf Dir$(ExePath) <> "" Then
			If WshShell.Run("""" & ExePath & """ " & File,1,RunStemp) <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)
				For i = 0 To UBound(Tools)
					If InStr(LCase$(Tools(i).FilePath),LCase$(ExeName)) Then
						File = ""
						Exit For
					End If
				Next i
				If File <> "" Then Call AddTools(Tools,ExeName,ExePath,Argument)
				OpenFile = True
			End If
		Else
			MsgBox Replace$(Replace$(Replace$(MsgList(6),"%s!1!",ExeName),"%s!2!",ExePath), _
					"%s!3!",Argument) & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
		End If
  	Case 3
		If CommandInput(ExePath,Argument) = True Then
			TempArray = ReSplit(ExePath,"%")
			If UBound(TempArray) = 2 Then
				ExePath = Replace$(ExePath,"%" & TempArray(1) & "%",Environ(TempArray(1)))
			End If
			ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)

			If Argument <> "" Then
				If InStr(Argument,"%1") Then
					File = Replace$(Argument,"%1",File)
				ElseIf InStr(Argument,"%L") Then
					File = Replace$(Argument,"%L",File)
				Else
					File = Argument & " " & """" & File & """"
				End If
			Else
				File = """" & File & """"
			End If

			If Dir$(ExePath) <> "" Then
				If WshShell.Run("""" & ExePath & """ " & File,1,RunStemp) <> 0 Then
					MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(0)
				Else
					ExeName = Mid$(ExePath,InStrRev(ExePath,"\") + 1)
					For i = 0 To UBound(Tools)
						If InStr(LCase$(Tools(i).FilePath),LCase$(ExeName)) Then
							File = ""
							Exit For
						End If
					Next i
					If File <> "" Then Call AddTools(Tools,ExeName,ExePath,Argument)
					OpenFile = True
				End If
			Else
				MsgBox ExeName & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
			End If
		End If
	Case Is > 3
		ExeName = Tools(x).sName
		ExePath = Tools(x).FilePath
		Argument = Tools(x).Argument
		If Argument <> "" Then
			If InStr(Argument,"%1") Then
				File = Replace$(Argument,"%1",File)
			ElseIf InStr(Argument,"%L") Then
				File = Replace$(Argument,"%L",File)
			Else
				File = Argument & " " & """" & File & """"
			End If
		Else
			File = """" & File & """"
		End If
		If ExePath <> "" Then
			If Dir$(ExePath) <> "" Then
				If WshShell.Run("""" & ExePath & """ " & File,1,RunStemp) <> 0 Then
					MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
				Else
					OpenFile = True
				End If
			Else
				MsgBox ExeName & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
			End If
		End If
	End Select
	Set WshShell = Nothing
End Function


'编辑文本文件
'Mode = True 编辑模式，如果打开文件成功返回字符编码和 True
'Mode = False 查看和确认字符编码模式，如果打开文件成功并按 [确定] 按钮返回字符编码和 True
Function EditFile(ByVal File As String,FileDataList() As String,ByVal Mode As Boolean) As Boolean
	Dim MsgList() As String,FileDataListBak() As String
	If getMsgList(UIDataList,MsgList,"EditFile",1) = False Then Exit Function
	'Dim objStream As Object
	'Set objStream = CreateObject("Adodb.Stream")
	'If objStream Is Nothing Then CodeList = getCodePageList(0,0)
	'If Not objStream Is Nothing Then CodeList = getCodePageList(0,49)
	'Set objStream = Nothing
	FileDataListBak = FileDataList
	Begin Dialog UserDialog 1020,595,IIf(Mode = True,MsgList(0),MsgList(1)) & " - " & File,.EditFileDlgFunc ' %GRID:10,7,1,1
		CheckBox 0,3,14,14,"",.OptionBox
		TextBox 0,0,0,21,.SuppValueBox
		TextBox 0,3,110,21,.TextLengthBox
		Text 10,7,920,14,File,.FilePath,2
		Text 10,7,90,14,MsgList(2),.FindText
		DropListBox 110,3,270,21,MsgList(),.FindTextBox,1
		PushButton 420,3,90,21,MsgList(4),.FindButton
		PushButton 520,3,90,21,MsgList(5),.FilterButton
		PushButton 520,3,90,21,MsgList(6),.CloseFilterButton
		PushButton 380,3,30,21,MsgList(3),.RegExpTipButton
		Text 630,7,90,14,MsgList(7),.CodeText
		DropListBox 730,3,280,21,MsgList(),.CodeNameList
		TextBox 0,28,1020,532,.InTextBox,1
		PushButton 20,567,90,21,MsgList(8),.HelpButton
		PushButton 120,567,90,21,"",.ReadButton
		PushButton 220,567,90,21,MsgList(9),.PreviousButton
		PushButton 320,567,90,21,MsgList(10),.NextButton
		PushButton 790,567,100,21,MsgList(11),.SaveButton
		PushButton 900,567,90,21,MsgList(12),.ExitButton
		OKButton 790,567,100,21,.OKButton
		CancelButton 900,567,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Mode = False Then dlg.OptionBox = 1
	If Dialog(dlg) = 0 Then
		FileDataList = FileDataListBak
		Erase AllStrList,UseStrList
		Exit Function
	End If
	EditFile = True
	Erase AllStrList,UseStrList
End Function


'编辑对话框函数
Private Function EditFileDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,n As Long
	Dim MsgList() As String,Temp As String,Code As String
	Dim TempArray() As String,TempList() As String,pt As POINTAPI

	Select Case Action%
	Case 1
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		DlgText "SuppValueBox",CStr$(SuppValue)
		DlgVisible "SuppValueBox",False
		DlgVisible "FilePath",False
		DlgVisible "TextLengthBox",False
		DlgVisible "OptionBox",False
		If DlgValue("OptionBox") = 0 Then
			DlgVisible "OKButton",False
			DlgVisible "CancelButton",False
		Else
			DlgVisible "SaveButton",False
			DlgVisible "ExitButton",False
		End If
		DlgVisible "CloseFilterButton",False
		GetHistory(TempList,"FindStrings","EditFileDlg")
		DlgListBoxArray "FindTextBox",TempList()
		DlgText "FindTextBox",TempList(0)
		Temp = DlgText("FilePath")
		For i = LBound(FileDataList) To UBound(FileDataList)
			TempList = ReSplit(FileDataList(i),JoinStr,-1)
			If TempList(0) = Temp Then
				Code = TempList(1)
				If Code = "" Then
					Code = CheckCode(Temp)
					TempList(1) = Code
					FileDataList(i) = Join$(TempList,JoinStr)
				End If
				j = i
				Exit For
			End If
		Next i
		ReDim TempList(UBound(CodeList)) As String
		For i = LBound(CodeList) To UBound(CodeList)
			TempList(i) = CodeList(i).sName
			If CodeList(i).CharSet = Code Then n = i
		Next i
		DlgListBoxArray "CodeNameList",TempList()
		DlgValue "CodeNameList",n
		Temp = ReadFile(Temp,Code)
		'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
		i = Len(Temp)
		If i > 25000 Then
			n = GetDlgItem(SuppValue,DlgControlId("InTextBox"))
			i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,False)
			DlgText "TextLengthBox",CStr$(i)
		End If
		DlgText "InTextBox",Temp
		If Temp <> "" Then
			DlgText "ReadButton",MsgList(9)
    	Else
    		DlgText "ReadButton",MsgList(8)
    		DlgEnable "FindButton",False
    		DlgEnable "FilterButton",False
    		DlgEnable "CloseFilterButton",False
    		DlgEnable "SaveButton",False
    	End If
    	If UBound(FileDataList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		ElseIf j = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",True
		ElseIf j = UBound(FileDataList) Then
			DlgEnable "PreviousButton",True
			DlgEnable "NextButton",False
    	End If
    	'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			j = CreateFont(0,LFList(0))
			If j = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,j,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		EditFileDlgFunc = True '防止按下按钮关闭对话框窗口
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		Select Case DlgItem$
		Case "HelpButton"
			Call Help("EditFileHelp")
			Exit Function
		Case "OKButton", "CancelButton"
			EditFileDlgFunc = False
			Exit Function
		Case "ExitButton"
			If DlgText("InTextBox") = ReadFile(DlgText("FilePath"),CodeList(DlgValue("CodeNameList")).CharSet) Then
				EditFileDlgFunc = False
				Exit Function
			End If
			Select Case MsgBox(MsgList(1),vbYesNoCancel+vbInformation,MsgList(0))
			Case vbYes
				Temp = DlgText("FilePath")
				If Dir$(Temp) <> "" Then SetAttr Temp,vbNormal
				If WriteToFile(Temp,DlgText("InTextBox"),CodeList(DlgValue("CodeNameList")).CharSet) = True Then
					MsgBox(MsgList(5),vbOkOnly+vbInformation,MsgList(0))
					EditFileDlgFunc = False
					Exit Function
				Else
					MsgBox(MsgList(6),vbOkOnly+vbInformation,MsgList(0))
				End If
			Case vbNo
				EditFileDlgFunc = False
				Exit Function
			End Select
		Case "SaveButton"
			If DlgText("InTextBox") = "" Then Exit Function
			Temp = DlgText("FilePath")
			If Dir$(Temp) <> "" Then SetAttr Temp,vbNormal
			If WriteToFile(Temp,DlgText("InTextBox"),CodeList(DlgValue("CodeNameList")).CharSet) = True Then
				MsgBox(MsgList(5),vbOkOnly+vbInformation,MsgList(0))
			Else
				MsgBox(MsgList(6),vbOkOnly+vbInformation,MsgList(0))
			End If
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "CodeNameList"
			Code = CodeList(DlgValue("CodeNameList")).CharSet
			If Code = "_autodetect_all" Or Code = "_autodetect" Or Code = "_autodetect_kr" Then
				Code = CheckCode(DlgText("FilePath"))
				For i = LBound(CodeList) To UBound(CodeList)
					If CodeList(i).CharSet = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
			End If
			Temp = ReadFile(DlgText("FilePath"),Code)
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = Len(Temp)
			If i > 25000 Then
				n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
				i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
				DlgText "TextLengthBox",CStr$(i)
			End If
			DlgText "InTextBox",Temp
			If DlgText("InTextBox") <> "" Then
				Temp = DlgText("FilePath")
				For i = LBound(FileDataList) To UBound(FileDataList)
					TempList = ReSplit(FileDataList(i),JoinStr)
					If TempList(0) = Temp Then
 						TempList(1) = Code
						FileDataList(i) = Join$(TempList,JoinStr)
 					End If
				Next i
			End If
			Erase AllStrList,UseStrList
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "ReadButton"
			If DlgText("ReadButton") = MsgList(8) Then
				Code = CodeList(DlgValue("CodeNameList")).CharSet
				Temp = ReadFile(DlgText("FilePath"),Code)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				i = Len(Temp)
				If i > 25000 Then
					n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
					i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
					DlgText "TextLengthBox",CStr$(i)
				End If
				DlgText "InTextBox",Temp
				If Temp <> "" Then DlgText "ReadButton",MsgList(9)
			Else
				DlgText "InTextBox",""
				DlgText "ReadButton",MsgList(8)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
				i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),0,True)
				DlgText "TextLengthBox",CStr$(i)
			End If
			Erase AllStrList,UseStrList
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "FindButton"
			If DlgText("FindTextBox") = "" Then
				MsgBox(MsgList(11),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'添加查找内容
			GetHistory(TempList,"FindStrings","EditFileDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","EditFileDlg")
				DlgListBoxArray "FindTextBox",TempList()
				DlgText "FindTextBox",TempList(0)
			End If
			'DlgFocus("InTextBox")  '设置焦点到文本框，2016 版会闪烁，使得光标位置移到最前面
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
			SendMessageLNG(n,WM_SETFOCUS,0,0)  '设置焦点到文本框
			Select Case FindCurPos(n,DlgText("FindTextBox"),False,Temp)
			Case 0
				MsgBox(MsgList(4),vbOkOnly+vbInformation,MsgList(0))
			Case -1
				MsgBox(Replace$(MsgList(12),"%s",Temp),vbOkOnly+vbInformation,MsgList(0))
			Case -2
				MsgBox(MsgList(13),vbOkOnly+vbInformation,MsgList(0))
			End Select
			Exit Function
		Case "FilterButton"
			If DlgText("FindTextBox") = "" Then
				MsgBox(MsgList(11),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'添加查找内容
			GetHistory(TempList,"FindStrings","EditFileDlg")
			If InsertArray(TempList,DlgText("FindTextBox"),0,True) = True Then
				WriteHistory(TempList,"FindStrings","EditFileDlg")
				DlgListBoxArray "FindTextBox",TempList()
				DlgText "FindTextBox",TempList(0)
			End If
			'检测查找内容的查找方式
			Temp = DlgText("FindTextBox")
			j = GetFindMode(Temp)
			If j = 1 Then Temp = "*" & Temp & "*"
			AllStrList = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
			ReDim UseStrList(UBound(AllStrList)) As String
			n = 0
			For i = 0 To UBound(AllStrList)
				Select Case FilterStr(AllStrList(i),Temp,j)
				Case Is < 0
					MsgBox(MsgList(13),vbOkOnly+vbInformation,MsgList(0))
					Exit Function
				Case Is > 0
					UseStrList(n) = "【" & CStr$(i + 1) & MsgList(10) & "】" & AllStrList(i)
					n = n + 1
				End Select
			Next i
			If n > 0 Then
				ReDim Preserve UseStrList(n - 1) As String
				DlgText "InTextBox",Join$(UseStrList,vbCrLf)
				DlgVisible "FilterButton",False
				DlgVisible "CloseFilterButton",True
			Else
				Erase AllStrList,UseStrList
				MsgBox(MsgList(4),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
		Case "CloseFilterButton"
			If DlgText("InTextBox") = "" Then
				MsgBox(MsgList(3),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			If DlgText("InTextBox") <> Join$(UseStrList,vbCrLf) Then
				If MsgBox(MsgList(2),vbYesNo+vbInformation,MsgList(0)) = vbYes Then
					TempArray = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
					If UBound(UseStrList) = UBound(TempArray) Then
						Temp = "^【[0-9]+" & MsgList(10) & "】"
						For i = 0 To UBound(TempArray)
							If CheckStrRegExp(TempArray(i),Temp,0,2) = True Then
								TempList = ReSplit(TempArray(i),MsgList(10) & "】",2)
								AllStrList(CLng(Mid$(TempList(0),2)) - 1) = TempList(1)
							End If
						Next i
					Else
						MsgBox(MsgList(3),vbOkOnly+vbInformation,MsgList(0))
						Exit Function
					End If
				End If
			End If
			Temp = Join(AllStrList,vbCrLf)
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = Len(Temp)
			If i > 25000 Then
				n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
				i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
				DlgText "TextLengthBox",CStr$(i)
			End If
			DlgText "InTextBox",Temp
			Erase AllStrList,UseStrList
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "PreviousButton","NextButton"
			Temp = DlgText("FilePath")
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = ReSplit(FileDataList(i),JoinStr)
				If TempList(0) = Temp Then
					j = i
					Exit For
				End If
			Next i
			If DlgItem$ = "PreviousButton" Then
				If j <> 0 Then j = j - 1
			Else
				If j < UBound(FileDataList) Then j = j + 1
			End If
			If i <> j Then
				TempList = ReSplit(FileDataList(j),JoinStr)
				DlgText "FilePath",TempList(0)
				DlgText -1,Left$(DlgText(-1),InStr(DlgText(-1),"-") + 1) & TempList(0)
				Code = TempList(1)
				If Code = "" Then Code = CheckCode(TempList(0))
				For i = LBound(CodeList) To UBound(CodeList)
					If CodeList(i).CharSet = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				Temp = ReadFile(TempList(0),Code)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				i = Len(Temp)
				If i > 25000 Then
					n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
					i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
					DlgText "TextLengthBox",CStr$(i)
				End If
				DlgText "InTextBox",Temp
				Erase AllStrList,UseStrList
			End If
			If UBound(FileDataList) = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",False
			ElseIf j = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf j = UBound(FileDataList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",True
			End If
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		Case "RegExpTipButton"
			If getMsgList(UIDataList,MsgList,"FindSetDlgFunc",1) = False Then Exit Function
			i = ShowPopupMenu(MsgList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(MsgList) Then
				Call Help("RegExpRuleHelp")
				Exit Function
			End If
			If DlgText("FindTextBox") = "" Then
				DlgText "FindTextBox",Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1)
			Else
				DlgFocus("FindTextBox")  '设置焦点到文本框
				j = GetFocus()
				DlgText "FindTextBox",InsertStr(j,DlgText("FindTextBox"), _
						Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1),pt)
				Call SetCurPos(j,pt,0)
			End If
			Exit Function
		End Select

    	If DlgText("InTextBox") <> "" Then
   			DlgText "ReadButton",MsgList(9)
   			DlgEnable "FindButton",True
   			DlgEnable "FilterButton",True
			DlgEnable "CloseFilterButton",True
			If DlgVisible("FilterButton") = True Then
				DlgEnable "SaveButton",True
				DlgEnable "ExitButton",True
				DlgEnable "CancelButton",True
			ElseIf DlgValue("OptionBox") = 0 Then
				DlgEnable "SaveButton",False
				DlgEnable "ExitButton",False
				DlgEnable "CancelButton",False
			End If
		Else
			DlgText "ReadButton",MsgList(8)
			DlgEnable "FindButton",False
			DlgEnable "FilterButton",False
			DlgEnable "CloseFilterButton",False
			DlgEnable "SaveButton",False
			DlgEnable "ExitButton",True
			DlgEnable "CancelButton",True
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		Select Case DlgItem$
		Case "InTextBox"
			If DlgText("InTextBox") <> "" Then
				DlgText "ReadButton",MsgList(9)
				DlgEnable "FindButton",True
				DlgEnable "FilterButton",True
				DlgEnable "CloseFilterButton",True
				If DlgVisible("FilterButton") = True Then
					DlgEnable "SaveButton",True
					DlgEnable "ExitButton",True
					DlgEnable "CancelButton",True
				ElseIf DlgValue("OptionBox") = 0 Then
					DlgEnable "SaveButton",False
					DlgEnable "ExitButton",False
					DlgEnable "CancelButton",False
				End If
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				i = Len(DlgText("InTextBox"))
				If i > 25000 Then
					n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
					i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
					DlgText "TextLengthBox",CStr$(i)
				End If
			Else
				DlgText "ReadButton",MsgList(8)
				DlgEnable "FindButton",False
				DlgEnable "FilterButton",False
				DlgEnable "CloseFilterButton",False
				DlgEnable "SaveButton",False
				DlgEnable "ExitButton",True
				DlgEnable "CancelButton",True
			End If
		End Select
	Case 4 ' 焦点被更改
		Select Case DlgItem$
		Case "InTextBox"
			'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
			i = Len(Clipboard)
			If i > 25000 Then
				n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
				i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),StrToLong(DlgText("TextLengthBox")) + i,False)
				DlgText "TextLengthBox",CStr$(i)
			End If
		End Select
	Case 6 ' 函数快捷键
		If getMsgList(UIDataList,MsgList,"EditFileDlgFunc",1) = False Then Exit Function
		Select Case SuppValue
		Case 1
			Call Help("EditFileHelp")
		Case 2
			If getMsgList(UIDataList,MsgList,"FindSetDlgFunc",1) = False Then Exit Function
			i = ShowPopupMenu(MsgList,vbPopupUseRightButton)
			If i < 0 Then Exit Function
			If i = UBound(MsgList) Then
				Call Help("RegExpRuleHelp")
				Exit Function
			End If
			If DlgText("FindTextBox") = "" Then
				DlgText "FindTextBox",Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1)
			Else
				DlgFocus("FindTextBox")  '设置焦点到文本框
				j = GetFocus()
				DlgText "FindTextBox",InsertStr(j,DlgText("FindTextBox"), _
						Mid$(MsgList(i),InStrRev(MsgList(i),vbTab) + 1),pt)
				Call SetCurPos(j,pt,0)
			End If
		Case 3
			If DlgText("FindTextBox") = "" Then
				MsgBox(MsgList(11),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			'DlgFocus("InTextBox")  '设置焦点到文本框，2016 版会闪烁，使得光标位置移到最前面
			n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
			SendMessageLNG(n,WM_SETFOCUS,0,0)  '设置焦点到文本框
			Select Case FindCurPos(n,DlgText("FindTextBox"),False,Temp)
			Case 0
				MsgBox(MsgList(4),vbOkOnly+vbInformation,MsgList(0))
			Case -1
				MsgBox(Replace$(MsgList(12),"%s",Temp),vbOkOnly+vbInformation,MsgList(0))
			Case -2
				MsgBox(MsgList(13),vbOkOnly+vbInformation,MsgList(0))
			End Select
		Case 4
			If DlgVisible("FilterButton") = True Then
				If DlgText("FindTextBox") = "" Then
					MsgBox(MsgList(11),vbOkOnly+vbInformation,MsgList(0))
					Exit Function
				End If
				'检测查找内容的查找方式
				Temp = DlgText("FindTextBox")
				j = GetFindMode(Temp)
				If j = 1 Then Temp = "*" & Temp & "*"
				AllStrList = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
				ReDim UseStrList(UBound(AllStrList)) As String
				n = 0
				For i = 0 To UBound(AllStrList)
					Select Case FilterStr(AllStrList(i),Temp,j)
					Case Is < 0
						MsgBox(MsgList(13),vbOkOnly+vbInformation,MsgList(0))
						Exit Function
					Case Is > 0
						UseStrList(n) = "【" & CStr$(i + 1) & MsgList(10) & "】" & AllStrList(i)
						n = n + 1
					End Select
				Next i
				If n > 0 Then
					ReDim Preserve UseStrList(n - 1) As String
					DlgText "InTextBox",Join$(UseStrList,vbCrLf)
					DlgVisible "FilterButton",False
					DlgVisible "CloseFilterButton",True
				Else
					Erase AllStrList,UseStrList
					MsgBox(MsgList(4),vbOkOnly+vbInformation,MsgList(0))
					Exit Function
				End If
			Else
				If DlgText("InTextBox") = "" Then
					MsgBox(MsgList(3),vbOkOnly+vbInformation,MsgList(0))
					Exit Function
				End If
				If DlgText("InTextBox") <> Join$(UseStrList,vbCrLf) Then
					If MsgBox(MsgList(2),vbYesNo+vbInformation,MsgList(0)) = vbYes Then
						TempArray = ReSplit(DlgText("InTextBox"),vbCrLf,-1)
						If UBound(UseStrList) = UBound(TempArray) Then
							Temp = "^【[0-9]+" & MsgList(10) & "】"
							For i = 0 To UBound(TempArray)
								If CheckStrRegExp(TempArray(i),Temp,0,2) = True Then
									TempList = ReSplit(TempArray(i),MsgList(10) & "】",2)
									AllStrList(CLng(Mid$(TempList(0),2)) - 1) = TempList(1)
								End If
							Next i
						Else
							MsgBox(MsgList(3),vbOkOnly+vbInformation,MsgList(0))
							Exit Function
						End If
					End If
				End If
				Temp = Join(AllStrList,vbCrLf)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				i = Len(Temp)
				If i > 25000 Then
					n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
					i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
					DlgText "TextLengthBox",CStr$(i)
				End If
				DlgText "InTextBox",Temp
				Erase AllStrList,UseStrList
				DlgVisible "FilterButton",True
				DlgVisible "CloseFilterButton",False
			End If
			If DlgVisible("FilterButton") = True Then
				DlgEnable "SaveButton",True
				DlgEnable "ExitButton",True
				DlgEnable "CancelButton",True
			ElseIf DlgValue("OptionBox") = 0 Then
				DlgEnable "SaveButton",False
				DlgEnable "ExitButton",False
				DlgEnable "CancelButton",False
			End If
		Case 5, 6
			Temp = DlgText("FilePath")
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = ReSplit(FileDataList(i),JoinStr)
				If TempList(0) = Temp Then
					j = i
					Exit For
				End If
			Next i
			If SuppValue = 5 Then
				If j <> 0 Then j = j - 1
			Else
				If j < UBound(FileDataList) Then j = j + 1
			End If
			If i <> j Then
				TempList = ReSplit(FileDataList(j),JoinStr)
				DlgText "FilePath",TempList(0)
				DlgText -1,Left$(DlgText(-1),InStr(DlgText(-1),"-") + 1) & TempList(0)
				Code = TempList(1)
				If Code = "" Then Code = CheckCode(TempList(0))
				For i = LBound(CodeList) To UBound(CodeList)
					If CodeList(i).CharSet = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				Temp = ReadFile(TempList(0),Code)
				'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
				i = Len(Temp)
				If i > 25000 Then
					n = GetDlgItem(CLng(DlgText("SuppValueBox")),DlgControlId("InTextBox"))
					i = SetTextBoxLength(n,StrToLong(DlgText("TextLengthBox")),i,True)
					DlgText "TextLengthBox",CStr$(i)
				End If
				DlgText "InTextBox",Temp
				Erase AllStrList,UseStrList
			End If
			If UBound(FileDataList) = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",False
			ElseIf j = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf j = UBound(FileDataList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",True
			End If
			DlgVisible "FilterButton",True
			DlgVisible "CloseFilterButton",False
		End Select
	End Select
End Function


'除去字串前后指定的 PreStr 和 AppStr
'fType = -1 不去除字串前后的空格和所有指定的 PreStr 和 AppStr，但不去除字串内前后空格
'fType = 0 去除字串前后的空格和所有指定的 PreStr 和 AppStr，但不去除字串内前后空格
'fType = 1 去除字串前后的空格和所有指定的 PreStr 和 AppStr，并去除字串内前后空格
'fType = 2 去除字串前后的空格和指定的 PreStr 和 AppStr 1 次，但不去除字串内前后空格
'fType > 2 去除字串前后的空格和指定的 PreStr 和 AppStr 1 次，并去除字串内前后空格
Public Function RemoveBackslash(ByVal Path As String,ByVal PreStr As String,ByVal AppStr As String,ByVal fType As Long) As String
	Dim i As Long,a As Long,p As Long,Stemp As Boolean
	RemoveBackslash = Path
	If Path = "" Then Exit Function
	a = Len(AppStr)
	p = Len(PreStr)
	If fType > -1 Then RemoveBackslash = Trim(RemoveBackslash)
	Do
		Stemp = False
		If p <> 0 Then
			If Left$(RemoveBackslash,p) = PreStr Then
				RemoveBackslash = Mid$(RemoveBackslash,p + 1)
				Stemp = True
			End If
		End If
		If a <> 0 Then
			If Right$(RemoveBackslash,a) = AppStr Then
				RemoveBackslash = Left$(RemoveBackslash,Len(RemoveBackslash) - a)
				Stemp = True
			End If
		End If
		If fType = 1 Or fType > 2 Then RemoveBackslash = Trim$(RemoveBackslash)
		If Stemp = True Then
			If fType < 2 Then i = 0 Else i = 1
		Else
			i = 1
		End If
	Loop Until i = 1
End Function


'输入编辑程序
Private Function CommandInput(CmdPath As String,Argument As String) As Boolean
	Dim MsgList() As String,TempList() As String
	If getMsgList(UIDataList,MsgList,"CommandInput",1) = False Then Exit Function
	ToolsBak = Tools
	Begin Dialog UserDialog 540,294,MsgList(0),.CommandInputDlgFunc ' %GRID:10,7,1,1
		Text 10,7,520,140,MsgList(1),.TipText
		Text 10,154,490,14,MsgList(2),.CmdPathText
		TextBox 10,175,490,21,.CmdPath
		PushButton 500,175,30,21,MsgList(3),.BrowseButton
		Text 10,210,490,14,MsgList(4),.ArgumentText
		TextBox 10,231,490,21,.Argument
		PushButton 500,231,30,21,MsgList(5),.FileArgButton

		Text 10,7,410,14,MsgList(6),.EditerListText
		ListBox 10,28,410,119,TempList(),.EditerList
		ListBox 10,28,410,119,TempList(),.EditerListBak
		PushButton 430,28,100,21,MsgList(7),.AddButton
		PushButton 430,49,100,21,MsgList(8),.ChangeButton
		PushButton 430,70,100,21,MsgList(9),.DelButton

		PushButton 20,266,100,21,MsgList(10),.ClearButton
		PushButton 130,266,120,21,MsgList(11),.EditerListButton
		PushButton 130,266,100,21,MsgList(12),.ResetButton
		PushButton 310,266,100,21,MsgList(13),.SaveButton
		OKButton 310,266,100,21,.OKButton
		CancelButton 420,266,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CmdPath = CmdPath
	dlg.Argument = Argument
	If Dialog(dlg) = 0 Then Exit Function
	If dlg.CmdPath <> "" Then
		CmdPath = dlg.CmdPath
		Argument = dlg.Argument
		CommandInput = True
	End If
End Function


'获取编辑程序对话框函数
Private Function CommandInputDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,x As Long,y As Long,Path As String
	Dim Temp As String,TempList() As String,TempArray() As String,MsgList() As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgVisible "EditerListText",False
		DlgVisible "EditerList",False
		DlgVisible "EditerListBak",False
		DlgVisible "AddButton",False
		DlgVisible "ChangeButton",False
		DlgVisible "DelButton",False
		DlgVisible "ResetButton",False
		DlgVisible "SaveButton",False
		If UBound(Tools) < 4 Then
			DlgVisible "EditerListButton",False
		End If
		'设置当前对话框字体
		If CheckFont(LFList(0)) = True Then
			x = CreateFont(0,LFList(0))
			If x = 0 Then Exit Function
			For i = 0 To DlgCount() - 1
				SendMessageLNG(GetDlgItem(SuppValue,DlgControlId(DlgName(i))),WM_SETFONT,x,0)
			Next i
		End If
	Case 2 ' 数值更改或者按下了按钮
		CommandInputDlgFunc = True ' 防止按下按钮关闭对话框窗口
		Select Case DlgItem$
		Case "CancelButton"
			If DlgVisible("DelButton") = False Then
				CommandInputDlgFunc = False
				Exit Function
			End If
			Tools = ToolsBak
			DlgVisible "TipText",True
			DlgVisible "OKButton",True
			DlgVisible "EditerListText",False
			DlgVisible "EditerList",False
			DlgVisible "AddButton",False
			DlgVisible "ChangeButton",False
			DlgVisible "DelButton",False
			DlgVisible "ResetButton",False
			DlgVisible "SaveButton",False
			If UBound(Tools) < 4 Then
				DlgVisible "EditerListButton",False
			Else
				DlgVisible "EditerListButton",True
			End If
			Exit Function
		Case "OKButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			Temp = Trim$(DlgText("CmdPath"))
			If Temp = "" Then
				MsgBox MsgList(6),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			Else
				TempList = ReSplit(Temp,"%")
				If UBound(TempList) = 2 Then
					Temp = Replace$(Temp,"%" & TempList(1) & "%",Environ(TempList(1)))
				End If
				If Dir$(Temp) = "" Then
					MsgBox MsgList(7),vbOkOnly+vbInformation,MsgList(0)
					Exit Function
				End If
			End If
			CommandInputDlgFunc = False
			Exit Function
		Case "SaveButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			ReDim TempList(UBound(Tools)) As String,TempArray(UBound(Tools)) As String
			For i = 4 To UBound(Tools)
				Temp = Trim$(Tools(i).FilePath)
				If Temp = "" Then
					TempList(x) = Tools(i).sName
					x = x + 1
				ElseIf Dir$(Temp) = "" Then
					TempArray(y) = Tools(i).FilePath
					y = y + 1
				End If
			Next i
			If x > 0 Then
				ReDim Preserve TempList(x - 1) As String
				MsgBox Replace$(MsgList(8),"%s",Join(TempList,vbCrLf)),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			ElseIf y > 0 Then
				ReDim Preserve TempArray(y - 1) As String
				MsgBox Replace$(MsgList(9),"%s",Join(TempArray,vbCrLf)),vbOkOnly+vbInformation,MsgList(0)
				Exit Function
			End If
			If WriteEngineSet(tWriteLoc,"Tools") = False Then
				MsgBox(Replace(MsgList(12),"%s",tWriteLoc),vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
			MsgBox(MsgList(13),vbOkOnly+vbInformation,MsgList(10))
			ToolsBak = Tools
			DlgVisible "TipText",True
			DlgVisible "OKButton",True
			DlgVisible "EditerListText",False
			DlgVisible "EditerList",False
			DlgVisible "AddButton",False
			DlgVisible "ChangeButton",False
			DlgVisible "DelButton",False
			DlgVisible "ResetButton",False
			DlgVisible "SaveButton",False
			If UBound(Tools) < 4 Then
				DlgVisible "EditerListButton",False
			Else
				DlgVisible "EditerListButton",True
			End If
			Exit Function
		Case "BrowseButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			If PSL.SelectFile(Path,True,MsgList(2),MsgList(1)) = False Then
				Exit Function
			End If
			DlgText "CmdPath",Path
			If DlgVisible("SaveButton") = True Then
				Temp = DlgText("EditerList")
				For i = 4 To UBound(Tools)
					If Tools(i).sName = Temp Then
						Tools(i).FilePath = Path
						Exit For
					End If
				Next i
			End If
		Case "FileArgButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			ReDim TempList(0) As String
			TempList(0) = MsgList(3)
			If ShowPopupMenu(TempList,vbPopupUseRightButton) < 0 Then Exit Function
			DlgText "Argument",DlgText("Argument") & " " & """%1"""
			If DlgVisible("SaveButton") = True Then
				Temp = DlgText("EditerList")
				For i = 4 To UBound(Tools)
					If Tools(i).sName = Temp Then
						Tools(i).Argument = DlgText("Argument")
						Exit For
					End If
				Next i
			End If
		Case "EditerListButton"
			Temp = Trim$(DlgText("CmdPath"))
			x = UBound(Tools) - 4
			ReDim TempList(x) As String
			For i = 0 To x
				TempList(i) = Tools(i + 4).sName
				If Tools(i + 4).FilePath = Temp Then y = i
			Next i
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",y
			DlgValue "EditerListBak",y
			DlgText "CmdPath",Tools(y + 4).FilePath
			DlgText "Argument",Tools(y + 4).Argument
			DlgVisible "TipText",False
			DlgVisible "OKButton",False
			DlgVisible "EditerListButton",False
			DlgVisible "EditerListText",True
			DlgVisible "EditerList",True
			DlgVisible "AddButton",True
			DlgVisible "ChangeButton",True
			DlgVisible "DelButton",True
			DlgVisible "ResetButton",True
			DlgVisible "SaveButton",True
		Case "AddButton"
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			If PSL.SelectFile(Path,True,MsgList(2),MsgList(1)) = False Then
				Exit Function
			End If
			Temp = Mid$(Path,InStrRev(Path,"\") + 1)
			For i = 0 To UBound(Tools)
				If InStr(LCase$(Tools(i).FilePath),LCase$(Temp)) Then
					MsgBox(MsgList(11),vbOkOnly+vbInformation,MsgList(10))
					Exit Function
				End If
			Next i
			Call AddTools(Tools,Temp,Path,"")
			x = UBound(Tools) - 4
			ReDim TempList(x) As String
			For i = 0 To x
				TempList(i) = Tools(i + 4).sName
			Next i
			DlgText "CmdPath",Path
			DlgText "Argument",""
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",x
			DlgValue "EditerListBak",x
		Case "ChangeButton"
			If DlgValue("EditerList") < 0 Then Exit Function
			x = UBound(Tools) - 4
			ReDim TempList(x) As String
			For i = 0 To x
				TempList(i) = Tools(i + 4).sName
			Next i
			x = DlgValue("EditerList")
			Temp = EditSet(TempList,x)
			If Temp = "" Then Exit Function
			TempList(x) = Temp
			Tools(x + 4).sName = Temp
			DlgListBoxArray "EditerList",TempList()
			DlgListBoxArray "EditerListBak",TempList()
			DlgValue "EditerList",x
			DlgValue "EditerListBak",x
			Exit Function
		Case "DelButton"
			If DlgValue("EditerList") < 0 Then Exit Function
			If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
			Temp = DlgText("EditerList")
			If MsgBox(Replace(MsgList(5),"%s",Temp),vbYesNo+vbInformation,MsgList(4)) = vbNo Then
				Exit Function
			End If
			x = UBound(Tools)
			ReDim TempList(x) As String
			For i = 0 To x
				TempList(i) = Tools(i).sName
			Next i
			y = DlgValue("EditerList") + 4
			Call DelToolsArrays(TempList,Tools,y)
			If y > 0 And y = x Then y = x - 1
			If UBound(Tools) > 3 Then
				x = UBound(Tools) - 4
				ReDim TempList(x) As String
				For i = 0 To x
					TempList(i) = Tools(i + 4).sName
				Next i
				DlgListBoxArray "EditerList",TempList()
				DlgListBoxArray "EditerListBak",TempList()
				DlgValue "EditerList",y - 4
				DlgValue "EditerListBak",y - 4
				DlgText "CmdPath",Tools(y).FilePath
				DlgText "Argument",Tools(y).Argument
			Else
				ReDim TempList(0) As String
				DlgListBoxArray "EditerList",TempList()
				DlgListBoxArray "EditerListBak",TempList()
				DlgValue "EditerList",0
				DlgValue "EditerListBak",0
				DlgText "CmdPath",""
				DlgText "Argument",""
			End If
		Case "ClearButton"
			DlgText "CmdPath",""
 			DlgText "Argument",""
 		Case "ResetButton"
			Temp = DlgText("EditerList")
			If Temp = "" Then Exit Function
			For i = 3 To UBound(ToolsBak)
				If Tools(i).sName = Temp Then
					DlgText "CmdPath",ToolsBak(i).FilePath
					DlgText "Argument",ToolsBak(i).Argument
					Exit For
				End If
			Next i
		Case "EditerList"
			x = DlgValue("EditerList")
			If x < 0 Then Exit Function
			DlgValue "EditerList",x
			DlgValue "EditerListBak",x
			DlgText "CmdPath",Tools(x + 4).FilePath
			DlgText "Argument",Tools(x + 4).Argument
		End Select
	Case 3 ' 文本框或者组合框文本被更改
		If DlgItem$ = "CmdPath" Or DlgItem$ = "Argument" Then
			If DlgItem$ = "CmdPath" Then
				Temp = Trim$(DlgText("CmdPath"))
				If Temp <> "" Then
					TempList = ReSplit(Temp,"%")
					If UBound(TempList) = 2 Then
						Temp = Replace$(Temp,"%" & TempList(1) & "%",Environ(TempList(1)))
					End If
					If Dir$(Temp) = "" Then
						If getMsgList(UIDataList,MsgList,"CommandInputDlgFunc",1) = False Then Exit Function
						MsgBox MsgList(7),vbOkOnly+vbInformation,MsgList(0)
					End If
				Else
					DlgText "CmdPath",Temp
				End If
			Else
				Temp = Trim$(DlgText("Argument"))
				DlgText "Argument",Temp
			End If
			If DlgVisible("DelButton") = True Then
				Temp = DlgText("EditerListBak")
				If Temp = "" Then Exit Function
				For i = 4 To UBound(Tools)
					If Tools(i).sName = Temp Then
						If DlgItem$ = "CmdPath" Then
							Tools(i).FilePath = DlgText("CmdPath")
						Else
							Tools(i).Argument = DlgText("Argument")
						End If
						Exit For
					End If
				Next i
			End If
		End If
	End Select
End Function


' 检查文件编码
' ----------------------------------------------------
' ANSI      无格式定义
' 2B2F 76[38|39|2B|2F] UTF-7
' EFBB BF   UTF-8
' FFFE      UTF-16LE/UCS-2, Little Endian with BOM
' FEFF      UTF-16BE/UCS-2, Big Endian with BOM
' XX00 XX00 UTF-16LE/UCS-2, Little Endian without BOM
' 00XX 00XX UTF-16BE/UCS-2, Big Endian without BOM
' FFFE 0000 UTF-32LE/UCS-4, Little Endian with BOM
' 0000 FEFF UTF-32BE/UCS-4, Big Endian with BOM
' XX00 0000 UTF-32LE/UCS-4, Little Endian without BOM
' 0000 00XX UTF-32BE/UCS-4, Big Endian without BOM
' 上述中的 XX 表示任意十六进制字符
Function CheckCode(ByVal FilePath As String) As String
	Dim i As Long,objStream As Object,Bytes() As Byte,Temp As String

	If Dir$(FilePath) = "" Then Exit Function
	i = FileLen(FilePath)
	If i = 0 Then
		CheckCode = "ANSI"
		Exit Function
	End If
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo 0
	If objStream Is Nothing Then
		CheckCode = "ANSI"
		Exit Function
	End If
	If i > 1 Then
		With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Position = 0
			.LoadFromFile FilePath
			Bytes = .read(IIf(i > 3,4,i))
			.Close
		End With
		If i > 2 Then
			If Bytes(0) = &HEF And Bytes(1) = &HBB And Bytes(2) = &HBF Then
				CheckCode = "utf-8EFBB"
			ElseIf Bytes(0) = &HFF And Bytes(1) = &HFE Then
				CheckCode = "unicodeFFFE"
			ElseIf Bytes(0) = &HFE And Bytes(1) = &HFF Then
				CheckCode = "unicodeFEFF"
			End If
		Else
			If Bytes(0) = &HFF And Bytes(1) = &HFE Then
				CheckCode = "unicodeFFFE"
			ElseIf Bytes(0) = &HFE And Bytes(1) = &HFF Then
				CheckCode = "unicodeFEFF"
			End If
		End If
		If i > 3 Then
			If Bytes(0) <> &H00 And Bytes(1) = &H00 And Bytes(2) <> &H00 And Bytes(3) = &H00 Then
				CheckCode = "utf-16LE"
			ElseIf Bytes(0) = &H00 And Bytes(1) <> &H00 And Bytes(2) = &H00 And Bytes(3) <> &H00 Then
				CheckCode = "utf-16BE"
			ElseIf Bytes(0) = &HFF And Bytes(1) = &HFE And Bytes(2) = &H00 And Bytes(3) = &H00 Then
				CheckCode = "unicode-32FFFE"
			ElseIf Bytes(0) = &H00 And Bytes(1) = &H00 And Bytes(2) = &HFE And Bytes(3) = &HFF Then
				CheckCode = "unicode-32FEFF"
			ElseIf Bytes(0) = &H00 And Bytes(1) = &H00 And Bytes(2) = &H00 And Bytes(3) <> &H00 Then
				CheckCode = "utf-32LE"
			ElseIf Bytes(0) <> &H00 And Bytes(1) = &H00 And Bytes(2) = &H00 And Bytes(3) = &H00 Then
				CheckCode = "utf-32BE"
			End If
		End If
		If CheckCode <> "" Then
			Set objStream = Nothing
			Exit Function
		End If
	End If
	On Error GoTo ExitFunction
	With objStream
		.Type = 2
		.Mode = 3
		.Open
		.Charset = "_autodetect_all"
		.Position = 0
		.LoadFromFile FilePath
		Temp = .ReadText
		'On Error GoTo NextNum
		For i = 38 To 2 Step -1
			If i <> 9 And i <> 13 Then
				.Position = 0
				.Charset = CodeList(i).CharSet
				If Temp = .ReadText Then
					CheckCode = .Charset
					.Close
					Set objStream = Nothing
					Exit Function
				End If
			End If
			NextNum:
		Next i
		'On Error GoTo NextPos
		For i = 41 To 39 Step -1
			If i <> 40 Then
				.Position = 0
				.Charset = CodeList(i).CharSet
				If Temp = .ReadText Then
					CheckCode = .Charset
					.Close
					Set objStream = Nothing
					Exit Function
				End If
			End If
			NextPos:
		Next i
		.Close
	End With
	ExitFunction:
	If CheckCode = "" Then CheckCode = "ANSI"
	Set objStream = Nothing
End Function


' 读取文本文件
Function ReadFile(ByVal FilePath As String,CharSet As String) As String
	Dim i As Long,objStream As Object,Code As String,FN As Variant
	If Dir$(FilePath) = "" Then Exit Function
	i = FileLen(FilePath)
	If i = 0 Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		If Code = "" Then Code = "_autodetect_all"
		If Code <> "ANSI" Then
			With objStream
				.Type = 2
				.Mode = 3
				.CharSet = IIf(Code = "utf-8EFBB","utf-8",Code)
				.Open
				.LoadFromFile FilePath
				ReadFile = .ReadText
				.Close
			End With
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		ReDim readByte(i - 1) As Byte
		FN = FreeFile
		Open FilePath For Binary Access Read As #FN
		Get #FN,,readByte
		Close #FN
		ReadFile = StrConv$(readByte,vbUnicode)
	End If
	If CharSet = "" Then CharSet = Code
	Set objStream = Nothing
	Exit Function

	ErrorMsg:
	ReadFile = ""
	If objStream Is Nothing Or Code = "ANSI" Then Close #FN
	Set objStream = Nothing
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & FilePath
	Call sysErrorMassage(Err,1)
End Function


' 写入文本文件
Function WriteToFile(ByVal FilePath As String,ByVal textStr As String,CharSet As String) As Boolean
	Dim objStream As Object,Code As String,Bytes() As Byte,FN As Variant
	If FilePath = "" Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		If LCase$(Code) = "_autodetect_all" Then Code = "ANSI"
		If Code <> "ANSI" Then
			With objStream
				.Type = 2
				.Mode = 3
				.CharSet = IIf(Code = "utf-8EFBB","utf-8",Code)
				.Open
				.WriteText textStr
				'去除不带 BOM 格式的 BOM
				If Code = "utf-16LE" Or Code = "utf-8" Then
					.Position = 0
					.Type = 1
					.Position = IIf(Code = "utf-16LE",2,3)
					Bytes = .Read(.Size - IIf(Code = "utf-16LE",2,3))
					.Position = 0
					.SetEOS
					.Write Bytes
				End If
				.SaveToFile FilePath,2
				.Close
			End With
			WriteToFile = True
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		Bytes = StrConv(textStr,vbFromUnicode)
		FN = FreeFile
		Open FilePath For Binary Access Write Lock Write As #FN
		Put #FN,,Bytes
		Close #FN
		WriteToFile = True
	End If
	If CharSet = "" Then CharSet = Code
	Set objStream = Nothing
	Exit Function

	ErrorMsg:
	If objStream Is Nothing Or Code = "ANSI" Then Close #FN
	Set objStream = Nothing
	Err.Source = "NotWriteFile"
	Err.Description = Err.Description & JoinStr & FilePath
	Call sysErrorMassage(Err,1)
End Function


'创建 Adodb.Stream 使用的代码页数组
Public Function getCodePageList(ByVal MinNum As Long,ByVal MaxNum As Long) As CODEPAGE_DATA()
	Dim i As Long,MsgList() As String,Code As String
	Dim CharSetList() As String,CodePage() As CODEPAGE_DATA

	If getMsgList(UIDataList,MsgList,"CodePageList",0) = False Then Exit Function
	Code = "ANSI|_autodetect_all|gb2312|hz-gb-2312|gb18030|big5|euc-jp|iso-2022-jp|shift_jis|" & _
			"_autodetect|ks_c_5601-1987|euc-kr|iso-2022-kr|_autodetect_kr|windows-874|" & _
			"windows-1258|iso-8859-4|windows-1257|ASMO-708|DOS-720|iso-8859-6|windows-1256|" & _
			"DOS-862|iso-8859-8-i|iso-8859-8|windows-1255|iso-8859-9|iso-8859-7|windows-1253|" & _
			"iso-8859-1|cp866|iso-8859-5|koi8-r|koi8-ru|windows-1251|ibm852|iso-8859-2|" & _
			"windows-1250|iso-8859-3|utf-7|utf-8EFBB|utf-8|unicodeFFFE|unicodeFEFF|utf-16LE|" & _
			"utf-16BE|unicode-32FFFE|unicode-32FEFF|utf-32LE|utf-32BE"
	CharSetList = ReSplit(Code,"|")

	i = UBound(MsgList)
	If MaxNum < 0 Or MaxNum > i Then MaxNum = i
	ReDim CodePage(MaxNum - MinNum) As CODEPAGE_DATA
	For i = MinNum To MaxNum
		CodePage(i - MinNum).sName = MsgList(i)
		CodePage(i - MinNum).CharSet = CharSetList(i)
	Next i
	getCodePageList = CodePage
End Function


'读取语言对
Function LangCodeList(ByVal DataName As String,ByVal MinNum As Long,ByVal MaxNum As Long) As String()
	Dim i As Long,n As Long,LangMaxNum As Long,LangName As String
	Dim LangPairs() As String,TempList() As String,Stemp As Boolean

	Stemp = CheckArray(PslLangDataList)
	If Stemp = False Then
		PslLangCode = "|af|sq|gsw|am|ar|hy|as|az|bn|ba|eu|be|BN|bs|br|bg|my|ca|tzm|ku|chr|zh-CN|zh-TW|co|hr|cs|da|prs|nl|" & _
				"en|et|fo|fa|fil-PH|fi|fr|fy|ff|gl|ka|de|el|kl|gn|gu|ha|haw|he|hi|hu|is|ig|id|iu|ga|xh|zu|it|ja|" & _
				"jv|kn|KS|kk|km|qut|rw|kok|ko|kz|ky|lo|lv|lt|dsb|lb|mk|ms|ML|mt|mi|arn|mr|moh|mn|ne|no|nb|nn|oc|" & _
				"or|om|ps|pl|pt|pa|qu|ro|rm|ru|sah|smn|smj|se|sms|sma|sa|gd|sr-Cyrl|sr|nso|st|tn|SD|si|sk|sl|so|st|es|" & _
				"sw|sv|sy|tg|tzm|ta|tt|te|th|bo|ti|ts|tr|tk|ug|uk|hsb|ur|uz|ca|vi|cy|wo|ii|yo"
		PslLangDataList = ReSplit(PslLangCode,LngJoinStr)
	End If

	BingLangCode = "|||||ar|||||||||bs-Latn||bg||ca||||zh-CHS|zh-CHT||hr|cs|da||nl|" & _
				"en|et||fa||fi|fr|||||de|el||||||he|hi|hu|||id|||||it|ja|" & _
				"||||||||ko||||lv|lt||||ms||mt|||||||no|no|no||" & _
				"|||pl|pt|||ro||ru|||||||||sr-Cyrl|sr-Latn||||||sk|sl|||es|" & _
				"|sv|||||||th||||tr|||uk||ur|||vi|cy|||"

	GoogleLangCode = "auto|af|sq|||ar|hy||az|||eu|be|bn|bs||bg|my|ca||||zh-CN|zh-TW||hr|cs|da||nl|" & _
				"en|et||fa|tl|fi|fr|||gl|ka|de|el|||gu|ha||iw|hi|hu|is|ig|id||ga||zu|it|ja|" & _
				"jw|kn||kk|km||||ko|||lo|lv|lt|||mk|ms|ml|mt|mi||mr||mn|ne|no|no|no||" & _
				"|||pl|pt|pa||ro||ru||||||||||sr|||||si|sk|sl|so|st|es|" & _
				"sw|sv||tg||ta||te|th||||tr|||uk||ur|uz||vi|cy|||yo"

	YahooLangCode = "||||||||||||||||||||||zh|zt||||||nl|" & _
				"en||||||fr|||||de|el||||||||||||||||it|ja|" & _
				"||||||||ko||||||||||||||||||||||" & _
				"||||pt|||||ru||||||||||||||||||||es|" & _
				"||||||||||||||||||||||||"

	en2zhCheck = "||||||||||||||||||||||zh-CN|zh-TW|||||||" & _
				"|||||||||||||||||||||||||||||ja|" & _
				"||||||||ko||||||||||||||||||||||" & _
				"||||||||||||||||||||||||||||||" & _
				"||||||||||||||||||||||||"

	zh2enCheck = "|af|sq|gsw|am|ar|hy|as|az|bn|ba|eu|be|BN|bs|br|bg|my|ca|tzm|ku|chr|||co|hr|cs|da|prs|nl|" & _
				"en|et|fo|fa|fil-PH|fi|fr|fy|ff|gl|ka|de|el|kl|gn|gu|ha|haw|he|hi|hu|is|ig|id|iu|ga|xh|zu|it||" & _
				"jv|kn|KS|kk|km|qut|rw|kok||kz|ky|lo|lv|lt|dsb|lb|mk|ms|ML|mt|mi|arn|mr|moh|mn|ne|no|nb|nn|oc|" & _
				"or|om|ps|pl|pt|pa|qu|ro|rm|ru|sah|smn|smj|se|sms|sma|sa|gd|sr-Cyrl|sr|nso|st|tn|SD|si|sk|sl|so|st|es|" & _
				"sw|sv|sy|tg|tzm|ta|tt|te|th|bo|ti|ts|tr|tk|ug|uk|hsb|ur|uz|ca|vi|cy|wo|ii|yo"

	LangMaxNum = UBound(PslLangDataList)
	If MaxNum < 0 Or MaxNum > LangMaxNum Then MaxNum = LangMaxNum
	ReDim TempList(LangMaxNum),LangPairs(MaxNum - MinNum)

	Select Case DataName
	Case DefaultEngineList(0)
		TempList = ReSplit(BingLangCode,LngJoinStr)
	Case DefaultEngineList(1)
		TempList = ReSplit(GoogleLangCode,LngJoinStr)
	Case DefaultEngineList(2)
		TempList = ReSplit(YahooLangCode,LngJoinStr)
	Case DefaultCheckList(0)
		TempList = ReSplit(en2zhCheck,LngJoinStr)
	Case DefaultCheckList(1)
		TempList = ReSplit(zh2enCheck,LngJoinStr)
	End Select

	For i = 0 To LangMaxNum
		If Stemp = False Then
			If LangCode = "zh-CN" Or LangCode = "zh-TW" Or LangCode = "fil-PH" Then
				LangName = PSL.GetLangCode(PSL.GetLangID(PslLangDataList(i),pslCodeLangRgn),pslCodeText)
			Else
				LangName = PSL.GetLangCode(PSL.GetLangID(PslLangDataList(i),pslCode639_1),pslCodeText)
				If LangName = "3FF3F" Then
					LangName = PSL.GetLangCode(PSL.GetLangID(PslLangDataList(i),pslCodeLangRgn),pslCodeText)
				End If
			End If
			PslLangDataList(i) = LangName & LngJoinStr & PslLangDataList(i)
			If LangName <> "3FF3F" Then
				If i >= MinNum And i <= MaxNum Then
					LangPairs(n) = PslLangDataList(i) & LngJoinStr & TempList(i)
					n = n + 1
				End If
			End If
		ElseIf i >= MinNum And i <= MaxNum Then
			If ReSplit(PslLangDataList(i),LngJoinStr)(0) <> "3FF3F" Then
				LangPairs(n) = PslLangDataList(i) & LngJoinStr & TempList(i)
				n = n + 1
			End If
		End If
	Next i
	If n > 0 Then n = n - 1
	ReDim Preserve LangPairs(n) As String
	LangCodeList = LangPairs
End Function


'从来源列表中获取指定项目的目标列表
'Mode = 0 发生错误直接退出程序，Mode = 1 发生错误给出是否退出程序提示
Function getMsgList(SourceList() As INIFILE_DATA,TargetList() As String,Titles As String,ByVal Mode As Long) As Boolean
	Dim i As Long,j As Long,n As Long,m As Long,Temp As String
	Temp = "|" & Titles & "|"
	n = -1
	For i = 0 To UBound(SourceList)
		If InStr(Temp,"|" & SourceList(i).Title & "|") > 0 Then
			If n = -1 Then
				TargetList = SourceList(i).Value
				n = UBound(TargetList)
			Else
				m = UBound(SourceList(i).Value)
				ReDim Preserve TargetList(n + m + 1) As String
				For j = 0 To m
					TargetList(n + 1) = SourceList(i).Value(j)
				Next j
			End If
			Temp = Replace$(Temp,"|" & SourceList(i).Title & "|","")
			If Temp = "" Then Exit For
		End If
	Next i
	If n > -1 Then
		Titles = RemoveBackslash(Temp,"|","|",1)
		getMsgList = True
	Else
		If Mode = 0 Then
			Err.Raise(1,"NotSection",LangFile & JoinStr & Titles)
		ElseIf Mode < 3 Then
			On Error GoTo ErrorMassage
			Err.Raise(1,"NotSection",LangFile & JoinStr & Titles)
			ErrorMassage:
			Call sysErrorMassage(Err,Mode)
		End If
	End If
End Function


'读取 INI 文件
'Mode = 0 删除项目值前后空格，并转义项目值
'Mode = 1 删除项目值前空格，不转义
'Mode > 1 不删除项目值前后空格，不转义
Function getINIFile(DataList() As INIFILE_DATA,ByVal File As String,Code As String,ByVal Mode As Long) As Boolean
	Dim i As Long,m As Long,n As Long,j As Long,Max As Long,iMax As Long,vMax As Long
	Dim ItemList() As String,ValueList() As String,TempArray() As String
	Dim l As String,Header As String,HeaderBak As String,Temp As String
	If Trim$(File) <> "" Then
		TempArray = ReSplit(ReadFile(File,Code),vbCrLf)
	ElseIf File = "" And Code <> "" Then
		TempArray = ReSplit(Code,vbCrLf)
	Else
		Exit Function
	End If
	If CheckArray(TempArray) = False Then Exit Function
	ReDim DataList(iMax) As INIFILE_DATA
	ReDim ItemList(vMax) As String,ValueList(vMax) As String
	Max = UBound(TempArray)
	For i = 0 To Max
		l$ = Trim$(TempArray(i))
		If l$ <> "" Then
			If l$ Like "[[]*]" Then
				Header$ = Trim$(Mid$(l$,2,Len(l$)-2))
			End If
			If Header$ <> "" And HeaderBak$ = "" Then HeaderBak$ = Header$
			If Header$ <> "" And Header$ = HeaderBak$ Then
				j = InStr(l$,"=")
				If j > 0 Then
					Temp$ = Trim$(Left$(l$,j - 1))
					If Temp$ <> "" Then
						If n > vMax Then
							 vMax = n * 100
							 ReDim Preserve ItemList(vMax) As String,ValueList(vMax) As String
						End If
						ItemList(n) = Temp$
						If Mode = 0 Then
							ValueList(n) = Convert(RemoveBackslash(Mid$(l$,j + 1),"""","""",2))
						ElseIf Mode = 1 Then
							ValueList(n) = LTrim$(Mid$(l$,j + 1))
						Else
							ValueList(n) = Mid$(l$,j + 1)
						End If
						n = n + 1
					End If
				End If
			End If
		End If
		If Header$ <> "" And (i = Max Or Header$ <> HeaderBak$) Then
			If n > 0 Then
				If m > iMax Then
					iMax = m * 50
					ReDim Preserve DataList(iMax) As INIFILE_DATA
				End If
				ReDim Preserve ItemList(n - 1) As String,ValueList(n - 1) As String
				DataList(m).Title = HeaderBak$
				DataList(m).Item = ItemList
				DataList(m).Value = ValueList
				m = m + 1: n = 0: vMax = 0
				ReDim ItemList(vMax) As String,ValueList(vMax) As String
			End If
			HeaderBak$ = Header$
		End If
	Next i
	If m > 0 Then
		m = m - 1
		getINIFile = True
	End If
	ReDim Preserve DataList(m) As INIFILE_DATA
End Function


'获取语言文件列表
Function GetUIList(ByVal UIDir As String,UIList() As UI_FILE) As Boolean
	Dim i As Long,j As Long,m As Long,n As Long,Max As Long
	Dim l As String,File As String,Header As String
	Dim TempArray() As String,tmpUIFile As UI_FILE
	Dim readByte() As Byte,FN As Variant
	ReDim UIList(0) As UI_FILE
	On Error Resume Next
		With tmpUIFile
		File = Dir$(UIDir & AppName & "_*.lng")
		Do While File <> ""
			If LCase$(Mid$(File,InStrRev(File,".") + 1)) = "lng" Then
				.FilePath = UIDir & File
				i = FileLen(.FilePath)
				ReDim readByte(i) As Byte
				FN = FreeFile
				Open .FilePath For Binary Access Read Lock Write As #FN
				Get #FN,,readByte
				Close #FN
				'On Error GoTo 0
				TempArray = ReSplit(readByte,vbCrLf)
				Max = UBound(TempArray)
				For i = 0 To Max
					l$ = Trim(TempArray(i))
					If l$ <> "" Then
						If l$ Like "[[]*]" Then
							Header$ = Trim$(Mid$(l$,2,Len(l$) - 2))
						End If
						If Header$ = "Option" Then
							j = InStr(l$,"=")
							If j > 0 Then
								Select Case Trim$(Left$(l$,j - 1))
								Case "AppName"
									.AppName = RemoveBackslash(Mid$(l$,j + 1),"""","""",1)
								Case "Version"
									.Version = RemoveBackslash(Mid$(l$,j + 1),"""","""",1)
								Case "LanguageName"
									.LangName = RemoveBackslash(Mid$(l$,j + 1),"""","""",1)
								Case "LanguageID"
									.LangID = RemoveBackslash(Mid$(l$,j + 1),"""","""",1)
								Case "Encoding"
									.Encoding = RemoveBackslash(Mid$(l$,j + 1),"""","""",1)
								End Select
							End If
						End If
					End If
					If Header$ <> "" And (i = Max Or Header$ <> "Option") Then
						Header = ""
						Exit For
					End If
				Next i
				If LCase$(.AppName) = LCase$(AppName) Then
					If .Version = Version And .LangName <> "" Then
						If n > m Then
							m = n * 10
							ReDim Preserve UIList(m) As UI_FILE
						End If
						UIList(n) = tmpUIFile
						n = n + 1
					End If
					.AppName = ""
					.Version = ""
					.LangName = ""
					.LangID = ""
					.Encoding = ""
				End If
			End If
			File = Dir$()
		Loop
	End With
	If n > 0 Then
		GetUIList = True
		n = n - 1
	End If
	ReDim Preserve UIList(n) As UI_FILE
End Function


'读取界面语言字串
Public Function GetUI(ByVal UIDir As String,ByVal UISet As String,ByVal OSLng As String,UIData() As INIFILE_DATA,UIFiles() As UI_FILE,LngFile As String) As Boolean
	Dim i As Long,j As Long,n As Long
	Dim TempList() As String,TempArray() As String
	If GetUIList(UIDir,UIFiles) = True Then
		If UISet = "" Or UISet = "0" Then
			UISet = OSLng
		ElseIf UISet = "1" Then
			UISet = Right$("0" & Hex$(PSL.Option(pslOptionSystemLanguage)),4)
		End If
		UISet = LCase$(UISet)
		TempList = ReSplit(UISet,";")
		For i = 0 To UBound(UIFiles)
			If LCase$(UIFiles(i).LangID) = UISet Then
				LngFile = UIFiles(i).FilePath
				Exit For
			End If
			TempArray = ReSplit(LCase$(UIFiles(i).LangID),";")
			For j = 0 To UBound(TempList)
				For n = 0 To UBound(TempArray)
					If TempList(j) = TempArray(n) Then
						LngFile = UIFiles(i).FilePath
						Exit For
					End If
				Next n
				If LngFile <> "" Then Exit For
			Next j
			If LngFile <> "" Then Exit For
		Next i
	End If
	If LngFile = "" Then LngFile = UIDir & AppName & "_" & OSLng & ".lng"
	If Dir$(LngFile) = "" Then
		LngFile = UIDir & AppName & "_" & OSLng & ".lng"
		If Dir$(LngFile) = "" Then
			For i = 0 To 2
				If i = 0 Then
					LngFile = UIDir & AppName & "_0804.lng"
				ElseIf i = 1 Then
					LngFile = UIDir & AppName & "_0404.lng"
				Else
					LngFile = UIFiles(0).FilePath
				End If
				If Dir$(LngFile) <> "" Then Exit For
				LngFile = ""
			Next i
			If LngFile = "" Then Err.Raise(1,"NotExitFile",UIDir & AppName & "_*.lng")
		End If
	End If
	GetUI = getINIFile(UIData,LngFile,"unicodeFFFE",0)
End Function


'设置光标位置（按字符数计算，起始行和起始列为0）
Public Sub SetCurPos(ByVal hwnd As Long,ptPos As POINTAPI,ByVal Length As Long)
	Dim nPos As Long
	'获取自定行的首字符在文本中的字符数偏移
	nPos = SendMessageLNG(hwnd, EM_LINEINDEX, ptPos.y, 0&)
	'选定指定文本的整个范围
	SendMessageLNG hwnd, EM_SETSEL, nPos + ptPos.x, nPos + ptPos.x + Length
	'将选定内容放到可视范围之内
	SendMessageLNG hwnd, EM_SCROLLCARET, 0&, 0&
End Sub


'获取光标位置（按字符数计算，起如行和起始列均为0）
Private Function GetCurPos(ByVal hwnd As Long) As POINTAPI
	'获取光标所在位置在文本中的字节数偏移(GetCurPos.x = 偏移，GetCurPos.y = 0)
	SendMessage(hwnd, EM_GETSEL, 0&, GetCurPos)
	'获得光标所在行的行号
	GetCurPos.y = SendMessageLNG(hwnd, EM_LINEFROMCHAR, GetCurPos.x, 0&)
	'获得光标所在行的列号(光标位置的字节数偏移 - 所在行首字符在文本中的字符偏移)
	GetCurPos.x = GetCurPos.x - SendMessageLNG(hwnd, EM_LINEINDEX, GetCurPos.y, 0&)
End Function


'获取光标所在行的整行文本
Private Function GetCurPosLine(ByVal hwnd As Long,ptPos As POINTAPI,Optional ByVal Mode As Boolean) As String
	Dim Length As Long
	'获取光标位置（按字符数计算，起如行和起始列均为0）
	If Mode = True Then ptPos = GetCurPos(hwnd)
	'获取光标所在行的首字符在文本中的字符数偏移
	Length = SendMessageLNG(hwnd, EM_LINEINDEX, ptPos.y, 0&)
	'获取光标所在行的文本长度(字符数)
	Length = SendMessageLNG(hwnd, EM_LINELENGTH, Length, 0&)
	If Length < 1 Then Exit Function
	'预设可接收文本内容的字节数，须预先赋空格
	ReDim byteBuffer(Length * 2 + 1) As Byte
	'最大允许存放 1024 个字符
	byteBuffer(1) = 4
	'获取光标所在行的文本字节数组
	SendMessage hwnd, EM_GETLINE, ptPos.y, byteBuffer(0)
	'转换为文本，并清除空字符
	GetCurPosLine = Replace$(StrConv$(byteBuffer,vbUnicode),vbNullChar,"")
End Function


'获取查找字串的查找方式
'GetFindMode = 0 常规，= 1 通配符, = 2 正则表达式
Public Function GetFindMode(FindStr As String) As Long
	'不含通配符和正则表达式专用字符时
	If (FindStr Like "*[$()+.^{|*?#[\]*") = False Then
		Exit Function
	End If
	'不含正则表达式专用字符时
	If (FindStr Like "*[$()+.^{|\]*") = False Then
		If (FindStr Like "*\[*?#[]*") = False Then
			GetFindMode = 1
		End If
		Exit Function
	End If
	GetFindMode = 2
End Function


'查找字串并移动光标位置
'Mode = False 从头到尾，否则从尾到头
'FindCurPos > 0 已找到，= 0 未找到, = -1 相同位置(只找到一项), = -2 查找内容语法
Public Function FindCurPos(ByVal hwnd As Long,ByVal FindStr As String,ByVal Mode As Boolean,Optional StrText As String) As Long
	Dim i As Long,Lines As Long,n As Long,Stemp As Integer,Length As Long
	Dim ptPos As POINTAPI,tmpPos As POINTAPI,Matches As Object
	Dim byteBuffer() As Byte,TempList() As String
	On Error GoTo errHandle
	'获取光标位置及文本框内字串的行数
	Lines = SendMessageLNG(hwnd, EM_GETLINECOUNT, 0&, 0&) - 1 'Lines 以 1 为起点
	If Lines = -1 Then Exit Function
	'检测查找内容的查找方式
	Stemp = GetFindMode(FindStr)
	Select Case Stemp
	Case 0
		If (FindStr Like "*\[*?#[]*") = True Then
			TempList = ReSplit("*,?,#,[",",",-1)
			For i = 0 To UBound(TempList)
				FindStr = Replace$(FindStr,"\" & TempList(i),TempList(i))
			Next i
		End If
		Length = Len(FindStr)
	Case 1
		'转换通配符为正则表达式模板
		TempList = ReSplit("\*,\#,\?,\[",",",-1)
		For i = 0 To UBound(TempList)
			FindStr = Replace$(FindStr,TempList(i),CStr$(i) & vbNullChar & CStr$(i) & vbNullChar & CStr$(i))
		Next i
		FindStr = Replace$(FindStr,"?",".")
		FindStr = Replace$(FindStr,"*",".*")
		FindStr = Replace$(FindStr,"#","\d")
		FindStr = Replace$(FindStr,"[!","[^")
		For i = 0 To UBound(TempList)
			FindStr = Replace(FindStr,CStr$(i) & vbNullChar & CStr$(i) & vbNullChar & CStr$(i),TempList(i))
		Next i
		FindStr = Replace$(FindStr,"\#","#")
	Case 2
		If CheckStrRegExp(FindStr,"(\\\(.+\\\))|(\\\(.+\))|(\(\?.+\))|(\(.+\).*\\[1-9]\d?)",0,2) = False Then
			If (FindStr Like "*(.*)*") = True Then Stemp = 3
		End If
	End Select
	'初始化正则表达式
	With RegExp
		.Global = Mode
		.IgnoreCase = False
		.Pattern = FindStr
	End With
	ptPos = GetCurPos(hwnd)	'获取光标位置  '.X 和 .Y 均以 1 为起点
	tmpPos = ptPos
	'查找字串
	With ptPos
	Do
		If .y > Lines Then
			.y = 0: .x = 0
		ElseIf .x < 0 Then
			.y = Lines: .x = 0
		End If
		StrText = GetCurPosLine(hwnd,ptPos)
		If StrText <> "" Then
			If Stemp = 0 Then
				If Mode = False Then
    				.x = InStr(.x + 1,StrText,FindStr)
    			Else
    				.x = InStrRev(StrText,FindStr,.x - 1)
    			End If
    			i = Length
    		Else
    			If Mode = False Then
    				StrText = Mid$(StrText,.x + 1)
    			ElseIf .x > 0 Then
					StrText = Left$(StrText,.x - 1)
				End If
				If StrText <> "" Then
   					Set Matches = RegExp.Execute(StrText)
					If Matches.Count > 0 Then
						If Mode = False Then
							If Stemp = 3 Then
								.x = .x + InStr(1,StrText,Matches(0).SubMatches(0))
								i = Len(Matches(0).SubMatches(0))
							Else
								.x = .x + Matches(0).FirstIndex + 1
								i = Matches(0).Length
							End If
						Else
							i = Matches.Count - 1
							If Stemp = 3 Then
								.x = InStrRev(StrText,Matches(i).SubMatches(0),-1)
								i = Len(Matches(i).SubMatches(0))
							Else
								.x = Matches(i).FirstIndex + 1
								i = Matches(i).Length
							End If
						End If
					Else
						.x = 0
					End If
				Else
					.x = 0
				End If
			End If
			If .x > 0 Then
				.x = .x - 1
				If .y = tmpPos.y And .x + i = tmpPos.x Then
					FindCurPos = -1
				Else
					FindCurPos = .y + 1
				End If
				Call SetCurPos(hwnd,ptPos,i)
				StrText = Mid$(GetCurPosLine(hwnd,ptPos),.x + 1,i)
				Exit Do
			Else
				.y = .y + IIf(Mode = False,1,-1)
			End If
		Else
			.y = .y + IIf(Mode = False,1,-1)
		End If
		n = n + 1
	Loop Until n > Lines + 1
	End With
	If FindCurPos = 0 Then StrText = ""
	Exit Function
	errHandle:
	StrText = ""
	FindCurPos = -2
End Function


'过滤字串
'Mode = 0 常规，= 1 通配符, = 2 正则表达式
'FilterStr = 1 已找到，= 0 未找到, = -2 查找内容语法
Public Function FilterStr(ByVal txtStr As String,ByVal FindStr As String,ByVal Mode As Long,Optional ByVal IgnoreCase As Boolean) As Long
	Dim i As Long,TempList() As String
	On Error GoTo errHandle
	Select Case Mode
	Case 0
		If (FindStr Like "*\[*?#[]*") = True Then
			TempList = ReSplit("*,?,#,[",",",-1)
			For i = 0 To UBound(TempList)
				FindStr = Replace$(FindStr,"\" & TempList(i),TempList(i))
			Next i
		End If
		If InStr(txtStr,FindStr) Then FilterStr = 1
	Case 1
		If (txtStr Like FindStr) = True Then FilterStr = 1
	Case 2
		If CheckStrRegExp(txtStr,FindStr,0,2,IgnoreCase) = True Then FilterStr = 1
	End Select
	Exit Function
	errHandle:
	FilterStr = -2
End Function


'插入文本到文本框光标所在开始处，返回 lpPoint 光标坐标
Public Function InsertStr(ByVal hwnd As Long,ByVal StrText As String,ByVal InsertText As String,lpPoint As POINTAPI) As String
	With lpPoint
		lpPoint = GetCurPos(hwnd)
		'SendMessage(hwnd, CB_GETEDITSEL, .x, .y)
		InsertStr = Left$(StrText,.x) & InsertText & Mid$(StrText,.x + 1)
		.x = .x + Len(InsertText)
	End With
	'Call SetCurPos(hwnd,pt,0)
End Function


'检查字体是否为空，非空返回 True
Public Function CheckFont(LF As LOG_FONT) As Boolean
	If ReSplit(StrConv$(LF.lfFaceName,vbUnicode),vbNullChar,2)(0) <> "" Then CheckFont = True
End Function


'获取字体名称和字号
Public Function GetFontText(ByVal hwnd As Long,LF As LOG_FONT) As String
	Dim LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd = 0 Then Exit Function
		GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
	End If
	GetFontText = ReSplit(StrConv$(LF2.lfFaceName,vbUnicode),vbNullChar,2)(0) & " " & CStr$(-LF2.lfHeight)
End Function


'比较二个字体数组是否相同，不相同返回 True
Public Function FontComps(LF() As LOG_FONT,LF2() As LOG_FONT) As Boolean
	Dim i As Long
	If UBound(LF) <> UBound(LF2) Then
		FontComps = True
		Exit Function
	End If
	For i = LBound(LF) To UBound(LF)
		If FontComp(LF(i),LF2(i)) = True Then
			FontComps = True
			Exit For
		End If
	Next i
End Function


'比较二个字体是否相同，不相同返回 True
Public Function FontComp(LF As LOG_FONT,LF2 As LOG_FONT) As Boolean
	FontComp = True
	With LF
		If .lfCharSet <> LF2.lfCharSet Then Exit Function
		If .lfClipPrecision <> LF2.lfClipPrecision Then Exit Function
		If .lfEscapement <> LF2.lfEscapement Then Exit Function
		If .lfFaceName <> LF2.lfFaceName Then Exit Function
		If .lfHeight <> LF2.lfHeight Then Exit Function
		If .lfItalic <> LF2.lfItalic Then Exit Function
		If .lfOrientation <> LF2.lfOrientation Then Exit Function
		If .lfOutPrecision <> LF2.lfOutPrecision Then Exit Function
		If .lfPitchAndFamily <> LF2.lfPitchAndFamily Then Exit Function
		If .lfQuality <> LF2.lfQuality Then Exit Function
		If .lfStrikeOut <> LF2.lfStrikeOut Then Exit Function
		If .lfUnderline <> LF2.lfUnderline Then Exit Function
		If .lfWeight <> LF2.lfWeight Then Exit Function
		If .lfWidth <> LF2.lfWidth Then Exit Function
		If .lfColor <> LF2.lfColor Then Exit Function
	End With
	FontComp = False
End Function


'弹出系统字体对话框选择字体，确定时返回非零
Public Function SelectFont(ByVal hwnd As Long,LF As LOG_FONT) As Long
	Dim CF As CHOOSE_FONT,LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd <> 0 Then
			GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
		End If
	End If
	With CF
		.lStructSize = Len(CF)			'size of structure
		.hwndOwner = hwnd				'window Form1 is opening this dialog box
		'.hDC = GetDC(hWnd)				'device context of default printer (using VB's mechanism)
		.lpLogFont = VarPtr(LF2)		'LogFont结构地址
		'.iPointSize = LF.lfHeight		'10 * size in points of selected font
		.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_INACTIVEFONTS
		.rgbColors = LF2.lfColor		'RGB(0,0,0)		'black
		'.lCustData = 0					' data passed to hook fn
		'.lpfnHook = 0					' ptr. to hook function
		'.lpTemplateName = ""			' custom template name
		'.hInstance = 0					' instance handle of.EXE that contains cust. dlg. template
		'.lpszStyle = LF2.lfFaceName	' return the style field here must be LF_FACESIZE or bigger
		.nFontType = LF2.lfWeight		'REGULAR_FONTTYPE	'regular font type i.e. not bold or anything
		.nSizeMin = 8 					'minimum point size
		.nSizeMax = 16 					'maximum point size
	End With
	SelectFont = ChooseFont(CF)
	If SelectFont = 0 Then Exit Function
	LF = LF2
	LF.lfColor = CF.rgbColors
End Function


'创建字体，返回字体句柄
Public Function CreateFont(ByVal hwnd As Long,LF As LOG_FONT) As Long
	Dim LF2 As LOG_FONT
	LF2 = LF
	If CheckFont(LF2) = False Then
		If hwnd = 0 Then Exit Function
		GetObjectAPI(SendMessageLNG(hwnd,WM_GETFONT,0,0),Len(LF2),VarPtr(LF2))
	End If
	CreateFont = CreateFontIndirect(LF2)
End Function


'重画整个对话框
Public Function DrawWindow(ByVal hwnd As Long,ByVal hFont As Long) As Long
	'Dim New_hFont As Long,hDC As Long
	'hDC = GetDC(hwnd)
	'New_hFont = SelectObject(hDC, hFont)
	'SendMessageLNG(hWnd,WM_SETREDRAW,True,0)
	SendMessageLNG(hwnd,WM_SETFONT,hFont,0)
	'SendMessageLNG(hWnd,WM_PAINT,0,0)
	DrawWindow = RedrawWindow(hwnd,0,0,RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
	'SelectObject hDC, New_hFont
	'DeleteObject(hFont)
	'ReleaseDC(hwnd, hDC)
End Function


'设置编辑控件中的最大文本长度，原最大长度为30000个字符（双字节字符算1个）
Public Function SetTextBoxLength(ByVal hwnd As Long,ByVal OldLength As Long,ByVal NewLength As Long,ByVal Mode As Boolean) As Long
	SetTextBoxLength = OldLength
	If NewLength < OldLength Then
		If Mode = False Then Exit Function
		If NewLength <= 25000 Then NewLength = 25000
	End If
	SetTextBoxLength = NewLength
	SendMessageLNG(hwnd,EM_LIMITTEXT,NewLength + 5000,0&)
End Function


'修正 PSL 2015 及以上版本宏引擎的 Split 函数拆分空字符串时返回未初始化数组的错误
Public Function ReSplit(ByVal TextStr As String,Optional ByVal Sep As String = " ",Optional ByVal Max As Integer = -1) As String()
	If TextStr = "" Then
		ReDim TempList(0) As String
		ReSplit = TempList
	Else
		ReSplit = Split(TextStr,Sep,Max)
	End If
End Function


'关于和帮助
Sub Help(ByVal HelpTip As String)
	Dim i As Long,MsgList() As String,HelpList(17) As String
	Dim Title As String,HelpTipTitle As String,HelpMsg As String

	For i = 0 To UBound(UIDataList)
		With UIDataList(i)
			Select Case .Title
			Case "Windows"
				MsgList = .Value
			Case "System"
				HelpList(0) = Replace$(Replace$(Join$(.Value,vbCrLf),"%s",Version),"%d",Build) & vbCrLf & vbCrLf
			Case "Description"
				HelpList(1) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Precondition"
				HelpList(2) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Setup"
				HelpList(3) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "CopyRight"
				HelpList(4) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Thank"
				HelpList(5) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Contact"
				HelpList(6) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "Logs"
				HelpList(7) = Join$(.Value,vbCrLf)
			Case "MainHelp"
				HelpList(8) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "EngineSetHelp"
				HelpList(9) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "EngineTestHelp"
				HelpList(10) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "CheckSetHelp"
				HelpList(11) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "ProjectHelp"
				HelpList(12) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "CheckTestHelp"
				HelpList(13) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "UpdateSetHelp"
				HelpList(14) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "UILangSetHelp"
				HelpList(15) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "EditFileHelp"
				HelpList(16) = Join$(.Value,vbCrLf) & vbCrLf & vbCrLf
			Case "RegExpRuleHelp"
				HelpList(17) = Join$(.Value,vbCrLf)
			End Select
		End With
	Next i

	Select Case HelpTip
	Case "About"
		HelpTipTitle = MsgList(2)
		HelpMsg = HelpList(0) & HelpList(1) & HelpList(2) & HelpList(3) & HelpList(4) & _
					HelpList(5) & HelpList(6) & HelpList(7)
	Case "MainHelp"
		HelpTipTitle = MsgList(3)
		HelpMsg = HelpList(8) & HelpList(7)
	Case "EngineSetHelp"
		HelpTipTitle = MsgList(4)
		HelpMsg = HelpList(9) & HelpList(7)
	Case "EngineTestHelp"
		HelpTipTitle = MsgList(5)
		HelpMsg = HelpList(10) & HelpList(7)
	Case "CheckSetHelp"
		HelpTipTitle = MsgList(6)
		HelpMsg = HelpList(11) & HelpList(7)
	Case "ProjectHelp"
		HelpTipTitle = MsgList(7)
		HelpMsg = HelpList(12) & HelpList(7)
	Case "CheckTestHelp"
		HelpTipTitle = MsgList(8)
		HelpMsg = HelpList(13) & HelpList(7)
	Case "UpdateSetHelp"
		HelpTipTitle = MsgList(9)
		HelpMsg = HelpList(14) & HelpList(7)
	Case "UILangSetHelp"
		HelpTipTitle = MsgList(10)
		HelpMsg = HelpList(15) & HelpList(7)
	Case "EditFileHelp"
		HelpTipTitle = MsgList(11)
		HelpMsg = Replace$(HelpList(16) & HelpList(7),"{RegExpRule}",HelpList(17))
	Case "RegExpRuleHelp"
		HelpTipTitle = MsgList(12)
		HelpMsg = HelpList(17) & vbCrLf & vbCrLf & HelpList(7)
	End Select

	Begin Dialog UserDialog 830,553,MsgList(0) & " - " & MsgList(1),.CommonDlgFunc ' %GRID:10,7,1,1
		Text 0,7,830,14,HelpTipTitle,.Text,2
		TextBox 0,28,830,490,.TextBox,1
		OKButton 370,525,100,21
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = HelpMsg
	Dialog dlg
End Sub

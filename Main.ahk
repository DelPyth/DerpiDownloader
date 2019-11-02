/*
	------------------------------------------------------
	Filename:				DerpiDownload.ahk
	Description:			Download a bunch of images from derpibooru.org with the tags of your liking.
	Version:				19w44a
	Created By:				Delta Pythagorean
	Author Email:			constantjarod@gmail.com
	Type:					AutoHotkey
	------------------------------------------------------
	TODO:
		+ Add a dialog to show images/webm with previews.
	------------------------------------------------------
*/

; Autoexec start
#NoEnv
#SingleInstance, Force
#KeyHistory, 0
#MaxThreadsPerHotkey, 1
#Persistent
SendMode, Input
SetBatchLines, -1
SetWinDelay, -1
SetMouseDelay, -1
SetKeyDelay, -1, -1
SetTitleMatchMode, 2
DetectHiddenWindows, On
SetWorkingDir, % A_ScriptDir

Global WM_USER               := 0x00000400
Global PBM_SETMARQUEE        := WM_USER + 10
Global PBM_SETSTATE          := WM_USER + 16
Global PBS_MARQUEE           := 0x00000008
Global PBS_SMOOTH            := 0x00000001
Global PBS_VERTICAL          := 0x00000004
Global PBST_NORMAL           := 0x00000001
Global PBST_ERROR            := 0x00000002
Global PBST_PAUSE            := 0x00000003
Global STAP_ALLOW_NONCLIENT  := 0x00000001
Global STAP_ALLOW_CONTROLS   := 0x00000002
Global STAP_ALLOW_WEBCONTENT := 0x00000004
Global WM_THEMECHANGED       := 0x0000031A

PonyList := "apple bloom|applejack|babs seed|berry punch|big macintosh|bon bon|braeburn|coco pommel|colgate|cutie mark crusaders|cutie mark|derpy hooves|diamond tiara|discord|doctor whooves|elements of harmony|fluttershy|granny smith|king sombra|lightning dust|mane six|maud pie|mayor mare|nightmare moon|nightmare night|octavia melody|opalescence|owlowiscious|pinkamena diane pie|pinkie pie|prince blueblood|princess cadance|princess celestia|princess flurry heart|princess luna|princess twilight|queen chrysalis|rainbow dash|rarity|royal guard|scootaloo|shining armor|silver spoon|soarin'|spike|spitfire|starlight glimmer|sunset shimmer|suri polomare|sweetie belle|trixie|twilight sparkle|vinyl scratch|zecora"

Menu, Gui, Add, Download..., StartDownload

Gui, New, +HwndGuiID -MinimizeBox
Gui, Menu, Gui
Gui, Add, GroupBox, x5 y5 w470 h335, Settings
Gui, Add, Text, x15 y25 w120 h20, Directory:
Gui, Add, Edit, x160 y25 w265 h21 vDir, % A_ScriptDir
Gui, Add, Button, x430 y25 w35 h21 gDirSelect, ...
Gui, Add, Text, x15 y50 w120 h20, Pony Name:
Gui, Add, ComboBox, x160 y50 w305 h200 vPonyName gCbAutoComplete, % PonyList
Gui, Add, Text, x15 y75 w120 h20, Amount Of Results:
Gui, Add, Edit, x160 y75 w305 h20 vMaxPerPage, 10
Gui, Add, UpDown, x440 y75 w17 h20 Range1-50
Gui, Add, Text, x15 y100 w120 h20, Filter
Gui, Add, ComboBox, x160 y100 w305 h200 gCbAutoComplete vFilter, Default - 100073||Everything - 56027|+18 Dark - 37429|Maximum Spoilers - 37430|+18 r34 - 37432|Legacy Default - 37431 ; 56027||37430|100073|37429|37432|37431|158584
Gui, Add, GroupBox, x15 y130 w450 h200, Tags
Gui, Add, Edit, x25 y145 w380 h21 vTag
Gui, Add, Button, x410 y145 w22 h22 +Default gAdd, +
Gui, Add, Button, x435 y145 w22 h22 gRemove, -
Gui, Add, ListBox, x25 y170 w430 h150 +0x100 vTagList
Gui, Add, Edit, x5 y345 w470 h165 +VScroll +Multi +ReadOnly vConsole
Gui, Add, Progress, x5 y515 w470 h15 vProg hwndMARQ -%PBS_SMOOTH%
Gui, Show, w480 h535, DerpiDownloader
GuiControl, Focus, PonyName

AddOn("Waiting for input...")
Return

GuiClose:
GuiEscape:
	ExitApp, 0

DirSelect:
	GuiControl,, Dir, % ((Folder := SelectFolder(A_ScriptDir, "Select A Folder To Save To", GuiID)) ? Folder : A_ScriptDir)
	Return

Add:
	Gui, Submit, NoHide
	If (Tag != "") {
		GuiControl,, Tag, % ""
		GuiControl,, TagList, % Tag "|"
	}
	Return

Remove:
	Gui, Submit, NoHide
	ControlGet, Content, List,, ListBox1, A
	TagList := StrReplace(Content, TagList)
	TagList := RegExReplace(TagList, "\n+", "|")
	GuiControl,, TagList, %  "|"
	GuiControl,, TagList, %  "|" TagList
	Return

StartDownload:
	Gui, Submit, NoHide
	MaxPages := 1
	Filter := RegExReplace(Filter, ".+\s-\s")

	; Get listbox results
	ControlGet, Content, List,, ListBox1, A

	AddOn("===========================================")
	AddOn("Close this dialog to cancel.")
	AddOn()
	AddOn("Checking for folders...")

	If (!FileExist(Dir))
		FileCreateDir, % Dir
	If (!FileExist(Dir "/" PonyName))
		FileCreateDir, % Dir "/" PonyName
	Tags := StrReplace(Content, "`n", "%2C+")
	Tags := StrReplace(Tags, " ", "+")
	Tags := StrReplace(Tags, ":", "%3A")
	Tags := StrExtra(Tags, "Lower")
	GuiControl,, Prog, 0
	Loop, % MaxPages
	{
		VarSetCapacity(Obj, 10240000)	; ~10mb for massive queries
		If (Tags != "") && (PonyName != "")
			Entry := "&q=" Tags "%2C+" PonyName
		If (Tags == "") && (PonyName != "")
			Entry := "&q=" PonyName
		If (Tags != "") && (PonyName == "")
			Entry := "&q=" Tags
		Str := "https://derpibooru.org/search.json?perpage=" MaxPerPage "&page=" A_Index "&filter_id=" Filter . Entry
		Obj := ReadFromPage(Str)
		X := JSON.Load(Obj)
		List := {}
		V_Index := A_Index
		MaxItems := MaxPerPage * MaxPages
		For Each, Item in X
		{
			For Index, Var in Item
				List[A_Index] := Var
		}

		For Each, Item in List {
			File := SplitPath("https:" X.Search[A_Index].Representations.Full)
			AddOn("Downloading... " X.Search[A_Index].ID "." File.Ext)
			If (PonyName != "")
				UrlDownloadToFile, % "https:" X.Search[A_Index].Representations.Full, % Dir "/" PonyName "/" X.Search[A_Index].ID "." File.Ext
			Else
				UrlDownloadToFile, % "https:" X.Search[A_Index].Representations.Full, % Dir "/RandomPonies/" X.Search[A_Index].ID "." File.Ext
			Prog(V_Index * Each)
		}
	}
	Prog(10)
	AddOn("Finished downloading your requested pool.")
	Return

AddOn(Txt := "") {
	GuiControlGet, Console
	GuiControl,, Console, % Console == "" ? Txt : Console "`n" Txt
}

Prog(Percent) {
	GuiControl,, Prog, % (Percent * 10)
}

SelectFolder(StartingFolder := "", Prompt := "", OwnerHwnd := 0, OkBtnLabel := "") {
	OwnerHwnd := WinExist(App.Name)
	Static OsVersion := DllCall("GetVersion", "UChar")
		  , IID_IShellItem := 0
		  , InitIID := VarSetCapacity(IID_IShellItem, 16, 0)
						& DllCall("Ole32.dll\IIDFromString", "WStr", "{43826d1e-e718-42ee-bc55-a1e261c37bfe}", "Ptr", &IID_IShellItem)
		  , Show := A_PtrSize * 3
		  , SetOptions := A_PtrSize * 9
		  , SetFolder := A_PtrSize * 12
		  , SetTitle := A_PtrSize * 17
		  , SetOkButtonLabel := A_PtrSize * 18
		  , GetResult := A_PtrSize * 20
	SelectedFolder := ""
	If (OsVersion < 6) { ; IFileDialog requires Win Vista+, so revert to FileSelectFolder
		FileSelectFolder, SelectedFolder, *%StartingFolder%, 3, %Prompt%
		Return SelectedFolder
	}
	OwnerHwnd := DllCall("IsWindow", "Ptr", OwnerHwnd, "UInt") ? OwnerHwnd : 0
	If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
		Return ""
	VTBL := NumGet(FileDialog + 0, "UPtr")
	; FOS_CREATEPROMPT | FOS_NOCHANGEDIR | FOS_PICKFOLDERS
	DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", 0x00002028, "UInt")
	If (StartingFolder <> "")
		If !DllCall("Shell32.dll\SHCreateItemFromParsingName", "WStr", StartingFolder, "Ptr", 0, "Ptr", &IID_IShellItem, "PtrP", FolderItem)
			DllCall(NumGet(VTBL + SetFolder, "UPtr"), "Ptr", FileDialog, "Ptr", FolderItem, "UInt")
	If (Prompt <> "")
		DllCall(NumGet(VTBL + SetTitle, "UPtr"), "Ptr", FileDialog, "WStr", Prompt, "UInt")
	If (OkBtnLabel <> "")
		DllCall(NumGet(VTBL + SetOkButtonLabel, "UPtr"), "Ptr", FileDialog, "WStr", OkBtnLabel, "UInt")
	If !DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", OwnerHwnd, "UInt") {
		If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt") {
			GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
			If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr) ; SIGDN_DESKTOPABSOLUTEPARSING
				SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
			ObjRelease(ShellItem)
	}  }
	If (FolderItem)
		ObjRelease(FolderItem)
	ObjRelease(FileDialog)
	Return SelectedFolder
}

;=======================================================================================
; Function:			SHAutoComplete
; Description:		Auto-completes typed values in an edit with various options.
; Usage:
;	Gui, Add, Edit, w200 h21 hwndEditCtrl1
;	SHAutoComplete(EditCtrl1)
;=======================================================================================
SHAutoComplete(hEdit, Option := 0x20000000) {
	; https://bit.ly/335nOYt		For more info on the function.
	DllCall("ole32\CoInitialize", "Uint", 0)
	; SHACF_AUTOSUGGEST_FORCE_OFF (0x20000000)
	;	Ignore the registry default and force the AutoSuggest feature off.
	;	This flag must be used in combination with one or more of the SHACF_FILESYS* or SHACF_URL* flags.

	; AKA. It won't autocomplete anything, but it will allow functionality such as Ctrl+Backspace deleting a word.
	DllCall("shlwapi\SHAutoComplete", "Uint", hEdit, "Uint", Option)
	DllCall("ole32\CoUninitialize")
}

;=======================================================================================
; Function:			CbAutoComplete
; Description:		Auto-completes typed values in a ComboBox.
;
; Author:			Pulover [Rodolfo U. Batista]
; Usage:
;	Gui, Add, ComboBox, w200 h50 gCbAutoComplete, Billy|Joel|Samual|Jim|Max|Jackson|George
;=======================================================================================
CbAutoComplete() {
	; CB_GETEDITSEL = 0x0140, CB_SETEDITSEL = 0x0142
	If ((GetKeyState("Delete", "P")) || (GetKeyState("Backspace", "P")))
		Return
	GuiControlGet, lHwnd, Hwnd, %A_GuiControl%
	SendMessage, 0x0140, 0, 0,, ahk_id %lHwnd%
	MakeShort(ErrorLevel, Start, End)
	GuiControlGet, CurContent,, %lHwnd%
	GuiControl, ChooseString, %A_GuiControl%, %CurContent%
	If (ErrorLevel) {
		ControlSetText,, %CurContent%, ahk_id %lHwnd%
		PostMessage, 0x0142, 0, MakeLong(Start, End),, ahk_id %lHwnd%
		Return
	}
	GuiControlGet, CurContent,, %lHwnd%
	PostMessage, 0x0142, 0, MakeLong(Start, StrLen(CurContent)),, ahk_id %lHwnd%
}

MakeLong(LoWord, HiWord) {
	return (HiWord << 16) | (LoWord & 0xffff)
}

MakeShort(Long, ByRef LoWord, ByRef HiWord) {
	LoWord := Long & 0xffff, HiWord := Long >> 16
}

ReadFromPage(Url) {
	Static hObject
	If (!(Url ~= "i)https?://"))
		Url := "https://" Url
	If (!IsObject(hObject))
		hObject := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	hObject.Open("GET", Url)
	hObject.Send()
	Return, hObject.ResponseText
}

SplitPath(File) {
	SplitPath, File, FileName, Dir, Ext, NameNoExt, Drive
	Return, {FileName: FileName, Dir: Dir, Ext: Ext, NameNoExt: NameNoExt, Drive: Drive}
}

StrExtra(Str, Command) {
	If (Command = "isInt") {
		; checks if the content is an integer
		if Str is integer
			out := true
		else
			out := false
	} Else If (Command ~= "i)capitalize|caps") {
		; Capitalizes first letter in the string
		out := Format("{1:U}{2:L}", Substr(Str, 1, 1), SubStr(Str, 2))
	} Else If (Command = "lower") {
		; converts string to lowercase
		out := Format("{:L}", Str)
	} Else If (Command = "title") {
		; Applies Title Case to the String
		out := Format("{:T}", Str)
	} Else If (Command = "upper") {
		; CONVERTS STRING TO UPPERCASE
		out := Format("{:U}", Str)
	} Else If (Command ~= "i)reverse|rev") {
		; reverses string from left->right to right->left
		; tfel>-thgir ot thgir>-tfel morf gnirts sesrever
		DllCall("msvcrt\_" (A_IsUnicode? "wcs":"str") "rev", "UInt", &Str, "CDecl")
		out := Str
	} Else If (Command = "i)length|len|size") {
		; returns the length of the string
		out := StrLen(Str)
	} Else If (Command = "isEmpty") {
		; checks if the string is empty
		out := StrLen(Str)? False:True
	} Else If (Command = "toHex") {
		; transforms a number into a hex value
		out := Format("{:#x}", Str)
	} Else If (Command = "toDec") {
		; transforms a hex value into a decimal value - needs to be prefixed with 0x; e.g. 0xFF
		out := Format("{:#d}", Str)
	} Else If (Command = "isStr") {
		; checks if variable is a string by converting it to hex.
		; if the conversion returns 0 instead of 0x it has to be a string
		out := SubStr(Format("{:#x}", Str), 1, 2)="0x"? False:True
	}
	Return, out
}

;=======================================================================================
; Class:			JSON
; Description:		Work with JSON styled information using AHK.
; Usage:
;	; Read from json file.
;	FileObj := FileOpen("file.json", "r")
;	Text := FileObj.Read()
;	BaseObject := JSON.Load(Text)
;	FileObj.Close()
;
;	; Save to json file.
;	Text := JSON.Dump(BaseObject, Func("ReplacerFunc"), "`t")
;	FileObj := FileOpen("file.json", "r")
;	FileObj.Write(Text)
;	FileObj.Close()
;=======================================================================================
class JSON
{
	class Load extends JSON.Functor
	{
		Call(self, ByRef text, reviver:="")
		{
			this.rev := IsObject(reviver) ? reviver : false
		; Object keys(and array indices) are temporarily stored in arrays so that
		; we can enumerate them in the order they appear in the document/text instead
		; of alphabetically. Skip if no reviver function is specified.
			this.keys := this.rev ? {} : false

			static quot := Chr(34), bashq := "\" . quot
			     , json_value := quot . "{[01234567890-tfn"
			     , json_value_or_array_closing := quot . "{[]01234567890-tfn"
			     , object_key_or_object_closing := quot . "}"

			key := ""
			is_key := false
			root := {}
			stack := [root]
			next := json_value
			pos := 0

			while ((ch := SubStr(text, ++pos, 1)) != "") {
				if InStr(" `t`r`n", ch)
					continue
				if !InStr(next, ch, 1)
					this.ParseError(next, text, pos)

				holder := stack[1]
				is_array := holder.IsArray

				if InStr(",:", ch) {
					next := (is_key := !is_array && ch == ",") ? quot : json_value

				} else if InStr("}]", ch) {
					ObjRemoveAt(stack, 1)
					next := stack[1]==root ? "" : stack[1].IsArray ? ",]" : ",}"

				} else {
					if InStr("{[", ch) {
					; Check if Array() is overridden and if its return value has
					; the 'IsArray' property. If so, Array() will be called normally,
					; otherwise, use a custom base object for arrays
						static json_array := Func("Array").IsBuiltIn || ![].IsArray ? {IsArray: true} : 0

					; sacrifice readability for minor(actually negligible) performance gain
						(ch == "{")
							? ( is_key := true
							  , value := {}
							  , next := object_key_or_object_closing )
						; ch == "["
							: ( value := json_array ? new json_array : []
							  , next := json_value_or_array_closing )

						ObjInsertAt(stack, 1, value)

						if (this.keys)
							this.keys[value] := []

					} else {
						if (ch == quot) {
							i := pos
							while (i := InStr(text, quot,, i+1)) {
								value := StrReplace(SubStr(text, pos+1, i-pos-1), "\\", "\u005c")

								static tail := A_AhkVersion<"2" ? 0 : -1
								if (SubStr(value, tail) != "\")
									break
							}

							if (!i)
								this.ParseError("'", text, pos)

							  value := StrReplace(value,  "\/",  "/")
							, value := StrReplace(value, bashq, quot)
							, value := StrReplace(value,  "\b", "`b")
							, value := StrReplace(value,  "\f", "`f")
							, value := StrReplace(value,  "\n", "`n")
							, value := StrReplace(value,  "\r", "`r")
							, value := StrReplace(value,  "\t", "`t")

							pos := i ; update pos

							i := 0
							while (i := InStr(value, "\",, i+1)) {
								if !(SubStr(value, i+1, 1) == "u")
									this.ParseError("\", text, pos - StrLen(SubStr(value, i+1)))

								uffff := Abs("0x" . SubStr(value, i+2, 4))
								if (A_IsUnicode || uffff < 0x100)
									value := SubStr(value, 1, i-1) . Chr(uffff) . SubStr(value, i+6)
							}

							if (is_key) {
								key := value, next := ":"
								continue
							}

						} else {
							value := SubStr(text, pos, i := RegExMatch(text, "[\]\},\s]|$",, pos)-pos)

							static number := "number", integer :="integer"
							if value is %number%
							{
								if value is %integer%
									value += 0
							}
							else if (value == "true" || value == "false")
								value := %value% + 0
							else if (value == "null")
								value := ""
							else
							; we can do more here to pinpoint the actual culprit
							; but that's just too much extra work.
								this.ParseError(next, text, pos, i)

							pos += i-1
						}

						next := holder==root ? "" : is_array ? ",]" : ",}"
					} ; If InStr("{[", ch) { ... } else

					is_array? key := ObjPush(holder, value) : holder[key] := value

					if (this.keys && this.keys.HasKey(holder))
						this.keys[holder].Push(key)
				}

			} ; while ( ... )

			return this.rev ? this.Walk(root, "") : root[""]
		}

		ParseError(expect, ByRef text, pos, len:=1)
		{
			static quot := Chr(34), qurly := quot . "}"

			line := StrSplit(SubStr(text, 1, pos), "`n", "`r").Length()
			col := pos - InStr(text, "`n",, -(StrLen(text)-pos+1))
			msg := Format("{1}`n`nLine:`t{2}`nCol:`t{3}`nChar:`t{4}"
			,     (expect == "")     ? "Extra data"
			    : (expect == "'")    ? "Unterminated string starting at"
			    : (expect == "\")    ? "Invalid \escape"
			    : (expect == ":")    ? "Expecting ':' delimiter"
			    : (expect == quot)   ? "Expecting object key enclosed in double quotes"
			    : (expect == qurly)  ? "Expecting object key enclosed in double quotes or object closing '}'"
			    : (expect == ",}")   ? "Expecting ',' delimiter or object closing '}'"
			    : (expect == ",]")   ? "Expecting ',' delimiter or array closing ']'"
			    : InStr(expect, "]") ? "Expecting JSON value or array closing ']'"
			    :                      "Expecting JSON value(string, number, true, false, null, object or array)"
			, line, col, pos)

			static offset := A_AhkVersion<"2" ? -3 : -4
			throw Exception(msg, offset, SubStr(text, pos, len))
		}

		Walk(holder, key)
		{
			value := holder[key]
			if IsObject(value) {
				for i, k in this.keys[value] {
					; check if ObjHasKey(value, k) ??
					v := this.Walk(value, k)
					if (v != JSON.Undefined)
						value[k] := v
					else
						ObjDelete(value, k)
				}
			}

			return this.rev.Call(holder, key, value)
		}
	}

	class Dump extends JSON.Functor
	{
		Call(self, value, replacer:="", space:="")
		{
			this.rep := IsObject(replacer) ? replacer : ""

			this.gap := ""
			if (space) {
				static integer := "integer"
				if space is %integer%
					Loop, % ((n := Abs(space))>10 ? 10 : n)
						this.gap .= " "
				else
					this.gap := SubStr(space, 1, 10)

				this.indent := "`n"
			}

			return this.Str({"": value}, "")
		}

		Str(holder, key)
		{
			value := holder[key]

			if (this.rep)
				value := this.rep.Call(holder, key, ObjHasKey(holder, key) ? value : JSON.Undefined)

			if IsObject(value) {
			; Check object type, skip serialization for other object types such as
			; ComObject, Func, BoundFunc, FileObject, RegExMatchObject, Property, etc.
				static type := A_AhkVersion<"2" ? "" : Func("Type")
				if (type ? type.Call(value) == "Object" : ObjGetCapacity(value) != "") {
					if (this.gap) {
						stepback := this.indent
						this.indent .= this.gap
					}

					is_array := value.IsArray
				; Array() is not overridden, rollback to old method of
				; identifying array-like objects. Due to the use of a for-loop
				; sparse arrays such as '[1,,3]' are detected as objects({}).
					if (!is_array) {
						for i in value
							is_array := i == A_Index
						until !is_array
					}

					str := ""
					if (is_array) {
						Loop, % value.Length() {
							if (this.gap)
								str .= this.indent

							v := this.Str(value, A_Index)
							str .= (v != "") ? v . "," : "null,"
						}
					} else {
						colon := this.gap ? ": " : ":"
						for k in value {
							v := this.Str(value, k)
							if (v != "") {
								if (this.gap)
									str .= this.indent

								str .= this.Quote(k) . colon . v . ","
							}
						}
					}

					if (str != "") {
						str := RTrim(str, ",")
						if (this.gap)
							str .= stepback
					}

					if (this.gap)
						this.indent := stepback

					return is_array ? "[" . str . "]" : "{" . str . "}"
				}

			} else ; is_number ? value : "value"
				return ObjGetCapacity([value], 1)=="" ? value : this.Quote(value)
		}

		Quote(string)
		{
			static quot := Chr(34), bashq := "\" . quot

			if (string != "") {
				  string := StrReplace(string,  "\",  "\\")
				; , string := StrReplace(string,  "/",  "\/") ; optional in ECMAScript
				, string := StrReplace(string, quot, bashq)
				, string := StrReplace(string, "`b",  "\b")
				, string := StrReplace(string, "`f",  "\f")
				, string := StrReplace(string, "`n",  "\n")
				, string := StrReplace(string, "`r",  "\r")
				, string := StrReplace(string, "`t",  "\t")

				static rx_escapable := A_AhkVersion<"2" ? "O)[^\x20-\x7e]" : "[^\x20-\x7e]"
				while RegExMatch(string, rx_escapable, m)
					string := StrReplace(string, m.Value, Format("\u{1:04x}", Ord(m.Value)))
			}

			return quot . string . quot
		}
	}

	Undefined[]
	{
		get {
			static empty := {}, vt_empty := ComObject(0, &empty, 1)
			return vt_empty
		}
	}

	class Functor
	{
		__Call(method, ByRef arg, args*)
		{
		; When casting to Call(), use a new instance of the "function object"
		; so as to avoid directly storing the properties(used across sub-methods)
		; into the "function object" itself.
			if IsObject(method)
				return (new this).Call(method, arg, args*)
			else if (method == "")
				return (new this).Call(arg, args*)
		}
	}
}

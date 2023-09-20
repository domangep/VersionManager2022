$PBExportHeader$n_pbupdater.sru
forward
global type n_pbupdater from datastore
end type
type filetime from structure within n_pbupdater
end type
type systemtime from structure within n_pbupdater
end type
type shellexecuteinfo from structure within n_pbupdater
end type
end forward

type filetime from structure
	unsignedlong		dwlowdatetime
	unsignedlong		dwhighdatetime
end type

type systemtime from structure
	unsignedinteger		wyear
	unsignedinteger		wmonth
	unsignedinteger		wdayofweek
	unsignedinteger		wday
	unsignedinteger		whour
	unsignedinteger		wminute
	unsignedinteger		wsecond
	unsignedinteger		wmilliseconds
end type

type shellexecuteinfo from structure
	long		cbsize
	long		fmask
	long		hwnd
	string		lpverb
	string		lpfile
	string		lpparameters
	string		lpdirectory
	long		nshow
	long		hinstapp
	long		lpidlist
	string		lpclass
	long		hkeyclass
	long		hicon
	long		hmonitor
	long		hprocess
end type

global type n_pbupdater from datastore
string dataobject = "d_filelist"
end type
global n_pbupdater n_pbupdater

type prototypes
Function ulong GetLastError( &
	) Library "kernel32.dll"

Function ulong FormatMessage( &
	ulong dwFlags, &
	ulong lpSource, &
	ulong dwMessageId, &
	ulong dwLanguageId, &
	Ref string lpBuffer, &
	ulong nSize, &
	ulong Arguments &
	) Library "kernel32.dll" Alias For "FormatMessageW"

Function long SHGetFolderPath ( &
	long hwndOwner, &
	long nFolder, &
	long hToken, &
	long dwFlags, &
	Ref string pszPath &
	) Library "shell32.dll" Alias For "SHGetFolderPathW"

Function boolean FileTimeToLocalFileTime ( &
	Ref FILETIME lpFileTime, &
	Ref FILETIME lpLocalFileTime &
	) Library "kernel32.dll" Alias For "FileTimeToLocalFileTime"

Function boolean FileTimeToSystemTime ( &
	Ref FILETIME lpFileTime, &
	Ref SYSTEMTIME lpSystemTime &
	) Library "kernel32.dll" Alias For "FileTimeToSystemTime"

Function boolean SystemTimeToFileTime ( &
	SYSTEMTIME lpSystemTime, &
	Ref FILETIME lpFileTime &
	) Library "kernel32.dll"

Function boolean GetFileTime ( &
	ulong hFile, &
	Ref FILETIME lpCreationTime, &
	Ref FILETIME lpLastAccessTime, &
	Ref FILETIME lpLastWriteTime &
	) Library "kernel32.dll"

Function boolean SetFileTime ( &
	ulong hFile, &
	FILETIME lpCreationTime, &
	FILETIME lpLastAccessTime, &
	FILETIME lpLastWriteTime &
	) Library "kernel32.dll"

Function ulong CreateFile ( &
	string lpFileName, &
	ulong dwDesiredAccess, &
	ulong dwShareMode, &
	ulong lpSecurityAttributes, &
	ulong dwCreationDisposition, &
	ulong dwFlagsAndAttributes, &
	ulong hTemplateFile &
	) Library "kernel32.dll" Alias For "CreateFileW"

Function boolean ShellExecuteEx ( &
	Ref SHELLEXECUTEINFO lpExecInfo &
	) Library "shell32.dll" Alias For "ShellExecuteExW"

Function boolean CloseHandle ( &
	ulong hObject &
	) Library "kernel32.dll"

Function long MessageBox ( &
	long hWnd, &
	string lpText, &
	string lpCaption, &
	ulong uType &
	) Library "user32.dll" Alias For "MessageBoxW"

end prototypes

type variables
Inet iinet_base
n_inetresult in_irdata
n_zlib in_zlib
Boolean ib_logfile
Boolean ib_UseZip = True
Integer ii_msgevent
Long il_hWnd
String is_company
String is_lasterrmsg
String is_location
String is_appdata
String is_urlroot
String is_vbscript
String is_zippassword

// of_Update return codes
Constant Long FAILURE   = -1
Constant Long NOUPDATE  = 1
Constant Long CANCELLED = 2
Constant Long COMPLETED = 3

// constants for CreateFile API function
Constant ULong INVALID_HANDLE_VALUE = -1
Constant ULong GENERIC_READ     = 2147483648
Constant ULong GENERIC_WRITE    = 1073741824
Constant ULong FILE_ATTRIBUTE_NORMAL = 128
Constant ULong FILE_SHARE_READ  = 1
Constant ULong FILE_SHARE_WRITE = 2
Constant ULong CREATE_NEW			= 1
Constant ULong CREATE_ALWAYS		= 2
Constant ULong OPEN_EXISTING		= 3
Constant ULong OPEN_ALWAYS			= 4
Constant ULong TRUNCATE_EXISTING = 5

// constants for MessageBox
Constant ULong MB_OK					= 0
Constant ULong MB_YESNO				= 4
Constant ULong MB_ICONHAND			= 16
Constant ULong MB_ICONQUESTION	= 32
Constant Long IDABORT	= 3
Constant Long IDCANCEL	= 2
Constant Long IDIGNORE	= 5
Constant Long IDNO		= 7
Constant Long IDOK		= 1
Constant Long IDRETRY	= 4
Constant Long IDYES		= 6

// constants for ShellExecuteEx
CONSTANT long SW_HIDDEN				= 0
CONSTANT long SW_SHOWNORMAL		= 1
CONSTANT long SW_SHOWMINIMIZED	= 2
CONSTANT long SW_SHOWMAXIMIZED	= 3

end variables

forward prototypes
public function datetime of_filedatetimetopb (filetime astr_filetime)
public function boolean of_getfiletime (string as_filename, ref datetime adt_create, ref datetime adt_access, ref datetime adt_update)
public subroutine of_start (long al_hwnd, integer ai_endevent, integer ai_msgevent)
public subroutine of_logfile (string as_msgtext)
public function string of_getappdata ()
public function boolean of_run (string as_program, string as_parms, string as_shellverb, long al_show)
public subroutine of_runscript ()
public function boolean of_isexecutable ()
public function string of_getlasterror ()
public subroutine of_seturlroot (string as_root)
public subroutine of_setpassword (string as_zippassword)
public subroutine of_setlogfile (boolean ab_logfile)
public subroutine of_setcompany (string as_company)
public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace)
public subroutine of_parse (string as_parse, string as_list, ref string as_array[])
public subroutine of_createlocalfolders (string as_filename)
public function long of_update ()
public subroutine of_deletelogfile ()
public subroutine of_pbdatetimetofile (datetime adt_filetime, ref filetime astr_filetime)
public function boolean of_setfiletime (string as_filename, datetime adt_updated)
public subroutine of_start (long al_hwnd)
public function string of_getappname ()
public function string of_stopmessage (long lparam)
public subroutine of_setlocation (string as_location)
public subroutine of_setusezip (boolean ab_usezip)
end prototypes

public function datetime of_filedatetimetopb (filetime astr_filetime);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_FileDateTimeToPB
//
// PURPOSE:    This function converts file datetime to UTC PB datetime.
//
// ARGUMENTS:  astr_filetime	- Filetime structure
//
// RETURN:     Datetime
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

SYSTEMTIME lstr_systime
DateTime ldt_filedate
String ls_time
Date ld_fdate
Time lt_ftime

SetNull(ldt_filedate)

If Not FileTimeToSystemTime(astr_filetime, &
			lstr_systime) Then Return ldt_filedate

ld_fdate = Date(lstr_systime.wYear, &
					lstr_systime.wMonth, lstr_systime.wDay)

ls_time = String(lstr_systime.wHour) + ":" + &
			 String(lstr_systime.wMinute) + ":" + &
			 String(lstr_systime.wSecond) + ":" + &
			 String(lstr_systime.wMilliseconds)
lt_ftime = Time(ls_Time)

ldt_filedate = DateTime(ld_fdate, lt_ftime)

Return ldt_filedate

end function

public function boolean of_getfiletime (string as_filename, ref datetime adt_create, ref datetime adt_access, ref datetime adt_update);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_GetFileTime
//
// PURPOSE:    This function gets file datetime attributes.
//
// ARGUMENTS:  as_filename	- File name
//					adt_create	- (By Ref) File Create datetime
//					adt_access	- (By Ref) Last access datetime
//					adt_update	- (By Ref) Last updated datetime
//
// RETURN:     True=Success, False=Failure
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

FILETIME lstr_create, lstr_access, lstr_update
Boolean lb_result
ULong lul_file

// open file for read
lul_file = CreateFile(as_filename, GENERIC_READ, &
					FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
If lul_file = INVALID_HANDLE_VALUE Then
	Return False
End If

// get the datetimes
lb_result = GetFileTime(lul_file, lstr_create, lstr_access, lstr_update)

// convert attributes to PB datetime
If lb_result Then
	adt_create = of_FileDateTimeToPB(lstr_create)
	adt_access = of_FileDateTimeToPB(lstr_access)
	adt_update = of_FileDateTimeToPB(lstr_update)
End If

// close the file
CloseHandle(lul_file)

Return lb_result

end function

public subroutine of_start (long al_hwnd, integer ai_endevent, integer ai_msgevent);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_Start
//
// PURPOSE:    This function is the main function called by the application.
//
// ARGUMENTS:  al_hWnd		- The application main window
//					ai_endevent	- The pbm_custom event called when complete
//					ai_msgevent	- The pbm_custom event called for microhelp updates
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_result

il_hWnd  = al_hWnd
ii_msgevent = ai_msgevent

ll_result = of_Update()

// let the window know process is done
Post(il_hWnd, 1023 + ai_endevent, 0, ll_result)

end subroutine

public subroutine of_logfile (string as_msgtext);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_LogFile
//
// PURPOSE:    This function writes messages to a logfile.
//
// ARGUMENTS:  as_msgtext	- Text to write to log file
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_logfile
Integer li_fnum

If ib_logfile Then
	ls_logfile = is_appdata + "updatelog.txt"
	li_fnum = FileOpen(ls_logfile, LineMode!, Write!)
	FileWrite(li_fnum, as_msgtext)
	FileClose(li_fnum)
End If

end subroutine

public function string of_getappdata ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_GetAppData
//
// PURPOSE:    This function finds the Application Data directory and adds
//					the company directory under it.
//
// RETURN:     Full path to ...\Application Data\company\
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

Constant Long SHGFP_TYPE_CURRENT = 0
Constant Long CSIDL_APPDATA = 26	// 0x001a
String ls_path
Long ll_rc

ls_path = Space(256)

ll_rc = SHGetFolderPath(il_hWnd, CSIDL_APPDATA, &
				0, SHGFP_TYPE_CURRENT, ls_path)

ls_path += "\" + is_company + "\"

Return ls_path

end function

public function boolean of_run (string as_program, string as_parms, string as_shellverb, long al_show);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_Run
//
// PURPOSE:    This function executes a program and optionally passes
//					parameters and a shell verb.
//
// ARGUMENTS:  as_program	- The program name to run.
//					as_parms		- The parameters to pass to the program.
//					al_show		- The show window option.
//
// RETURN:     True=Success, False=Failure
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

CONSTANT Long SEE_MASK_NOCLOSEPROCESS = 64
SHELLEXECUTEINFO lstr_sei
Long ll_ExitCode

// initialize structure values
lstr_sei.cbSize = 60
lstr_sei.fMask  = SEE_MASK_NOCLOSEPROCESS
lstr_sei.hWnd   = Handle(this)
lstr_sei.lpVerb = as_shellverb
lstr_sei.lpFile = as_program
lstr_sei.lpParameters = as_parms
lstr_sei.nShow  = al_show

If ShellExecuteEx(lstr_sei) Then
	// close process handle
	CloseHandle(lstr_sei.hProcess)
Else
	// return failure
	is_lasterrmsg = of_GetLastError()
	MessageBox( il_hWnd, "An error was returned by ShellExecuteEx:" + &
					"~r~n~r~n" + is_lasterrmsg, &
					of_GetAppName() + " - Program Update Error", MB_OK + MB_ICONHAND)
	Return False
End If

Return True

end function

public subroutine of_runscript ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_RunScript
//
// PURPOSE:    This function sets up call to of_Run to execute script.
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

Environment le_env
String ls_verb

// run update script
If FileExists(is_vbscript) Then
	GetEnvironment(le_env)
	If le_env.OSMajorRevision > 5 Then
		ls_verb = "runas"
	End If
	of_Run( "wscript.exe", "~"" + is_vbscript + "~"", &
				ls_verb, SW_SHOWNORMAL )
End If

end subroutine

public function boolean of_isexecutable ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_IsExecutable
//
// PURPOSE:    This function checks to see if running from IDE.
//
// RETURN:     True=Executable, False=IDE
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

Application la_app

la_app = GetApplication()

If Handle(la_app) = 0 Then
	Return False
Else
	Return True
End If

end function

public function string of_getlasterror ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_GetLastError
//
// PURPOSE:    This function returns the last Windows API error.
//
// RETURN:     Error message text
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

Constant ULong FORMAT_MESSAGE_FROM_SYSTEM = 4096
Constant ULong LANG_NEUTRAL = 0
String ls_buffer, ls_errmsg
ULong lul_error

lul_error = GetLastError()

ls_buffer = Space(255)

FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, &
		lul_error, LANG_NEUTRAL, ls_buffer, 255, 0)

ls_errmsg = "Error# " + String(lul_error) + "~r~n~r~n" + ls_buffer

Return ls_errmsg

end function

public subroutine of_seturlroot (string as_root);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetURLRoot
//
// PURPOSE:    This function sets the URL of the file directory.
//
// ARGUMENTS:  as_root	- URL of the download directory
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

is_urlroot = as_root

end subroutine

public subroutine of_setpassword (string as_zippassword);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetPassword
//
// PURPOSE:    This function sets the zipfile password.
//
// ARGUMENTS:  as_zippassword	- Password for the zip files
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

is_zippassword = as_zippassword

end subroutine

public subroutine of_setlogfile (boolean ab_logfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetLogFile
//
// PURPOSE:    This function sets the URL of the file directory.
//
// ARGUMENTS:  ab_logfile	- True=Create logfile
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

ib_logfile = ab_logfile

end subroutine

public subroutine of_setcompany (string as_company);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetCompany
//
// PURPOSE:    This function sets the application name used by messages.
//
// ARGUMENTS:  as_company	- Company name
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

is_company = as_company

end subroutine

public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_ReplaceAll
//
// PURPOSE:    This function replaces all occurrences of a string within
//					another string.
//
// ARGUMENTS:  as_oldstring	- The string to be updated
//					as_findstr		- The value being searched for
//					as_replace		- The value to replace with
//
// RETURN:     The updated string
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_newstring
Long ll_findstr, ll_replace, ll_pos

// get length of strings
ll_findstr = Len(as_findstr)
ll_replace = Len(as_replace)

// find first occurrence
ls_newstring = as_oldstring
ll_pos = Pos(ls_newstring, as_findstr)

Do While ll_pos > 0
	// replace old with new
	ls_newstring = Replace(ls_newstring, ll_pos, ll_findstr, as_replace)
	// find next occurrence
	ll_pos = Pos(ls_newstring, as_findstr, (ll_pos + ll_replace))
Loop

Return ls_newstring

end function

public subroutine of_parse (string as_parse, string as_list, ref string as_array[]);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_Parse
//
// PURPOSE:    This function parses a string into an array.
//
// ARGUMENTS:  as_parse	- The separator character
//					as_list	- The string to be separated
//					as_array	- (By Ref) The output array
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_pos, ll_cnt, ll_start
String ls_empty[]
Integer li_next

as_array = ls_empty
as_list = Trim(as_list)
If Right(as_list, 1) <> as_parse Then
	as_list = as_list + as_parse
End If

ll_start = 1
ll_pos = Pos(as_list, as_parse, ll_start)
do while ll_pos > 1
	li_next = UpperBound(as_array) + 1
	as_array[li_next] = Mid(as_list, ll_start, (ll_pos - ll_start))
	ll_start = ll_pos + 1
	ll_pos = Pos(as_list, as_parse, ll_start)
loop

end subroutine

public subroutine of_createlocalfolders (string as_filename);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_CreateLocalFolders
//
// PURPOSE:    This function creates any missing local folders.
//
// ARGUMENTS:  as_filename	- Name of the file that needs to be downloaded
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_parts[], ls_path
Integer li_idx, li_max

of_Parse("\", as_filename, ls_parts)

li_max = UpperBound(ls_parts) - 1
For li_idx = 1 To li_max
	If li_idx > 1 Then
		ls_path += "\"
	End If
	ls_path += ls_parts[li_idx]
	If Not DirectoryExists(ls_path) Then
		CreateDirectory(ls_path)
	End If
Next

end subroutine

public function long of_update ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_Update
//
// PURPOSE:    This function downloads files and creates the install script.
//
// RETURN:     -1		=	Error occurred
//					 1		=  No updates available
//					 2		=	Update request cancelled
//					 3		=	Updates complete
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// 11/15/2015	RolandS		Location must now be set by the caller
// -----------------------------------------------------------------------------

Integer li_fnum
DateTime ldt_filedate, ldt_create, ldt_access, ldt_update
String ls_errmsg, ls_data, ls_filename, ls_fullname
String ls_zipfile, ls_target, ls_source
String ls_action, ls_inzipname, ls_appname
Long ll_row, ll_max, ll_rtn
ULong lul_unzFile

// make sure location set
If is_location = "" Then
	MessageBox("Auto Update", "Location not set!")
	is_lasterrmsg = "Location not set!"
	of_LogFile(is_lasterrmsg)
	Return FAILURE
End If

// get the temp path
is_appdata = of_GetAppData()

// get the application name
ls_appname = of_GetAppName()

// initialize logfile
of_DeleteLogFile()
of_LogFile("Date " + String(DateTime(Today(),Now())) + "~r~n")
of_LogFile("Installation Directory: " + is_location + "~r~n")

// set script name and delete any existing file
is_vbscript = is_appdata + "copyfiles.vbs"
If FileExists(is_vbscript) Then
	FileDelete(is_vbscript)
End If

// get the file list from website
in_irdata.is_filename = ""
in_irdata.is_filedata = ""
iinet_base.GetURL(is_urlroot + "/filelist.txt", in_irdata)
ls_data = in_irdata.is_filedata
If ls_data = "" Or Pos(ls_data, "<body>") > 0 Then
	is_lasterrmsg = "Get FileList failed"
	of_LogFile(is_lasterrmsg)
	Return FAILURE
End If

// import the file list
this.ImportString(ls_data)
this.Sort()

of_LogFile("Checking files...")

// check the files
ll_max = this.RowCount()
For ll_row = 1 To ll_max
	ls_action = this.GetItemString(ll_row, "action")
	// check the files to see if any are out of date
	If ls_action = "cpy" Then
		ls_filename  = this.GetItemString(ll_row, "filename")
		ls_fullname = is_location + "\" + ls_filename
		of_LogFile(ls_fullname)
		If FileExists(ls_fullname) Then
			ldt_filedate = this.GetItemDateTime(ll_row, "filedate")
			If Not of_GetFileTime(ls_fullname, ldt_create, ldt_access, ldt_update) Then
				is_lasterrmsg = "An error occurred getting the datetime for " + ls_filename
				of_LogFile(is_lasterrmsg)
				Return FAILURE
			End If
			If ldt_filedate = ldt_update Then
				of_LogFile("Dates are the same: " + String(ldt_filedate, "mm/dd/yyyy hh:mm:ss.fff"))
				// set action to blank so it can get filtered out
				this.SetItem(ll_row, "action", "")
			Else
				If ldt_filedate < ldt_update Then
					ls_errmsg  = "Installed date greater -"
					ls_errmsg += " Current: " + String(ldt_filedate, "mm/dd/yyyy hh:mm:ss.fff")
					ls_errmsg += " Installed: " + String(ldt_update, "mm/dd/yyyy hh:mm:ss.fff")
					of_LogFile(ls_errmsg)
					// set action to blank so it can get filtered out
					this.SetItem(ll_row, "action", "")
				Else
					ls_errmsg  = "Dates are different -"
					ls_errmsg += " Current: " + String(ldt_filedate, "mm/dd/yyyy hh:mm:ss.fff")
					ls_errmsg += " Installed: " + String(ldt_update, "mm/dd/yyyy hh:mm:ss.fff")
					of_LogFile(ls_errmsg)
				End If
			End If
		Else
			of_LogFile("File does not exist!")
		End If
	End If
	// check the delete files to see if they exist
	If ls_action = "del" Then
		ls_filename  = this.GetItemString(ll_row, "filename")
		ls_fullname = is_location + "\" + ls_filename
		If Not FileExists(ls_fullname) Then
			// set action to blank so it can get filtered out
			this.SetItem(ll_row, "action", "")
		End If
	End If
Next

// filter out unchanged files
this.Filter()

// return if no updates
If this.RowCount() = 0 Then
	of_LogFile("No updates found")
	Return NOUPDATE
End If

// set microhelp to Ready
Post(il_hWnd, 1023 + ii_msgevent, 0, 0)

// verify download
ll_rtn = MessageBox( il_hWnd, "A new version of the program is available." + &
					"~r~n~r~n" + "Do you wish to download updates?", &
					ls_appname + " - Program Update", MB_YESNO + MB_ICONQUESTION)
If ll_rtn = IDNO Then
	is_lasterrmsg = ""
	of_LogFile("Download Cancelled")
	Return CANCELLED
End If

// set zipfile password
in_zlib.of_SetPassword(is_zippassword)

of_LogFile("Downloading files...")

// download the files
ll_max = this.RowCount()
For ll_row = 1 To ll_max
	// display file count in microhelp
	Post(il_hWnd, 1023 + ii_msgevent, ll_row, ll_max)
	// get action
	ls_action = this.GetItemString(ll_row, "action")
	If ls_action = "cpy" Then
		If ib_UseZip Then
			// build the zipfile and target filenames
			ls_filename  = this.GetItemString(ll_row, "filename")
			ls_zipfile = ls_filename + ".zip"
			ls_target = is_appdata + ls_zipfile
			ls_zipfile = of_ReplaceAll(ls_zipfile, "\", "/")
			// make sure the local directory structure exists
			of_CreateLocalFolders(ls_target)
			// download the file
			of_LogFile("Downloading Zipfile: " + ls_zipfile)
			in_irdata.is_filename = ls_target
			iinet_base.GetURL(is_urlroot + "/" + ls_zipfile, in_irdata)
			// download worked if file exists
			If FileExists(ls_target) Then
				of_LogFile("Target found: "  + ls_target)
				// uncompress the file
				lul_unzFile = in_zlib.of_unzOpen(ls_target)
				ls_target = is_appdata + ls_filename
				ls_inzipname = Mid(ls_filename, LastPos(ls_filename, "\") + 1)
				in_zlib.of_ExtractFile(lul_unzFile, ls_target, ls_inzipname)
				in_zlib.of_unzClose(lul_unzFile)
				// set the file datetime
				ldt_filedate = this.GetItemDateTime(ll_row, "filedate")
				of_SetFileTime(ls_target, ldt_filedate)
				// delete the zip file
				ls_target = is_appdata + ls_zipfile
				FileDelete(ls_target)
			Else
				of_LogFile("Target not found: " + ls_target)
			End If
		Else
			// build the target filenames
			ls_filename = this.GetItemString(ll_row, "filename")
			ls_target = is_appdata + ls_filename
			ls_filename = of_ReplaceAll(ls_filename, "\", "/")
			// make sure the local directory structure exists
			of_CreateLocalFolders(ls_target)
			// download the file
			of_LogFile("Downloading File: " + ls_filename)
			in_irdata.is_filename = ls_target
			iinet_base.GetURL(is_urlroot + "/" + ls_filename, in_irdata)
			// download worked if file exists
			If FileExists(ls_target) Then
				of_LogFile("Target found: "  + ls_target)
				// set the file datetime
				ldt_filedate = this.GetItemDateTime(ll_row, "filedate")
				of_SetFileTime(ls_target, ldt_filedate)
			Else
				of_LogFile("Target not found: " + ls_target)
			End If
		End If
	End If
Next

of_LogFile("Creating the vbscript...")

// create the file copy script
li_fnum = FileOpen(is_vbscript, LineMode!, Write!)

// write the start section
ls_data = "Dim FSO, WshShell, BtnCode"
FileWrite(li_fnum, ls_data)
FileWrite(li_fnum, "")
ls_data = "WScript.Sleep 3000"
FileWrite(li_fnum, ls_data)
FileWrite(li_fnum, "")
ls_data = "Set FSO = CreateObject(~"Scripting.FileSystemObject~")"
FileWrite(li_fnum, ls_data)

// write the file copy actions
ll_max = this.RowCount()
If ll_max > 0 Then
	FileWrite(li_fnum, "")
	For ll_row = 1 To ll_max
		ls_action = this.GetItemString(ll_row, "action")
		ls_filename = this.GetItemString(ll_row, "filename")
		ls_source = is_appdata + ls_filename
		ls_target = is_location + "\" + ls_filename
		ls_data = ""
		If ls_action = "cpy" Then
			If FileExists(ls_source) Then
				of_LogFile("Copy: "  + ls_source)
				// copy new file
				ls_data  = "FSO.CopyFile "
				ls_data += "~"" + ls_source + "~", "
				ls_data += "~"" + ls_target + "~""
				FileWrite(li_fnum, ls_data)
			Else
				of_LogFile("Source file not found: "  + ls_source)
			End If
		End If
	Next
End If

// write the file delete actions
ll_max = this.RowCount()
If ll_max > 0 Then
	FileWrite(li_fnum, "")
	For ll_row = 1 To ll_max
		ls_action = this.GetItemString(ll_row, "action")
		ls_filename = this.GetItemString(ll_row, "filename")
		ls_source = is_appdata + ls_filename
		ls_target = is_location + "\" + ls_filename
		ls_data = ""
		choose case ls_action
			case "cpy"
				If FileExists(ls_source) Then
					// delete new file from the temp folder
					ls_data  = "FSO.DeleteFile "
					ls_data += "~"" + ls_source + "~""
					FileWrite(li_fnum, ls_data)
				End If
			case "del"
				// delete existing file from installation folder
				If FileExists(ls_target) Then
					ls_data  = "FSO.DeleteFile "
					ls_data += "~"" + ls_target + "~""
					FileWrite(li_fnum, ls_data)
				End If
		end choose
	Next
End If

// write the end section
FileWrite(li_fnum, "")
ls_data = "Set WshShell = WScript.CreateObject(~"WScript.Shell~")"
FileWrite(li_fnum, ls_data)
FileWrite(li_fnum, "")
ls_data  = "BtnCode = WshShell.Popup(~"Installation of updated "
ls_data += "program files is complete.~", 5, ~""
ls_data += ls_appname + " - Program Update~", 0)"
FileWrite(li_fnum, ls_data)

FileClose(li_fnum)

of_LogFile("Updates Complete")

Return COMPLETED

end function

public subroutine of_deletelogfile ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_DeleteLogFile
//
// PURPOSE:    This function deletes the logfile.
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_logfile

ls_logfile = is_appdata + "updatelog.txt"

If FileExists(ls_logfile) Then
	FileDelete(ls_logfile)
End If

end subroutine

public subroutine of_pbdatetimetofile (datetime adt_filetime, ref filetime astr_filetime);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_PBDateTimeToFile
//
// PURPOSE:    This function converts PB datetime to UTC file datetime.
//
// ARGUMENTS:  adt_filetime	- PowerBuilder DateTime
//					astr_filetime	- (By Ref) Filetime structure
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

SYSTEMTIME lstr_systime
String ls_time
Date ld_fdate
Time lt_ftime

ld_fdate = Date(adt_filetime)
lstr_systime.wYear  = Year(ld_fdate)
lstr_systime.wMonth = Month(ld_fdate)
lstr_systime.wDay   = Day(ld_fdate)

ls_time = String(adt_filetime, "hh:mm:ss.fff")
lstr_systime.wHour   = Long(Left(ls_time, 2))
lstr_systime.wMinute = Long(Mid(ls_time, 4, 2))
lstr_systime.wSecond = Long(Mid(ls_time, 7, 2))
lstr_systime.wMilliseconds = Long(Mid(ls_time, 10, 3))

SystemTimeToFileTime(lstr_systime, astr_filetime)

end subroutine

public function boolean of_setfiletime (string as_filename, datetime adt_updated);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetFileTime
//
// PURPOSE:    This function sets file datetime attributes.
//
// ARGUMENTS:  as_filename	- File name
//					adt_updated	- Last updated datetime
//
// RETURN:     True=Success, False=Failure
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

FILETIME lstr_updated
ULong lul_file

of_PBDateTimeToFile(adt_updated, lstr_updated)

// open file for update
lul_File = CreateFile(as_FileName, GENERIC_WRITE, &
					FILE_SHARE_READ, 0, OPEN_EXISTING, &
					FILE_ATTRIBUTE_NORMAL, 0)
If lul_file = INVALID_HANDLE_VALUE Then
	Return False
End If

// set the filetimes
If Not SetFileTime(lul_file, lstr_updated, &
				lstr_updated, lstr_updated) Then
	CloseHandle(lul_file)
	Return False
End If

// close the file
CloseHandle(lul_file)

Return True

end function

public subroutine of_start (long al_hwnd);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_Start
//
// PURPOSE:    This function is the main function called by the application.
//					The event arguments default to pbm_custom01 and pbm_custom02.
//
// ARGUMENTS:  al_hWnd		- The application main window
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

of_Start(al_hWnd, 1, 2)

end subroutine

public function string of_getappname ();// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_GetAppName
//
// PURPOSE:    This function returns the application name.
//
// RETURN:     Applicaton name.
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_appname
Application la_app

la_app = GetApplication()
If la_app.DisplayName = "" Then
	ls_appname = WordCap(la_app.AppName)
Else
	ls_appname = la_app.DisplayName
End If

Return ls_appname

end function

public function string of_stopmessage (long lparam);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_StopMessage
//
// PURPOSE:    This function returns a message to be displayed when the
//					update process was stopped.
//
// ARGUMENTS:  lparam	- The code passed to ue_updatefinish
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/24/2012  RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_msgtext

choose case lparam
	case FAILURE
		ls_msgtext = "Program Update - Errors occurred."
	case NOUPDATE
		ls_msgtext = "Program Update - No updates available."
	case CANCELLED
		ls_msgtext = "Program Update - Download cancelled."
	case else
		ls_msgtext = "Ready"
end choose

Return ls_msgtext

end function

public subroutine of_setlocation (string as_location);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetLocation
//
// PURPOSE:    This function sets the application installation folder.
//
// ARGUMENTS:  as_location	- Installation folder
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 11/15/2015  RolandS		Initial creation
// -----------------------------------------------------------------------------

is_location = as_location

end subroutine

public subroutine of_setusezip (boolean ab_usezip);// -----------------------------------------------------------------------------
// SCRIPT:     n_pbupdater.of_SetUseZip
//
// PURPOSE:    This function sets the UseZip option.
//
// ARGUMENTS:  ab_usezip	- True=Use Zip files, False=Do not use Zip files
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 01/06/2018  RolandS		Initial creation
// -----------------------------------------------------------------------------

ib_usezip = ab_usezip

end subroutine

on n_pbupdater.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_pbupdater.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

event constructor;// instantiate internet service objects
this.GetContextService("Internet", iinet_base)
in_irdata = CREATE n_inetresult

end event


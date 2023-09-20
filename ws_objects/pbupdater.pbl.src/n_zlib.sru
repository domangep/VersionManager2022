$PBExportHeader$n_zlib.sru
forward
global type n_zlib from nonvisualobject
end type
type filetime from structure within n_zlib
end type
type win32_find_data from structure within n_zlib
end type
type systemtime from structure within n_zlib
end type
type tm_unz from structure within n_zlib
end type
type unz_fileinfo from structure within n_zlib
end type
type shfileinfo from structure within n_zlib
end type
type zip_fileinfo from structure within n_zlib
end type
end forward

type filetime from structure
	unsignedlong		dwlowdatetime
	unsignedlong		dwhighdatetime
end type

type win32_find_data from structure
	unsignedlong		dwfileattributes
	filetime		ftcreationtime
	filetime		ftlastaccesstime
	filetime		ftlastwritetime
	unsignedlong		nfilesizehigh
	unsignedlong		nfilesizelow
	unsignedlong		dwreserved0
	unsignedlong		dwreserved1
	character		cfilename[260]
	character		calternatefilename[14]
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

type tm_unz from structure
	unsignedinteger		tm_sec
	unsignedinteger		tm_gap1
	unsignedinteger		tm_min
	unsignedinteger		tm_gap2
	unsignedinteger		tm_hour
	unsignedinteger		tm_gap3
	unsignedinteger		tm_mday
	unsignedinteger		tm_gap4
	unsignedinteger		tm_mon
	unsignedinteger		tm_gap5
	unsignedinteger		tm_year
	unsignedinteger		tm_gap6
end type

type unz_fileinfo from structure
	unsignedlong		version
	unsignedlong		version_needed
	unsignedlong		flag
	unsignedlong		compression_method
	unsignedlong		dosdate
	unsignedlong		crc
	unsignedlong		compressed_size
	unsignedlong		uncompressed_size
	unsignedlong		size_filename
	unsignedlong		size_file_extra
	unsignedlong		size_file_comment
	unsignedlong		disk_num_start
	unsignedlong		internal_fa
	unsignedlong		external_fa
	tm_unz		tmu_date
end type

type shfileinfo from structure
	long		hicon
	long		iicon
	long		dwattributes
	character		szdisplayname[260]
	character		sztypename[80]
end type

type zip_fileinfo from structure
	tm_unz		tmz_date
	unsignedlong		dos_date
	unsignedlong		internal_fa
	unsignedlong		external_fa
end type

global type n_zlib from nonvisualobject autoinstantiate
end type

type prototypes
// ZLib API functions

Function ulong crc32 ( &
	ulong crc, &
	Ref blob buf, &
	ulong len &
	) Library "zlibwapi.dll"

Function long compress ( &
	Ref blob dest, &
	Ref ulong destLen, &
	Ref blob source, &
	ulong sourceLen &
	) Library "zlibwapi.dll"

Function long compress2 ( &
	Ref blob dest, &
	Ref ulong destLen, &
	Ref blob source, &
	ulong sourceLen, &
	long level &
	) Library "zlibwapi.dll"

Function long uncompress ( &
	Ref blob dest, &
	Ref ulong destLen, &
	Ref blob source, &
	ulong sourceLen &
	) Library "zlibwapi.dll"

Function long uncompress2 ( &
	Ref blob dest, &
	Ref ulong destLen, &
	Ref blob source, &
	Ref ulong sourceLen &
	) Library "zlibwapi.dll"

Function string zlibVersion ( &
	) Library "zlibwapi.dll" Alias For "zlibVersion;Ansi"

Function long unzOpen ( &
	string filename &
	) Library "zlibwapi.dll" Alias For "unzOpen;Ansi"

Function long unzClose ( &
	long unzFile &
	) Library "zlibwapi.dll"

Function long unzGetGlobalComment ( &
	long unzFile, &
	Ref string szComment, &
	ulong uSizeBuf &
	) Library "zlibwapi.dll" Alias For "unzGetGlobalComment;Ansi"

Function long unzGoToFirstFile ( &
	long unzFile &
	) Library "zlibwapi.dll"

Function long unzGoToNextFile ( &
	long unzFile &
	) Library "zlibwapi.dll"

Function long unzGetCurrentFileInfo ( &
	long unzFile, &
	Ref unz_fileinfo file_info, &
	Ref string szFileName, &
	ulong filenamesize, &
	Ref ulong extrafield, &
	ulong extrasize, &
	Ref string comment, &
	ulong cmtsize &
	) Library "zlibwapi.dll" Alias For "unzGetCurrentFileInfo;Ansi"

Function long unzLocateFile ( &
	long unzFile, &
	string szFileName, &
	int iCase &
	) Library "zlibwapi.dll" Alias For "unzLocateFile;Ansi"

Function long unzOpenCurrentFile ( &
	long unzFile &
	) Library "zlibwapi.dll"

Function long unzCloseCurrentFile ( &
	long unzFile &
	) Library "zlibwapi.dll"

Function long unzReadCurrentFile ( &
	long unzFile, &
	Ref blob data, &
	ulong dataLen &
	) Library "zlibwapi.dll"

Function long unzOpenCurrentFilePassword ( &
	long unzFile, &
	Ref string password &
	) Library "zlibwapi.dll" Alias For "unzOpenCurrentFilePassword;Ansi"

Function long zipOpen ( &
	Ref string filename, &
	int append &
	) Library "zlibwapi.dll" Alias For "zipOpen;Ansi"

Function long zipClose ( &
	long zipFile, &
	Ref string szComment &
	) Library "zlibwapi.dll" Alias For "zipClose;Ansi"

Function long zipOpenNewFileInZip ( &
	long file, &
	Ref string filename, &
	Ref zip_fileinfo zipfi, &
	Ref ulong extrafield_local, &
	uint size_extrafield_local, &
	Ref ulong extrafield_global, &
	uint size_extrafield_global, &
	Ref string comment, &
	int method, &
	int level &
	) Library "zlibwapi.dll" Alias For "zipOpenNewFileInZip;Ansi"

Function long zipOpenNewFileInZip3 ( &
	long file, &
	Ref string filename, &
	Ref zip_fileinfo zipfi, &
	Ref ulong extrafield_local, &
	uint size_extrafield_local, &
	Ref ulong extrafield_global, &
	uint size_extrafield_global, &
	Ref string comment, &
	int method, &
	int level, &
	int raw, &
	int windowBits, &
	int memLevel, &
	int strategy, &
	Ref string password, &
	ulong crcForCrypting &
	) Library "zlibwapi.dll" Alias For "zipOpenNewFileInZip3;Ansi"

Function long zipWriteInFileInZip ( &
	long zipFile, &
	Ref blob buffer, &
	long uSizeBuf &
	) Library "zlibwapi.dll"

Function long zipCloseFileInZip ( &
	long zipFile &
	) Library "zlibwapi.dll"

Function long gzclose ( &
	long gzFile &
	) Library "zlibwapi.dll"

Function long gzopen ( &
	string filename, &
	string mode &
	) Library "zlibwapi.dll" Alias For "gzopen;Ansi"

Function long gzread ( &
	long gzFile, &
	Ref blob data, &
	ulong dataLen &
	) Library "zlibwapi.dll"

Function long gzwrite ( &
	long gzFile, &
	Ref blob buffer, &
	ulong uSizeBuf &
	) Library "zlibwapi.dll"

// Win API functions
Function long CreateFile ( &
	string lpFileName, &
	ulong dwDesiredAccess, &
	ulong dwShareMode, &
	ulong lpSecurityAttributes, &
	ulong dwCreationDisposition, &
	ulong dwFlagsAndAttributes, &
	ulong hTemplateFile &
	) Library "kernel32.dll" Alias For "CreateFileW"

Function boolean ReadFile ( &
	long hFile, &
	Ref blob lpBuffer, &
	ulong nNumberOfBytesToRead, &
	Ref ulong lpNumberOfBytesRead, &
	ulong lpOverlapped &
	) Library "kernel32.dll"

Function boolean WriteFile ( &
	long hFile, &
	blob lpBuffer, &
	ulong nNumberOfBytesToWrite, &
	Ref ulong lpNumberOfBytesWritten, &
	ulong lpOverlapped &
	) Library "kernel32.dll"

Function boolean CloseHandle ( &
	long hObject &
	) Library "kernel32.dll"

Function long FindFirstFile ( &
	Ref string filename, &
	Ref win32_find_data findfiledata &
	) Library "kernel32.dll" Alias For "FindFirstFileW"

Function boolean FindNextFile ( &
	ulong handle, &
	Ref win32_find_data findfiledata &
	) Library "kernel32.dll" Alias For "FindNextFileW"

Function boolean FindClose ( &
	ulong handle &
	) Library "kernel32.dll" Alias For "FindClose"

Function boolean FileTimeToLocalFileTime ( &
	Ref filetime lpFileTime, &
	Ref filetime lpLocalFileTime &
	) Library "kernel32.dll" Alias For "FileTimeToLocalFileTime"

Function boolean FileTimeToSystemTime ( &
	Ref filetime lpFileTime, &
	Ref systemtime lpSystemTime &
	) Library "kernel32.dll" Alias For "FileTimeToSystemTime"

Function long FindMimeFromData ( &
	long pBC, &
	string pwzUrl, &
	blob pBuffer, &
	ulong cbSize, &
	long pwzMimeProposed, &
	ulong dwMimeFlags, &
	Ref long ppwzMimeOut, &
	ulong dwReserved &
	) Library "urlmon.dll"

Function ulong SHGetFileInfo ( &
	string pszPath, &
	long dwFileAttributes, &
	Ref SHFILEINFO psfi, &
	long cbFileInfo, &
	long uFlags &
	) Library "shell32.dll" Alias For "SHGetFileInfoW"

Subroutine CopyMemory ( &
	Ref blob Destination, &
	ulong Source, &
	long Length &
	) Library  "kernel32.dll" Alias For "RtlMoveMemory"

Function ulong GetLastError( ) Library "kernel32.dll"

Function ulong FormatMessage( &
	ulong dwFlags, &
	ulong lpSource, &
	ulong dwMessageId, &
	ulong dwLanguageId, &
	Ref string lpBuffer, &
	ulong nSize, &
	ulong Arguments &
	) Library "kernel32.dll" Alias For "FormatMessageW"

end prototypes

type variables
String is_password

// Return codes for the compression/decompression functions. Negative values
// are errors, positive values are used for special but normal events.
Constant Long Z_OK				= 0
Constant Long Z_STREAM_END		= 1
Constant Long Z_NEED_DICT		= 2
Constant Long Z_ERRNO			= -1
Constant Long Z_STREAM_ERROR	= -2
Constant Long Z_DATA_ERROR		= -3
Constant Long Z_MEM_ERROR		= -4
Constant Long Z_BUF_ERROR		= -5
Constant Long Z_VERSION_ERROR	= -6

// compression levels
Constant Long Z_NO_COMPRESSION		= 0
Constant Long Z_BEST_SPEED				= 1
Constant Long Z_BEST_COMPRESSION		= 9
Constant Long Z_DEFAULT_COMPRESSION	= -1

// compression strategy
Constant Long Z_FILTERED			= 1
Constant Long Z_HUFFMAN_ONLY		= 2
Constant Long Z_RLE					= 3
Constant Long Z_FIXED				= 4
Constant Long Z_DEFAULT_STRATEGY	= 0

// Possible values of the data_type field
Constant Long Z_BINARY	= 0
Constant Long Z_TEXT		= 1
Constant Long Z_ASCII	= Z_TEXT
Constant Long Z_UNKNOWN	= 2

// The deflate compression method (the only one supported in this version)
Constant Long Z_DEFLATED = 8

// Unzip constants
Constant Long UNZ_OK							= 0
Constant Long UNZ_END_OF_LIST_OF_FILE	= -100
Constant Long UNZ_ERRNO						= Z_ERRNO
Constant Long UNZ_EOF						= 0
Constant Long UNZ_PARAMERROR				= -102
Constant Long UNZ_BADZIPFILE				= -103
Constant Long UNZ_INTERNALERROR			= -104
Constant Long UNZ_CRCERROR					= -105

// Zip constants
Constant Long ZIP_OK					= 0
Constant Long ZIP_ERRNO				= Z_ERRNO
Constant Long ZIP_PARAMERROR		= -102
Constant Long ZIP_INTERNALERROR	= -104

// Utility constants
Constant Long DEF_MEM_LEVEL = 8
Constant Long MAX_WBITS = 15

// Constants for CreateFile API function
Constant Long  INVALID_HANDLE_VALUE = -1
Constant ULong GENERIC_READ		= 2147483648
Constant ULong GENERIC_WRITE		= 1073741824
Constant ULong FILE_SHARE_READ	= 1
Constant ULong FILE_SHARE_WRITE	= 2
Constant ULong CREATE_NEW			= 1
Constant ULong CREATE_ALWAYS		= 2
Constant ULong OPEN_EXISTING		= 3
Constant ULong OPEN_ALWAYS			= 4
Constant ULong TRUNCATE_EXISTING = 5

end variables

forward prototypes
public function unsignedlong of_getfilecrc (string as_filename)
public function boolean of_readfile (string as_filename, ref blob ablob_data)
public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace)
public function boolean of_writefile (string as_filename, blob ablob_data)
public function string of_zlibversion ()
public function string of_directory (string as_filename, boolean ab_subdir)
public function string of_getfiledescription (string as_filename)
public function string of_getlasterror ()
public function boolean of_compress (blob ablob_source, ref blob ablob_dest)
public function boolean of_compress (string as_source, ref blob ablob_dest)
public function boolean of_uncompress (blob ablob_src, unsignedlong aul_destlen, ref blob ablob_dest)
public subroutine of_setpassword (string as_password)
public function integer of_internal_fa (string as_filename)
public function string of_getpassword ()
public function boolean of_uncompress (blob ablob_src, unsignedlong aul_destlen, ref string as_dest)
public function boolean of_checkbit (unsignedlong aul_number, unsignedinteger ai_bit)
public function string of_findmimefromdata (string as_filename, ref blob ablob_data)
public function long of_gzclose (unsignedlong aul_unzfile)
public function boolean of_compress2 (blob ablob_source, ref blob ablob_dest, long al_level)
public function boolean of_compress2 (string as_source, ref blob ablob_dest, long al_level)
public function boolean of_currentfilehaspassword (long al_unzfile)
public function boolean of_extractblob (long al_unzfile, string as_filename, ref blob ablob_data)
public function boolean of_extractfile (long al_unzfile, string as_filename, string as_nameinzip)
public function long of_gzopen (string as_filename, string as_mode)
public function unsignedlong of_gzread (unsignedlong aul_unzfile, ref blob ablob_data, unsignedlong aul_destlen)
public function long of_gzwrite (long al_gzipfile, ref blob ablob_data)
public function long of_importblob (long al_zipfile, string as_filename, blob ablob_data)
public function long of_importfile (long al_zipfile, string as_filename, string as_nameinzip)
public function long of_importfolder (long al_zipfile, string as_folder)
public function long of_importfolder (long al_zipfile, string as_folder, string as_subfolder)
public function long of_unzclose (long al_unzfile)
public function long of_unzclosecurrentfile (long al_unzfile)
public function long of_unzgetcurrentfileinfo (long al_unzfile, ref string as_fullname, ref string as_name, ref string as_path, ref unsignedlong aul_usize, ref unsignedlong aul_csize, ref datetime adt_datetime, ref boolean ab_password, ref boolean ab_readonly, ref boolean ab_hidden, ref boolean ab_system, ref boolean ab_subdir, ref boolean ab_archive)
public function string of_unzgetglobalcomment (long al_unzfile)
public function long of_unzgotofirstfile (long al_unzfile)
public function integer of_unzgotonextfile (long al_unzfile)
public function long of_unzopen (string as_filename)
public function long of_unzreadcurrentfile (long al_unzfile, ref blob ablob_data, unsignedlong aul_destlen)
public function long of_zipclose (long al_zipfile, string as_comment)
public function integer of_zipclosefileinzip (long al_zipfile)
public function long of_zipopen (string as_filename)
public function unsignedlong of_getfilecrc (string as_filename)
public function boolean of_readfile (string as_filename, ref blob ablob_data)
public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace)
public function boolean of_writefile (string as_filename, blob ablob_data)
public function string of_zlibversion ()
public function string of_directory (string as_filename, boolean ab_subdir)
public function string of_getfiledescription (string as_filename)
public function string of_getlasterror ()
public function boolean of_compress (blob ablob_source, ref blob ablob_dest)
public function boolean of_compress (string as_source, ref blob ablob_dest)
public function boolean of_uncompress (blob ablob_src, unsignedlong aul_destlen, ref blob ablob_dest)
public subroutine of_setpassword (string as_password)
public function integer of_internal_fa (string as_filename)
public function string of_getpassword ()
public function boolean of_uncompress (blob ablob_src, unsignedlong aul_destlen, ref string as_dest)
public function boolean of_checkbit (unsignedlong aul_number, unsignedinteger ai_bit)
public function string of_findmimefromdata (string as_filename, ref blob ablob_data)
public function long of_gzclose (unsignedlong aul_unzfile)
public function boolean of_compress2 (blob ablob_source, ref blob ablob_dest, long al_level)
public function boolean of_compress2 (string as_source, ref blob ablob_dest, long al_level)
public function boolean of_currentfilehaspassword (long al_unzfile)
public function boolean of_extractblob (long al_unzfile, string as_filename, ref blob ablob_data)
public function boolean of_extractfile (long al_unzfile, string as_filename, string as_nameinzip)
public function long of_gzopen (string as_filename, string as_mode)
public function unsignedlong of_gzread (unsignedlong aul_unzfile, ref blob ablob_data, unsignedlong aul_destlen)
public function long of_gzwrite (long al_gzipfile, ref blob ablob_data)
public function long of_importblob (long al_zipfile, string as_filename, blob ablob_data)
public function long of_importfile (long al_zipfile, string as_filename, string as_nameinzip)
public function long of_importfolder (long al_zipfile, string as_folder)
public function long of_importfolder (long al_zipfile, string as_folder, string as_subfolder)
public function long of_unzclose (long al_unzfile)
public function long of_unzclosecurrentfile (long al_unzfile)
public function long of_unzgetcurrentfileinfo (long al_unzfile, ref string as_fullname, ref string as_name, ref string as_path, ref unsignedlong aul_usize, ref unsignedlong aul_csize, ref datetime adt_datetime, ref boolean ab_password, ref boolean ab_readonly, ref boolean ab_hidden, ref boolean ab_system, ref boolean ab_subdir, ref boolean ab_archive)
public function string of_unzgetglobalcomment (long al_unzfile)
public function long of_unzgotofirstfile (long al_unzfile)
public function integer of_unzgotonextfile (long al_unzfile)
public function long of_unzopen (string as_filename)
public function long of_unzreadcurrentfile (long al_unzfile, ref blob ablob_data, unsignedlong aul_destlen)
public function long of_zipclose (long al_zipfile, string as_comment)
public function integer of_zipclosefileinzip (long al_zipfile)
public function long of_zipopen (string as_filename)
public function long of_zipopennewfileinzip (long al_zipfile, string as_filename, string as_nameinzip)
public function long of_zipwriteinfileinzip (long al_zipfile, ref blob ablob_data)
end prototypes

public function unsignedlong of_getfilecrc (string as_filename);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_GetFileCRC
//
// PURPOSE:    This function calculates a files CRC32 value.
//
// ARGUMENTS:  as_filename	- The name of the file
//
// RETURN:     CRC32 value, Zero=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

ULong lul_crc, lul_length
Blob lblob_buffer

lul_length = FileLength(as_filename)

If of_ReadFile(as_filename, lblob_buffer) Then
	lul_crc = crc32(lul_crc, lblob_buffer, lul_length);
Else
	lul_crc = 0
End If

Return lul_crc

end function

public function boolean of_readfile (string as_filename, ref blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ReadFile
//
// PURPOSE:    This function reads a file into a blob variable.
//
// ARGUMENTS:  as_filename	- The name of the file
//					ablob_data	- Contents of the file ( by ref )
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 08/21/2009	RolandS		Changed file handle to be Long instead of ULong
// -----------------------------------------------------------------------------

Long ll_file
ULong lul_bytes, lul_length
Blob lblob_filedata
Boolean lb_result

// open file for read
ll_file = CreateFile(as_filename, GENERIC_READ, &
					FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
If ll_file = INVALID_HANDLE_VALUE Then
	Return False
End If

// get file length
lul_length = FileLength(as_filename)

// allocate buffer space
lblob_filedata = Blob(Space(lul_length), EncodingAnsi!)

// read the entire file contents in one shot
lb_result = ReadFile(ll_file, lblob_filedata, &
							lul_length, lul_bytes, 0)

If lb_result Then
	// copy buffer to by reference blob
	ablob_data = BlobMid(lblob_filedata, 1, lul_length)
Else
	SetNull(ablob_data)
End If

// close the file
CloseHandle(ll_file)

Return lb_result

end function

public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ReplaceAll
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
// 03/26/2008	RolandS		Initial creation
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

public function boolean of_writefile (string as_filename, blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_WriteFile
//
// PURPOSE:    This function writes a blob variable to a file.
//
// ARGUMENTS:  as_filename	- The name of the file
//					ablob_data	- Contents of the file
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 08/21/2009	RolandS		Changed file handle to be Long instead of ULong
// -----------------------------------------------------------------------------

Long ll_file
ULong lul_written, lul_length
Boolean lb_result

// open file for write
ll_file = CreateFile(as_filename, GENERIC_WRITE, &
					FILE_SHARE_WRITE, 0, CREATE_ALWAYS, 0, 0)
If ll_file = INVALID_HANDLE_VALUE Then
	Return False
End If

// get file length
lul_length = Len(ablob_data)

// write file to disk
lb_result = WriteFile(ll_file, ablob_data, &
							lul_length, lul_written, 0)

// close the file
CloseHandle(ll_file)

Return lb_result

end function

public function string of_zlibversion ();// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ZlibVersion
//
// PURPOSE:    This function returns the version of Zlib.
//
// RETURN:     Version string
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_version

ls_version = Space(20)

ls_version = zlibVersion()

Return ls_version

end function

public function string of_directory (string as_filename, boolean ab_subdir);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Directory
//
// PURPOSE:    This function returns directory information about the
//					current file in the zip archive file.
//
// ARGUMENTS:  as_filename	- The name of the file
//					ab_subdir	- Whether subdirectory entries should be returned
//
// RETURN:     Datawindow import string
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 07/25/2009	RolandS		Formatting of DataWindow string is now here
// -----------------------------------------------------------------------------

Integer li_rc
DateTime ldt_datetime
Long ll_unzFile
ULong lul_usize, lul_csize
String ls_dwdata, ls_rowdata, ls_fullname, ls_name, ls_path, ls_attrib
Boolean lb_password, lb_readonly, lb_hidden
Boolean lb_system, lb_subdir, lb_archive

// open the zip archive
ll_unzFile = this.of_unzOpen(as_filename)

If ll_unzFile > 0 Then
	// walk thru each file in the zip archive
	li_rc = this.of_unzGoToFirstFile(ll_unzFile)
	DO WHILE li_rc = 0
		// get information about the file
		of_unzGetCurrentFileInfo(ll_unzFile, ls_fullname, ls_name, ls_path, lul_usize, &
					lul_csize, ldt_datetime, lb_password, lb_readonly, &
					lb_hidden, lb_system, lb_subdir, lb_archive)
		// skip subdirectories if not needed
		If lb_subdir And Not ab_subdir Then
			li_rc = this.of_unzGoToNextFile(ll_unzFile)
			Continue
		End If
		// build datawindow import string
		ls_rowdata  = Trim(ls_name) + "~t"
		If lb_subdir Then
			ls_rowdata += "Folder~t"
		Else
			ls_rowdata += of_GetFileDescription(ls_name) + "~t"
		End If
		ls_rowdata += String(ldt_datetime) + "~t"
		ls_rowdata += String(lul_usize) + "~t"
		If lul_usize = 0 Then
			ls_rowdata += "0~t"
		Else
			ls_rowdata += String(lul_csize / lul_usize) + "~t"
		End If
		ls_rowdata += String(lul_csize) + "~t"
		ls_attrib = ""
		If lb_readonly Then
			ls_attrib += "R"
		End If
		If lb_hidden Then
			ls_attrib += "H"
		End If
		If lb_system Then
			ls_attrib += "S"
		End If
		If lb_archive Then
			ls_attrib += "A"
		End If
		ls_rowdata += ls_attrib + "~t"
		If lb_password Then
			ls_rowdata += "Y~t"
		Else
			ls_rowdata += "N~t"
		End If
		ls_rowdata += ls_path + "~t"
		ls_rowdata += ls_fullname + "~t"
		If Len(ls_dwdata) = 0 Then
			ls_dwdata = ls_rowdata
		Else
			ls_dwdata = ls_dwdata + "~r~n" + ls_rowdata
		End If
		li_rc = this.of_unzGoToNextFile(ll_unzFile)
	LOOP
	// close the zip archive
	this.of_unzClose(ll_unzFile)
End If

Return ls_dwdata

end function

public function string of_getfiledescription (string as_filename);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_GetFileDescription
//
// PURPOSE:    This function returns the file description used by Win Explorer.
//
// ARGUMENTS:  as_filename	- The name of the file
//
// RETURN:     Description
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Constant Long SHGFI_ICON = 256 
Constant Long SHGFI_DISPLAYNAME = 512
Constant Long SHGFI_TYPENAME = 1024
Constant Long SHGFI_USEFILEATTRIBUTES = 16

SHFILEINFO lstr_SFI
String ls_typename
Long ll_rc

ll_rc = SHGetFileInfo(as_filename, 0, lstr_SFI, &
					352, SHGFI_TYPENAME + SHGFI_USEFILEATTRIBUTES)
If ll_rc = 1 Then
	ls_typename    = String(lstr_SFI.szTypeName)
End If

Return ls_typename

end function

public function string of_getlasterror ();// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_GetLastError
//
// PURPOSE:    This function returns the text of the last Win API error.
//
// RETURN:     The error message
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Constant ULong FORMAT_MESSAGE_FROM_SYSTEM = 4096
Constant ULong LANG_NEUTRAL = 0
String ls_buffer
ULong lul_rtn, lul_error

ls_buffer = Space(200)

lul_error = GetLastError()

lul_rtn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, &
				lul_error, LANG_NEUTRAL, ls_buffer, 200, 0)

Return ls_buffer

end function

public function boolean of_compress (blob ablob_source, ref blob ablob_dest);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Compress
//
// PURPOSE:    This function will compress the passed blob.
//
// ARGUMENTS:  ablob_source	-	Input blob variable
//					ablob_dest		-	Output blob variable
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
ULong lul_srclen, lul_destlen
Blob lblob_dest

// allocate buffer space
lul_srclen  = Len(ablob_source)
lul_destlen = (lul_srclen * 1.01) + 12  // zlib needs a little extra room
lblob_dest = Blob(Space(lul_destlen))

// compress the blob
ll_rc = compress(lblob_dest, lul_destlen, ablob_source, lul_srclen)
CHOOSE CASE ll_rc
	CASE Z_OK
		ablob_dest = BlobMid(lblob_dest, 1, lul_destlen)
	CASE Z_MEM_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_Compress", &
				"Not enough memory!", StopSign!)
		Return False
	CASE Z_BUF_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_Compress", &
				"Not enough room in the output buffer!", StopSign!)
		Return False
	CASE ELSE
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_Compress", &
				"Undefined error: " + String(ll_rc), StopSign!)
		Return False
END CHOOSE

Return True

end function

public function boolean of_compress (string as_source, ref blob ablob_dest);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Compress
//
// PURPOSE:    This function will compress the passed string.
//
// ARGUMENTS:  as_source	-	Input string variable
//					ablob_dest	-	Output blob variable
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Blob lblob_src

// convert string to blob
lblob_src  = Blob(as_source)

Return of_Compress(lblob_src, ablob_dest)

end function

public function boolean of_uncompress (blob ablob_src, unsignedlong aul_destlen, ref blob ablob_dest);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Uncompress
//
// PURPOSE:    This function will uncompress the passed blob to a blob.
//
// ARGUMENTS:  ablob_source	-	Input blob variable
//					aul_destlen		-	Original length of the input
//					ablob_dest		-	Output blob variable
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
ULong lul_destlen, lul_srclen
Blob lblob_dest

// allocate buffer space
lblob_dest = Blob(Space(aul_destlen))
lul_destlen = Len(lblob_dest)
lul_srclen  = Len(ablob_src)

// uncompress the blob
ll_rc = uncompress(lblob_dest, lul_destlen, ablob_src, lul_srclen)
CHOOSE CASE ll_rc
	CASE Z_OK
		ablob_dest = BlobMid(lblob_dest, 1, aul_destlen)
	CASE Z_MEM_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_UnCompress", &
				"Not enough memory!", StopSign!)
		Return False
	CASE Z_BUF_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_UnCompress", &
				"Not enough room in the output buffer!", StopSign!)
		Return False
	CASE Z_DATA_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_UnCompress", &
				"The input data was corrupted!", StopSign!)
		Return False
	CASE ELSE
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_UnCompress", &
				"Undefined error: " + String(ll_rc), StopSign!)
		Return False
END CHOOSE

Return True

end function

public subroutine of_setpassword (string as_password);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_SetPassword
//
// PURPOSE:    This function sets the password.
//
// ARGUMENTS:  as_password	- Password to save
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

is_password = as_password

end subroutine

public function integer of_internal_fa (string as_filename);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Internal_FA
//
// PURPOSE:    This function will determine if the file is likely to be a
//					text file or a binary file.  This is used for the internal_fa
//					property when adding a file to a .zip archive.  It is only
//					for documentation, it doesn't need to be accurate.
//
// ARGUMENTS:  as_filename	- The name of the file
//
// RETURN:     2=unknown, 1=text, 0=binary
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 08/21/2009	RolandS		Changed file handle to be Long instead of ULong
// -----------------------------------------------------------------------------

Long ll_file
ULong lul_bytes
String ls_mime
Blob lblob_filedata
Boolean lb_result

// open file for read
ll_file = CreateFile(as_filename, GENERIC_READ, &
					FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
If ll_file = INVALID_HANDLE_VALUE Then
	Return 2
End If

// allocate buffer space
lblob_filedata = Blob(Space(256), EncodingAnsi!)

// read first 256 bytes of the file
lb_result = ReadFile(ll_file, lblob_filedata, &
							256, lul_bytes, 0)

// close the file
CloseHandle(ll_file)

ls_mime = of_FindMimeFromData(as_filename, lblob_filedata)

If Lower(Left(ls_mime, 4)) = "text" Then
	Return 1
End If

Return 0

end function

public function string of_getpassword ();// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_GetPassword
//
// PURPOSE:    This function returns the current password.
//
// RETURN:     Password string
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Return is_password

end function

public function boolean of_uncompress (blob ablob_src, unsignedlong aul_destlen, ref string as_dest);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Uncompress
//
// PURPOSE:    This function will uncompress the passed blob to a blob.
//
// ARGUMENTS:  ablob_source	-	Input blob variable
//					aul_destlen		-	Original length of the input
//					ablob_dest		-	Output blob variable
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Blob lblob_dest

// allocate buffer space
lblob_dest = Blob(Space(aul_destlen))

If of_Uncompress(ablob_src, Len(lblob_dest), lblob_dest) Then
	as_dest = String(lblob_dest)
	Return True
Else
	as_dest = ""
End If

Return False

end function

public function boolean of_checkbit (unsignedlong aul_number, unsignedinteger ai_bit);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Checkbit
//
// PURPOSE:    This function determines if a certain bit is on or off within
//					the number.
//
// ARGUMENTS:  aul_number	- Number to check bits
//             ai_bit		- Bit number ( starting at 1 )
//
// RETURN:		True = On, False = Off
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 07/23/2009	RolandS		Initial Coding
// -----------------------------------------------------------------------------

If Int(Mod(aul_number / (2 ^(ai_bit - 1)), 2)) > 0 Then
	Return True
End If

Return False

end function

public function string of_findmimefromdata (string as_filename, ref blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_FindMimeFromData
//
// PURPOSE:    This function is determines the file MIME type.
//
// ARGUMENTS:  as_filename	- The name of the file
//					ablob_data	- Contents of the file
//
// RETURN:     MIME Type, Null=Failure
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Constant ulong NOERROR			= 0
Constant ulong FMFD_DEFAULT	= 0
Constant Long NULL = 0
String ls_mimetype
Long ll_ptr
ULong lul_rtn

lul_rtn = FindMimeFromData(NULL, as_filename, ablob_data, &
				Len(ablob_data), NULL, FMFD_DEFAULT, ll_ptr, 0)

If lul_rtn = NOERROR Then
	ls_mimetype = String(ll_ptr, "address")
Else
	SetNull(ls_mimetype)
End If

Return ls_mimetype

end function

public function long of_gzclose (unsignedlong aul_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_gzClose
//
// PURPOSE:    This function will close the gzip archive file.
//
// ARGUMENTS:  aul_unzfile	- Handle to open gzip archive file
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/08/2010	SebK			Initial creation
// -----------------------------------------------------------------------------

Long ll_rc

ll_rc = gzclose(aul_unzfile)

Return ll_rc

end function

public function boolean of_compress2 (blob ablob_source, ref blob ablob_dest, long al_level);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Compress2
//
// PURPOSE:    This function will compress the passed blob using the
//					requested compression level.
//
// ARGUMENTS:  ablob_source	-	Input blob variable
//					ablob_dest		-	Output blob variable
//					al_level			-	Compression level
//												Z_NO_COMPRESSION
//												Z_BEST_SPEED
//												Z_BEST_COMPRESSION
//												Z_DEFAULT_COMPRESSION
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 12/29/2011	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
ULong lul_srclen, lul_destlen
Blob lblob_dest

// allocate buffer space
lul_srclen  = Len(ablob_source)
lul_destlen = (lul_srclen * 1.01) + 12  // zlib needs a little extra room
lblob_dest = Blob(Space(lul_destlen))

// compress the blob
ll_rc = compress2(lblob_dest, lul_destlen, ablob_source, lul_srclen, al_level)
CHOOSE CASE ll_rc
	CASE Z_OK
		ablob_dest = BlobMid(lblob_dest, 1, lul_destlen)
	CASE Z_MEM_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_Compress2", &
				"Not enough memory!", StopSign!)
		Return False
	CASE Z_BUF_ERROR
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_Compress2", &
				"Not enough room in the output buffer!", StopSign!)
		Return False
	CASE ELSE
		SetNull(ablob_dest)
		MessageBox("ZLib Error in of_Compress2", &
				"Undefined error: " + String(ll_rc), StopSign!)
		Return False
END CHOOSE

Return True

end function

public function boolean of_compress2 (string as_source, ref blob ablob_dest, long al_level);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_Compress2
//
// PURPOSE:    This function will compress the passed string using the
//					requested compression level.
//
// ARGUMENTS:  ablob_source	-	Input blob variable
//					ablob_dest		-	Output blob variable
//					al_level			-	Compression level
//												Z_NO_COMPRESSION
//												Z_BEST_SPEED
//												Z_BEST_COMPRESSION
//												Z_DEFAULT_COMPRESSION
//
// RETURN:     True=Success, False=Failure
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 12/29/2011	RolandS		Initial creation
// -----------------------------------------------------------------------------

Blob lblob_src

// convert string to blob
lblob_src  = Blob(as_source)

Return of_Compress2(lblob_src, ablob_dest, al_level)

end function

public function boolean of_currentfilehaspassword (long al_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_CurrentFileHasPassword
//
// PURPOSE:    This function determines if the current file has a password.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//
// RETURN:     True=password, False=no password
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 12/29/2011	RolandS		Changed how password is detected
// -----------------------------------------------------------------------------

Long ll_rc
ULong lul_fname, lul_extraptr, lul_extra, lul_cmts
String ls_fullname, ls_cmts
unz_fileinfo lstr_info

// initialize buffers
lul_fname = 256
ls_fullname  = Space(lul_fname)
lul_cmts  = 500
ls_cmts   = Space(lul_cmts)

// get file info
ll_rc = unzGetCurrentFileInfo(al_unzfile, lstr_info, ls_fullname, &
					lul_fname, lul_extraptr, lul_extra, ls_cmts, lul_cmts)

// check for password
Return of_CheckBit(lstr_info.flag, 1)

end function

public function boolean of_extractblob (long al_unzfile, string as_filename, ref blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ExtractBlob
//
// PURPOSE:    This function will extract a file from the zip archive and
//             save it in a blob.  The zip archive must have already been
//					opened using the using the of_unzOpen function.
//
// ARGUMENTS:  al_unzfile	-	Handle of currently open zip archive
//					as_filename	-	Name of file within the archive
//					ablob_dest	-	Output blob containing the file
//
// RETURN:     True=Success, False=Error
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
String ls_cmts, ls_filename
ULong lul_fname, lul_extraptr, lul_extra, lul_cmts, lul_datalen
unz_fileinfo lstr_info

// locate the file within the archive
If Len(as_filename) > 0 Then
	ll_rc = unzLocateFile(al_unzfile, as_filename, 2)
	If ll_rc < 0 Then
		MessageBox("ZLib Error in of_ExtractBlob", &
				"Unable to locate file '" + as_filename + "' within archive!", StopSign!)
		Return False
	End If
End If

// if file has password and none given, display error
If of_CurrentFileHasPassword(al_unzfile) Then
	If is_password = "" Then
		MessageBox("ZLib Error in of_ExtractBlob", &
				"The file '" + as_filename + "' requires a password!", StopSign!)
		Return False
	End If
End If

// initialize buffers
lul_fname = 256
ls_filename = Space(lul_fname)
lul_cmts  = 500
ls_cmts   = Space(lul_cmts)

// get file info
ll_rc = unzGetCurrentFileInfo(al_unzfile, lstr_info, ls_filename, &
					lul_fname, lul_extraptr, lul_extra, ls_cmts, lul_cmts)

// open the current file
If is_password = "" Then
	ll_rc = unzOpenCurrentFile(al_unzfile)
Else
	ll_rc = unzOpenCurrentFilePassword(al_unzfile, is_password)
End If

// extract the current file
lul_datalen = lstr_info.uncompressed_size
ll_rc = this.of_unzReadCurrentFile(al_unzfile, ablob_data, lul_datalen)

// close the current file
ll_rc = this.of_unzCloseCurrentFile(al_unzfile)

Return True

end function

public function boolean of_extractfile (long al_unzfile, string as_filename, string as_nameinzip);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ExtractFile
//
// PURPOSE:    This function will extract a file from the zip archive and
//             save it to disk.  The zip archive must have already been
//					opened using the using the of_unzOpen function.
//
// ARGUMENTS:  al_unzfile		- Handle of currently open zip archive
//					as_filename		- Name of file being extracted from archive
//					as_nameinzip	- Name of the file within the zip file
//
// RETURN:		True=Success, False=Error
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 01/23/2012	RolandS		Added error handling to open current file
// -----------------------------------------------------------------------------

Integer li_fnum
Long ll_rc
Blob lblob_data
ULong lul_datalen

// locate the file within the archive
If Len(as_filename) > 0 Then
	ll_rc = unzLocateFile(al_unzfile, as_nameinzip, 2)
	If ll_rc < 0 Then
		MessageBox("ZLib Error in of_ExtractFile", &
				"Unable to locate file '" + as_nameinzip + "' within archive!", StopSign!)
		Return False
	End If
End If

// if file has password and none given, display error
If of_CurrentFileHasPassword(al_unzfile) Then
	If is_password = "" Then
		MessageBox("ZLib Error in of_ExtractFile", &
				"The file '" + as_filename + "' requires a password!", StopSign!)
		Return False
	End If
End If

// open the current file
If is_password = "" Then
	ll_rc = unzOpenCurrentFile(al_unzfile)
	If ll_rc < 0 Then
		MessageBox("ZLib Error in of_ExtractFile", &
				"Error opening file '" + as_nameinzip + "' within archive!", StopSign!)
		Return False
	End If
Else
	ll_rc = unzOpenCurrentFilePassword(al_unzfile, is_password)
	If ll_rc < 0 Then
		MessageBox("ZLib Error in of_ExtractFile", &
				"Error opening password protected file '" + as_nameinzip + "' within archive!", StopSign!)
		Return False
	End If
End If

// allocate buffer space
lul_datalen = 8192
lblob_data = Blob(Space(lul_datalen))

// open file
li_fnum = FileOpen(as_filename, StreamMode!, Write!, LockReadWrite!, Replace!)
If li_fnum > 0 Then
	// read file 8k bytes at a time until done
	ll_rc = unzReadCurrentFile(al_unzfile, lblob_data, lul_datalen)
	If ll_rc < 0 Then
		FileClose(li_fnum)
		MessageBox("ZLib Error in of_ExtractFile", &
				"Error extracting file '" + as_nameinzip + "' within archive!", StopSign!)
		Return False
	End If
	DO WHILE ll_rc > 0
		// write data to file
		ll_rc = FileWrite(li_fnum, BlobMid(lblob_data, 1, ll_rc))
		// read next 8k bytes
		ll_rc = unzReadCurrentFile(al_unzfile, lblob_data, lul_datalen)
	LOOP
	// close file
	FileClose(li_fnum)
Else
	MessageBox("ZLib Error in of_ExtractFile", &
			"Unable to open the output file!", StopSign!)
	Return False
End If

// close the current file
ll_rc = this.of_unzCloseCurrentFile(al_unzfile)

Return True

end function

public function long of_gzopen (string as_filename, string as_mode);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_gzOpen
//
// PURPOSE:    This function will open the gzipped file.
//
// ARGUMENTS:  as_filename	- The name of the file
//             as_mode - the mode for opening the file : rb (=read binary) / wb (=write binary)
//
// RETURN:     Handle to open gzip archive file
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/08/2010	SebK			Initial creation
// -----------------------------------------------------------------------------

Long ll_unzFile

ll_unzFile = gzopen(as_filename, as_mode)
If ll_unzFile < 0 Then
	MessageBox("ZLib Error in of_gzOpen", &
			"The gzip file is invalid or can't be found!", StopSign!)
End If

Return ll_unzFile

end function

public function unsignedlong of_gzread (unsignedlong aul_unzfile, ref blob ablob_data, unsignedlong aul_destlen);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_gzRead
//
// PURPOSE:    This function uncompresses the current file in the gzip archive
//					file and places it in a blob variable.
//
// ARGUMENTS:  aul_unzfile		-	Handle to open gzip archive file.
//					ablob_data		-	Blob to contain file (output).
//					aul_destlen		-	The buffer size for uncompressed data
//
// RETURN:     the size of read data
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/08/2010	SebK			Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
Blob lblob_data
ULong lul_next, lul_total = 0
UInt lui_datalen

// allocate buffer space
lui_datalen = 8192
lblob_data = Blob(Space(lui_datalen))
ablob_data = Blob(Space(aul_destlen))
lul_next = 1

// read file 8k bytes at a time until done
ll_rc = gzread(aul_unzfile, lblob_data, lui_datalen)
DO WHILE ll_rc > 0
	lul_total += ll_rc
	// concatenate to output blob
	lul_next = BlobEdit(ablob_data, lul_next, BlobMid(lblob_data, 1, ll_rc))
	// read next 8k bytes
	ll_rc = gzRead(aul_unzfile, lblob_data, lui_datalen)
LOOP
lul_total += ll_rc

Return lul_total


end function

public function long of_gzwrite (long al_gzipfile, ref blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_gzWrite
//
// PURPOSE:    This function will write into a gzip file
//
// ARGUMENTS:  al_zipfile	-	Handle to open gzip archive file.
//             ablob_data  -  the data to write
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/08/2010	SebK			Initial creation
// -----------------------------------------------------------------------------

Integer ll_rc
Long ll_sizebuf

ll_sizebuf = Len(ablob_data)

ll_rc = gzwrite(al_gzipfile, ablob_data, ll_sizebuf)

Return ll_rc

end function

public function long of_importblob (long al_zipfile, string as_filename, blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ImportBlob
//
// PURPOSE:    This function will add a new blob to the zip archive file.
//
// ARGUMENTS:  al_unzfile	- Handle of currently open zip archive
//					as_filename	- Name of file being added to archive
//					ablob_data	- Blob data to be imported
//
// RETURN:		0=Success
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 07/26/2009	RolandS		Moved zipOpen code to be here so could set datetime
// -----------------------------------------------------------------------------

zip_fileinfo lstr_fileinfo
String ls_cmts
ULong lul_null, lul_crc
Long ll_rc

SetNull(ls_cmts)
SetNull(lul_null)

lstr_fileinfo.tmz_date.tm_mon  = Month(Today()) - 1
lstr_fileinfo.tmz_date.tm_mday = Day(Today())
lstr_fileinfo.tmz_date.tm_year = Year(Today())
lstr_fileinfo.tmz_date.tm_hour = Hour(Now())
lstr_fileinfo.tmz_date.tm_min  = Minute(Now())
lstr_fileinfo.tmz_date.tm_sec  = Second(Now())

If is_password = "" Then
	ll_rc = zipOpenNewFileInZip(al_zipfile, as_filename, &
						lstr_fileinfo, lul_null, 0, lul_null, 0, &
						ls_cmts, Z_DEFLATED, Z_DEFAULT_COMPRESSION)
Else
	lul_crc = of_GetFileCRC(as_filename)
	ll_rc = zipOpenNewFileInZip3(al_zipfile, as_filename, &
						lstr_fileinfo, lul_null, 0, lul_null, 0, &
						ls_cmts, Z_DEFLATED, Z_DEFAULT_COMPRESSION, &
						0, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY, &
						is_password, lul_crc)
End If

If ll_rc = 0 Then
	ll_rc = this.of_zipWriteInFileInZip(al_zipFile, ablob_data)
	ll_rc = this.of_zipCloseFileInZip(al_zipFile)
End If

Return ll_rc

end function

public function long of_importfile (long al_zipfile, string as_filename, string as_nameinzip);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ImportFile
//
// PURPOSE:    This function will add a new file to the zip archive file.
//
// ARGUMENTS:  al_unzfile		- Handle of currently open zip archive
//					as_filename		- Name of file being added to archive
//					as_nameinzip	- Name of the file within the zip file
//
// RETURN:		0=Success
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Integer li_fnum
Blob lblob_file
Long ll_rc

li_fnum = FileOpen(as_filename, StreamMode!, Read!, Shared!)
If li_fnum > 0 Then
	ll_rc = this.of_zipOpenNewFileInZip(al_zipFile, as_filename, as_nameinzip)
	If ll_rc = 0 Then
		DO WHILE FileRead(li_fnum, lblob_file) > 0
			ll_rc = this.of_zipWriteInFileInZip(al_zipFile, lblob_file)
		LOOP
		ll_rc = this.of_zipCloseFileInZip(al_zipFile)
	End If
	FileClose(li_fnum)
Else
	ll_rc = -1
End If

Return ll_rc

end function

public function long of_importfolder (long al_zipfile, string as_folder);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ImportFolder
//
// PURPOSE:    This function will add all the files within a folder
//					to the zip archive file.
//
// ARGUMENTS:  al_unzfile		- Handle of currently open zip archive
//					as_folder		- Name of folder being added to archive
//
// RETURN:		0=Success
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 07/23/2009	RolandS		Initial creation
// -----------------------------------------------------------------------------

Return of_ImportFolder(al_zipfile, as_folder, "")

end function

public function long of_importfolder (long al_zipfile, string as_folder, string as_subfolder);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ImportFolder
//
// PURPOSE:    This function will add all the files within a folder
//					to the zip archive file.
//
// ARGUMENTS:  al_unzfile		- Handle of currently open zip archive
//					as_folder		- Name of folder being added to archive
//					as_subfolder	- Subfolder name when called recursively
//
// RETURN:		0=Success
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 07/23/2009	RolandS		Initial creation
// -----------------------------------------------------------------------------

win32_find_data lstr_fd
Long ll_Handle
Boolean lb_found, lb_subdir
String ls_filespec, ls_nameinzip, ls_filename

// append filename pattern
If as_subfolder = "" Then
	ls_filespec = as_folder + "\*.*"
Else
	ls_filespec = as_folder + "\" + as_subfolder + "\*.*"
End If

// find first file
ll_Handle = FindFirstFile(ls_filespec, lstr_fd)
If ll_Handle < 1 Then Return -1

// loop through each file
Do
	ls_nameinzip = String(lstr_fd.cFilename)
	If ls_nameinzip = "." Or ls_nameinzip = ".." Then
	Else
		If as_subfolder <> "" Then
			ls_nameinzip = as_subfolder + "\" + ls_nameinzip
		End If
		lb_subdir = of_checkbit(lstr_fd.dwFileAttributes, 5)
		If lb_subdir Then
			// recursively call self to process subfolder
			of_ImportFolder(al_zipfile, as_folder, ls_nameinzip)
		Else
			// add file to the zipfile
			ls_filename = as_folder + "\" + ls_nameinzip
			of_ImportFile(al_zipFile, ls_filename, ls_nameinzip)
		End If
	End If
	// find next file
	lb_Found = FindNextFile(ll_Handle, lstr_fd)
Loop Until Not lb_Found

// close find handle
FindClose(ll_Handle)

Return 0

end function

public function long of_unzclose (long al_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzClose
//
// PURPOSE:    This function will close the zip archive file.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc

ll_rc = unzClose(al_unzfile)

Return ll_rc

end function

public function long of_unzclosecurrentfile (long al_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzCloseCurrentFile
//
// PURPOSE:    This function closes current file in the zip archive file.
//
// ARGUMENTS:  al_zipfile	-	Handle to open zip archive file.
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc

ll_rc = unzCloseCurrentFile(al_unzfile)

Return ll_rc

end function

public function long of_unzgetcurrentfileinfo (long al_unzfile, ref string as_fullname, ref string as_name, ref string as_path, ref unsignedlong aul_usize, ref unsignedlong aul_csize, ref datetime adt_datetime, ref boolean ab_password, ref boolean ab_readonly, ref boolean ab_hidden, ref boolean ab_system, ref boolean ab_subdir, ref boolean ab_archive);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzGetCurrentFileInfo
//
// PURPOSE:    This function returns directory information about the
//					current file in the zip archive file.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 07/25/2009	RolandS		Changed to return info as by ref arguments
// 12/29/2011	RolandS		Changed how password is detected
// -----------------------------------------------------------------------------

unz_fileinfo lstr_info
Long ll_rc, ll_pos
String ls_cmts, ls_temp
ULong lul_fname, lul_extraptr, lul_extra, lul_cmts
Date ld_date
Time lt_time

// initialize buffers
lul_fname = 256
as_fullname  = Space(lul_fname)
lul_cmts  = 500
ls_cmts   = Space(lul_cmts)

// get file info
ll_rc = unzGetCurrentFileInfo(al_unzfile, lstr_info, as_fullname, &
					lul_fname, lul_extraptr, lul_extra, ls_cmts, lul_cmts)
If ll_rc <> UNZ_OK Then
	Return ll_rc
End If

// extract filename/filepath from full path
ls_temp = "\" + this.of_replaceall(as_fullname, "/", "\")
If Right(ls_temp, 1) = "\" Then
	ls_temp = Left(ls_temp, Len(ls_temp) - 1)
End If
ll_pos = LastPos(ls_temp, "\")
as_name = Mid(ls_temp, ll_pos + 1)
as_path = Mid(Left(ls_temp, ll_pos), 2)

// file size
aul_usize = lstr_info.uncompressed_size
aul_csize = lstr_info.compressed_size

// build datetime
ld_date = Date(lstr_info.tmu_date.tm_year, &
					lstr_info.tmu_date.tm_mon + 1, &
					lstr_info.tmu_date.tm_mday)
lt_time = Time(lstr_info.tmu_date.tm_hour, &
					lstr_info.tmu_date.tm_min, &
					lstr_info.tmu_date.tm_sec)
adt_datetime = DateTime(ld_date, lt_time)

// has a password?
ab_password = of_CheckBit(lstr_info.flag, 1)

// file attributes
ab_readonly = of_CheckBit(lstr_info.external_fa, 1)
ab_hidden   = of_CheckBit(lstr_info.external_fa, 2)
ab_system   = of_CheckBit(lstr_info.external_fa, 3)
ab_subdir   = of_CheckBit(lstr_info.external_fa, 5)
ab_archive  = of_CheckBit(lstr_info.external_fa, 6)

Return ll_rc

end function

public function string of_unzgetglobalcomment (long al_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzGetGlobalComment
//
// PURPOSE:    This function returns the global comment for the file.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//
// RETURN:     The comment string
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

String ls_comment
Ulong lul_sizebuf
Long ll_rc

lul_sizebuf = 500
ls_comment = Space(lul_sizebuf)

ll_rc = unzGetGlobalComment(al_unzfile, ls_comment, lul_sizebuf)

Return ls_comment

end function

public function long of_unzgotofirstfile (long al_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzGotoFirstFile
//
// PURPOSE:    This function makes the first file in the zip archive
//					the current file.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc

ll_rc = unzGoToFirstFile(al_unzfile)
CHOOSE CASE ll_rc
	CASE UNZ_PARAMERROR
		MessageBox("ZLib Error in of_unzGotoFirstFile", &
				"Parameter Error!", StopSign!)
	CASE UNZ_BADZIPFILE
		MessageBox("ZLib Error in of_unzGotoFirstFile", &
				"Bad zip file!", StopSign!)
	CASE UNZ_INTERNALERROR
		MessageBox("ZLib Error in of_unzGotoFirstFile", &
				"Internal Error!", StopSign!)
	CASE UNZ_CRCERROR
		MessageBox("ZLib Error in of_unzGotoFirstFile", &
				"CRC Error!", StopSign!)
END CHOOSE

Return ll_rc

end function

public function integer of_unzgotonextfile (long al_unzfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzGotoNextFile
//
// PURPOSE:    This function makes the next file in the zip archive
//					the current file.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc

ll_rc = unzGoToNextFile(al_unzfile)
CHOOSE CASE ll_rc
	CASE UNZ_PARAMERROR
		MessageBox("ZLib Error in of_unzGotoNextFile", &
				"Parameter Error!", StopSign!)
	CASE UNZ_BADZIPFILE
		MessageBox("ZLib Error in of_unzGotoNextFile", &
				"Bad zip file!", StopSign!)
	CASE UNZ_INTERNALERROR
		MessageBox("ZLib Error in of_unzGotoNextFile", &
				"Internal Error!", StopSign!)
	CASE UNZ_CRCERROR
		MessageBox("ZLib Error in of_unzGotoNextFile", &
				"CRC Error!", StopSign!)
END CHOOSE

Return ll_rc

end function

public function long of_unzopen (string as_filename);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzOpen
//
// PURPOSE:    This function will open the zip archive file for unzipping.
//
// ARGUMENTS:  as_filename	- The name of the file
//
// RETURN:     Handle to open zip archive file
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_unzFile

ll_unzFile = unzOpen(as_filename)
If ll_unzFile < 0 Then
	MessageBox("ZLib Error in of_unzOpen", &
			"The zip file is invalid or can't be found!", StopSign!)
End If

Return ll_unzFile

end function

public function long of_unzreadcurrentfile (long al_unzfile, ref blob ablob_data, unsignedlong aul_destlen);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_unzReadCurrentFile
//
// PURPOSE:    This function uncompresses the current file in the zip archive
//					file and places it in a blob variable.
//
// ARGUMENTS:  al_unzfile		-	Handle to open zip archive file.
//					ablob_data		-	Blob to contain file (output).
//					aul_destlen		-	The uncompressed size of the file.
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
Blob lblob_data
ULong lul_next
UInt lui_datalen

// allocate buffer space
lui_datalen = 8192
lblob_data = Blob(Space(lui_datalen))
ablob_data = Blob(Space(aul_destlen))
lul_next = 1

// read file 8k bytes at a time until done
ll_rc = unzReadCurrentFile(al_unzfile, lblob_data, lui_datalen)
DO WHILE ll_rc > 0
	// concatenate to output blob
	lul_next = BlobEdit(ablob_data, lul_next, BlobMid(lblob_data, 1, ll_rc))
	// read next 8k bytes
	ll_rc = unzReadCurrentFile(al_unzfile, lblob_data, lui_datalen)
LOOP

Return ll_rc

end function

public function long of_zipclose (long al_zipfile, string as_comment);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_zipClose
//
// PURPOSE:    This function will close the zip archive file.
//
// ARGUMENTS:  al_unzfile	- Handle to open zip archive file
//					as_comment	- Global comment
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc
String ls_cmt

ll_rc = zipClose(al_zipfile, as_comment)

Return ll_rc

end function

public function integer of_zipclosefileinzip (long al_zipfile);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_zipCloseFileInZip
//
// PURPOSE:    This function closes a new file in the zip archive file.
//
// ARGUMENTS:  al_zipfile	-	Handle to open zip archive file.
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc

ll_rc = zipCloseFileInZip(al_zipfile)

Return ll_rc

end function

public function long of_zipopen (string as_filename);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_ZipOpen
//
// PURPOSE:    This function will open the zip archive file for adding files.
//
// ARGUMENTS:  as_filename	- The name of the file
//
// RETURN:     Handle to open zip archive file
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_zipFile

// open zip file
ll_zipFile = zipOpen(as_filename, 0)
If ll_zipFile < 0 Then
	MessageBox("ZLib Error in of_zipOpen", &
			"The zip file could not be opened!", StopSign!)
End If

Return ll_zipFile

end function

public function long of_zipopennewfileinzip (long al_zipfile, string as_filename, string as_nameinzip);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_zipOpenNewFileInZip
//
// PURPOSE:    This function opens a new file in the zip archive file.
//
// ARGUMENTS:  al_zipfile		- Handle to open zip archive file.
//					as_filename		- Name of file being added to archive.
//					as_nameinzip	- Name of file within the archive.
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// 07/26/2009	RolandS		Added as_nameinzip
// -----------------------------------------------------------------------------

Long ll_rc
ULong lul_Handle, lul_null, lul_crc
zip_fileinfo lstr_fileinfo
WIN32_FIND_DATA lstr_FD
FILETIME lstr_LT
SYSTEMTIME lstr_ST
String ls_cmts, ls_filename

SetNull(ls_cmts)
SetNull(lul_null)

// Get the file information
lul_Handle = FindFirstFile(as_filename, lstr_FD)
If lul_Handle <= 0 Then Return -1
FindClose(lul_Handle)

// Value of 1 indicates text file, 0 indicates binary
lstr_fileinfo.internal_fa = this.of_Internal_FA(as_filename)

// get file attributes (Read-Only, Hidden, System, Archive)
lstr_fileinfo.external_fa = lstr_FD.dwfileattributes

// convert create datetime to usable format
FileTimeToLocalFileTime(lstr_FD.ftlastwritetime, lstr_LT)
FileTimeToSystemTime(lstr_LT, lstr_ST)

// .zip standard is to use forward slash
ls_filename = this.of_replaceall(as_nameinzip, "\", "/")

// copy datetime to structure
lstr_fileinfo.tmz_date.tm_mon = lstr_ST.wMonth - 1
lstr_fileinfo.tmz_date.tm_mday = lstr_ST.wDay
lstr_fileinfo.tmz_date.tm_year = lstr_ST.wYear
lstr_fileinfo.tmz_date.tm_hour = lstr_ST.wHour
lstr_fileinfo.tmz_date.tm_min = lstr_ST.wMinute
lstr_fileinfo.tmz_date.tm_sec = lstr_ST.wSecond

If is_password = "" Then
	ll_rc = zipOpenNewFileInZip(al_zipfile, ls_filename, &
						lstr_fileinfo, lul_null, 0, lul_null, 0, &
						ls_cmts, Z_DEFLATED, Z_DEFAULT_COMPRESSION)
Else
	lul_crc = of_GetFileCRC(as_filename)
	ll_rc = zipOpenNewFileInZip3(al_zipfile, ls_filename, &
						lstr_fileinfo, lul_null, 0, lul_null, 0, &
						ls_cmts, Z_DEFLATED, Z_DEFAULT_COMPRESSION, &
						0, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY, &
						is_password, lul_crc)
End If

Return ll_rc

end function

public function long of_zipwriteinfileinzip (long al_zipfile, ref blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_zlib.of_zipWriteInFileInZip
//
// PURPOSE:    This function will write a new file to the zip archive file.
//
// ARGUMENTS:  al_zipfile	-	Handle to open zip archive file.
//
// RETURN:     Zero=Success
//
// DATE        PROG/ID 		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 03/26/2008	RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_rc, ll_sizebuf

ll_sizebuf = Len(ablob_data)

ll_rc = zipWriteInFileInZip(al_zipfile, ablob_data, ll_sizebuf)

Return ll_rc

end function

on n_zlib.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_zlib.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


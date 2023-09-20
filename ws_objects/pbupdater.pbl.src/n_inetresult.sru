$PBExportHeader$n_inetresult.sru
forward
global type n_inetresult from internetresult
end type
end forward

global type n_inetresult from internetresult
end type
global n_inetresult n_inetresult

type prototypes
Function ulong CreateFile ( &
	string lpFileName, &
	ulong dwDesiredAccess, &
	ulong dwShareMode, &
	ulong lpSecurityAttributes, &
	ulong dwCreationDisposition, &
	ulong dwFlagsAndAttributes, &
	ulong hTemplateFile &
	) Library "kernel32.dll" Alias For "CreateFileW"

Function boolean CloseHandle ( &
	ulong hObject &
	) Library "kernel32.dll"

Function boolean WriteFile ( &
	ulong hFile, &
	blob lpBuffer, &
	ulong nNumberOfBytesToWrite, &
	Ref ulong lpNumberOfBytesWritten, &
	ulong lpOverlapped &
	) Library "kernel32.dll"

end prototypes

type variables
String is_filedata
String is_filename

// constants for CreateFile API function
Constant ULong INVALID_HANDLE_VALUE		= -1
Constant ULong GENERIC_READ				= 2147483648
Constant ULong GENERIC_WRITE				= 1073741824
Constant ULong FILE_ATTRIBUTE_NORMAL	= 128
Constant ULong FILE_SHARE_READ	= 1
Constant ULong FILE_SHARE_WRITE	= 2
Constant ULong CREATE_NEW			= 1
Constant ULong CREATE_ALWAYS		= 2
Constant ULong OPEN_EXISTING		= 3
Constant ULong OPEN_ALWAYS			= 4
Constant ULong TRUNCATE_EXISTING	= 5

end variables

forward prototypes
public function integer internetdata (blob data)
public function integer of_writefile (blob ablob_data)
end prototypes

public function integer internetdata (blob data);// -----------------------------------------------------------------------------
// SCRIPT:     n_inetresult.InternetData
//
// PURPOSE:    This function receives data from the GetURL function.
//
// ARGUMENTS:  data	- The file data
//
// RETURN:     1=Success
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 03/17/2008  RolandS		Initial creation
// -----------------------------------------------------------------------------

If is_filename = "" Then
	// save the data to string
	is_filedata = String(data, EncodingAnsi!)
Else
	// write the data to disk
	of_writefile(data)
End If

Return 1

end function

public function integer of_writefile (blob ablob_data);// -----------------------------------------------------------------------------
// SCRIPT:     n_inetresult.of_WriteFile
//
// PURPOSE:    This function writes data to a file.
//
// ARGUMENTS:  ablob_data	- The file data
//
// RETURN:     1=Success, -1=Failure
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 03/17/2008  RolandS		Initial creation
// -----------------------------------------------------------------------------

ULong lul_file, lul_written
Boolean lb_rtn

// open file for write
lul_file = CreateFile(is_filename, GENERIC_WRITE, &
					FILE_SHARE_WRITE, 0, CREATE_ALWAYS, 0, 0)
If lul_file = INVALID_HANDLE_VALUE Then
	Return -1
End If

// write file to disk
lb_rtn = WriteFile(lul_file, ablob_data, &
					Len(ablob_data), lul_written, 0)

// close the file
CloseHandle(lul_file)

Return 1

end function

on n_inetresult.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_inetresult.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on


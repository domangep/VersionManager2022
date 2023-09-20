$PBExportHeader$version_manager.sra
$PBExportComments$Generated Application Object
forward
global type version_manager from application
end type
global transaction sqlca
global dynamicdescriptionarea sqlda
global dynamicstagingarea sqlsa
global error error
global message message
end forward

global variables
n_pbupdater gn_pbupdater
end variables
global type version_manager from application
string appname = "version_manager"
string themepath = "C:\Program Files (x86)\Appeon\PowerBuilder 22.0\IDE\theme"
string themename = "Do Not Use Themes"
boolean nativepdfvalid = false
boolean nativepdfincludecustomfont = false
string nativepdfappname = ""
long richtextedittype = 5
long richtexteditx64type = 5
long richtexteditversion = 3
string richtexteditkey = ""
string appicon = "vm_ico.ico"
string appruntimeversion = "22.1.0.2819"
boolean manualsession = false
boolean unsupportedapierror = false
boolean bignoreservercertificate = false
uint ignoreservercertificate = 0
end type
global version_manager version_manager

on version_manager.create
appname="version_manager"
message=create message
sqlca=create transaction
sqlda=create dynamicdescriptionarea
sqlsa=create dynamicstagingarea
error=create error
end on

on version_manager.destroy
destroy(sqlca)
destroy(sqlda)
destroy(sqlsa)
destroy(error)
destroy(message)
end on

event open;// create updater object
gn_pbupdater = Create n_pbupdater

openwithparm(w_main, commandline)
end event

event close;// run update script if exists
gn_pbupdater.of_RunScript()

// destroy the updater object
Destroy gn_pbupdater
end event


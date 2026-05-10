Attribute VB_Name = "libErrorCodes"
' libErrorCodes
Option Explicit
Option Private Module

' libInterpolation
Public Const ERR_BASE_INTERPOLATION As Long = vbObjectError + 2100
Public Const ERR_INTERPOLATION_ZERO_RANGE As Long = ERR_BASE_INTERPOLATION + 1

' libControls
Public Const ERR_BASE_CONTROLS As Long = vbObjectError + 2200
Public Const ERR_UNSUPPORTED_CONTROL As Long = ERR_BASE_CONTROLS + 1

' libString
Public Const ERR_BASE_STRING As Long = vbObjectError + 2300
Public Const ERR_INVALID_STRING_BOOL As Long = ERR_BASE_STRING + 1


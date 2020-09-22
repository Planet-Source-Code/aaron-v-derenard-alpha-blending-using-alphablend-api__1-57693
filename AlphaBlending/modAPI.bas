Attribute VB_Name = "modAPI"
'===================================='
'=  Alpha Blending Tutorial By:     ='
'=      Aaron DeRenard              ='
'=  MSU-Fall04-Ex                   ='
'=----------------------------------='
'=  Free-Use Advisory:              ='
'=      Give credit where credit is ='
'=      Due!  Turn  me in as hw     ='
'=      And I find out              ='
'=      I'll report your ass!       ='
'===================================='


'This project requires three picture boxes
'Both picture boxes should contain a picture
'Also, there should be a third picture used to refresh
'the hDC of the underlying picture!
Const AC_SRC_OVER = &H0
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)


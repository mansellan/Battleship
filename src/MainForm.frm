VERSION 5.00
Begin VB.Form GameForm 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Battleship.View.MsForms")
Option Explicit

Private Const InfoBoxMessage As String = _
    "ENEMY FLEET DETECTED" & vbNewLine & _
    "ALL SYSTEMS READY" & vbNewLine & vbNewLine & _
    "DOUBLE CLICK IN THE ENEMY GRID TO FIRE A MISSILE." & vbNewLine & vbNewLine & _
    "FIND AND DESTROY ALL ENEMY SHIPS BEFORE THEY DESTROY YOUR OWN FLEET!"

Private Const InfoBoxPlaceShips As String = _
    "FLEET DEPLOYMENT" & vbNewLine & _
    "ACTION REQUIRED: DEPLOY %SHIP%" & vbNewLine & vbNewLine & _
    " -CLICK TO PREVIEW" & vbNewLine & _
    " -RIGHT CLICK TO ROTATE" & vbNewLine & _
    " -DOUBLE CLICK TO CONFIRM" & vbNewLine & vbNewLine
    
Private Const ErrorBoxInvalidPosition As String = _
    "FLEET DEPLOYMENT" & vbNewLine & _
    "SYSTEM ERROR" & vbNewLine & vbNewLine & _
    " -SHIPS CANNOT OVERLAP." & vbNewLine & _
    " -SHIPS MUST BE ENTIRELY WITHIN THE GRID." & vbNewLine & vbNewLine & _
    "DEPLOY SHIP TO ANOTHER POSITION."
    
Private Const ErrorBoxInvalidKnownAttackPosition As String = _
    "TARGETING SYSTEM" & vbNewLine & vbNewLine & _
    "SPECIFIED GRID LOCATION IS ALREADY IN A KNOWN STATE." & vbNewLine & vbNewLine & _
    "NEW VALID COORDINATES REQUIRED."

Private previousMode As ViewMode
Private Mode As ViewMode

Public Event CreatePlayer(ByVal gridId As Byte, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
Public Event PlayerReady()
Public Event SelectionChange(ByVal gridId As Byte, ByVal position As IGridCoord, ByVal Mode As ViewMode)
Public Event RightClick(ByVal gridId As Byte, ByVal position As IGridCoord, ByVal Mode As ViewMode)
Public Event DoubleClick(ByVal gridId As Byte, ByVal position As IGridCoord, ByVal Mode As ViewMode)

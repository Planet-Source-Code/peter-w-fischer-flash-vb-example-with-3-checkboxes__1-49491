VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5805
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _cx             =   3836
      _cy             =   1296
      FlashVars       =   ""
      Movie           =   "swf"
      Src             =   "swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1095
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================
' an easy to understand little example for a concise way of bidirectional 'talking'
' between Flash and VB
'
' written Oktober, 28, 2003
' by Peter-W. Fischer
'
' any comments welcome: peter-w.fischer@epost.de
'
' you may modify this example in any way you like to but
' it would be nice to name the source: me!
'
' hope this code inspires people to do more programming with both
' flash and VB, one for the front-end, the other for the calculations!
'
'========================================================
Option Explicit
Dim CheckButs As String 'the turntable for the checkbox-actions!
Dim noaction As Boolean 'what it says!

Private Sub Form_Load()
    CheckButs = "0;0;0" 'init the CheckBoxes in VB as well
    Label1.Caption = "Press checkboxes in the flash-movie or on the VB-formular. They will react simultanously! But don't the Flasbuttons look much more 'flashy'?"
    ShockwaveFlash1.LoadMovie 0, App.Path & "\VB-Flash-Tutorial.swf" 'this flashmovie should be in the App-folder!
End Sub

'Action taken in VB and send to Flash by two Variables
'1. the CheckButs-String
'2. the CheckByVB -controlVariable for the timer in flash to react
'in this order, of course ;-)
Private Sub Check1_Click(Index As Integer)
    If noaction Then Exit Sub
    Dim arr As Variant
    'prepare the datastring for Flash
    arr = Split(CheckButs, ";")
    arr(Index) = Check1(Index).Value
    CheckButs = Join(arr, ";")
    With ShockwaveFlash1
        .SetVariable "CheckButs", CheckButs 'send it to Flash
        .SetVariable "CheckByVB", 1 ' and tell the movie to update the his Checkboxes!
    End With
End Sub

'================== Some actionscript on level0 of the flash-movie:
'
'CheckButs = "0;0;0";   => the same StringVar like in the VB-Projekt, same Name to avoid confusion!
'CheckByVB = 0; => control-var for the timer below
'setInterval(CheckNeueWerte, 100); => timer should react every tenth of a second
'function CheckNeueWerte() {
'    if (CheckByVB<>0) { => changes from VB
'        CheckByVB = 0; => set back to avoid next call
'        arr = CheckButs.split(";"); => make an array (like in VB)
'        for (i=0; i<3; i++) {
'                => move the movieClip(s) with the Name 'Chk0' to 'Chk2' to the corresponding position
'                  => of the CheckValue
'            _root["Chk"+i].gotoAndStop(arr[i] == 0 ? "NotChecked" : "Checked");
'        }
'    }
'}
' this function does the job the other way round
'it's getting a new value from the CheckBox-movieclip
'function SetCheckString(ChkNr, NewValue) {
'    arr = CheckButs.split(";");
'    arr[ChkNr] = NewValue;
'    CheckButs = arr.join(";");
'    fscommand("CheckButs", CheckButs); => and sends the whole string to VB! that's it!
'}
'======================================================================
' and here is the pipeline for messages from Flash
Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
    Dim i As Integer
    Select Case command
        Case "CheckButs" 'in this case the only command we aspect from Flash
            noaction = True 'necessary to avoid reaction of the VB-Checkboxes
            CheckButs = args 'in args is the CheckButs-String from Flash!
            For i = 0 To 2
                Check1(i).Value = Split(args, ";")(i) 'set the VB-CheckBoxen according to the values in args
            Next
            noaction = False 'and give them free again!
    End Select
End Sub


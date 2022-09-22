VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MetadataMaker V1.0"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7905
   Icon            =   "MetadataMaker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   7905
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame2 
      Caption         =   "Step 1"
      Height          =   855
      Left            =   360
      TabIndex        =   13
      Top             =   360
      Width           =   7215
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Fill in the NFT attributes in the \src\attributes.csv file."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 2"
      Height          =   4695
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   7215
      Begin VB.TextBox txtExternal_url 
         Height          =   495
         Left            =   1485
         TabIndex        =   15
         Top             =   2760
         Width           =   5175
      End
      Begin VB.CheckBox chkWhiteSpace 
         Caption         =   "Add Whitespace"
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   4080
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtNamePrefix 
         Height          =   495
         Left            =   1485
         TabIndex        =   7
         Text            =   "MyNFT #"
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox txtImageBaseURL 
         Height          =   495
         Left            =   1485
         TabIndex        =   6
         Text            =   "ipfs://QmeSjSinHpPnmXmspMjwiXyN6zS4E9zccariGR3jxcaWtq/"
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox txtDescription 
         Height          =   495
         Left            =   1485
         TabIndex        =   5
         Top             =   3480
         Width           =   5175
      End
      Begin VB.TextBox txtImageFormat 
         Height          =   495
         Left            =   1485
         TabIndex        =   4
         Text            =   "png"
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label Label8 
         Caption         =   "*"
         Height          =   255
         Left            =   6840
         TabIndex        =   18
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "*"
         Height          =   255
         Left            =   6840
         TabIndex        =   17
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "external_url"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "NamePrefix"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ImageBaseURL"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ImageFormat"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label LabelStatusBar 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   6960
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private myMetadata As JsonBag

Private Sub Command1_Click()
    Dim srcPath As String, buildPath As String, jsonPath As String
    Dim tempS As String
    Dim attributeName As Variant
    Dim totalAttribute As Integer
    Dim myAttribute As Variant
    Dim myID As Long
    Dim i As Integer
   
    On Error Resume Next
    LabelStatusBar.Caption = ""
    DoEvents
    
    If txtImageBaseURL = "" Then
        LabelStatusBar.Caption = "Error! ImageBaseURL cannot be empty."
        txtImageBaseURL.SetFocus
        Exit Sub
    End If
    
    If txtImageFormat = "" Then
        LabelStatusBar.Caption = "Error! ImageFormat cannot be empty."
        txtImageFormat.SetFocus
        Exit Sub
    End If

    If Right(App.Path, 1) = "\" Then srcPath = App.Path & "src" Else srcPath = App.Path & "\src"
    If Dir(srcPath, vbDirectory) = "" Then MkDir srcPath
    If Dir(srcPath & "\attributes.csv") = "" Then
        LabelStatusBar.Caption = "Error! src\attributes.csv file not found."
        Exit Sub
    End If
    
    If Right(App.Path, 1) = "\" Then buildPath = App.Path & "build" Else buildPath = App.Path & "\build"
    If Dir(buildPath, vbDirectory) = "" Then MkDir buildPath
    jsonPath = buildPath & "\json"
    If Dir(jsonPath, vbDirectory) = "" Then MkDir jsonPath
    Kill jsonPath & "\*.*"
    
    Open srcPath & "\attributes.csv" For Input As #1
    Line Input #1, tempS
    attributeName = Split(tempS, ",")
    totalAttribute = UBound(attributeName) + 1
    
    LabelStatusBar.Caption = " Making........."
    DoEvents
        
    Do While Not EOF(1)
        Line Input #1, tempS
        If Len(tempS) < 2 Then Exit Do
        myAttribute = Split(tempS, ",")
        myID = myAttribute(0)
        Set myMetadata = New JsonBag
        myMetadata.Whitespace = chkWhiteSpace.Value
        
        With myMetadata
            .Clear
            .IsArray = False
            If txtNamePrefix <> "" Then .Item("name") = txtNamePrefix & myID
            .Item("image") = txtImageBaseURL & myID & "." & txtImageFormat
            If txtExternal_url <> "" Then .Item("external_url") = txtExternal_url
            If txtDescription <> "" Then .Item("description") = txtDescription
            With .AddNewArray("attributes")
                For i = 1 To totalAttribute - 1
                    If myAttribute(i) <> "" Then
                        With .AddNewObject()
                            .Item("trait_type") = attributeName(i)
                             If IsNumeric(myAttribute(i)) Then .Item("value") = Val(myAttribute(i)) Else .Item("value") = myAttribute(i)
                        End With
                    End If
                Next i
            End With
        End With
        
        Open jsonPath & "\" & myID & ".json" For Output As #2
            Print #2, myMetadata.JSON
        Close #2
    Loop
    
    Close #1
    LabelStatusBar.Caption = "Done! Metadata files are in the \build\json\ directory."
    Shell "explorer " & jsonPath, 1
End Sub


Private Sub Command2_Click()
    End
End Sub


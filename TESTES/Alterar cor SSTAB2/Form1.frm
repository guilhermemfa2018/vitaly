VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   0
      ScaleHeight     =   1170
      ScaleWidth      =   11940
      TabIndex        =   64
      Top             =   0
      Width           =   11940
      Begin VB.Image Image1 
         Height          =   900
         Left            =   330
         Picture         =   "Form1.frx":0000
         Top             =   135
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Entry"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   1545
         TabIndex        =   65
         Top             =   345
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E7C5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8250
      Left            =   0
      ScaleHeight     =   8250
      ScaleWidth      =   11940
      TabIndex        =   0
      Top             =   1170
      Width           =   11940
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6E7C5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8175
         Left            =   -45
         ScaleHeight     =   8175
         ScaleWidth      =   12555
         TabIndex        =   1
         Top             =   60
         Width           =   12555
         Begin VB.CommandButton Command1 
            Caption         =   "End Program"
            Height          =   450
            Left            =   360
            TabIndex        =   66
            Top             =   7020
            Width           =   1755
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   6450
            Left            =   300
            TabIndex        =   2
            Top             =   360
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   11377
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            BackColor       =   -2147483643
            TabCaption(0)   =   "Personal Details"
            TabPicture(0)   =   "Form1.frx":06B6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label6"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label26"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label1"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label10"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label9"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label8"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label5"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label4"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label3"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label11"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label43"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Label31"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label30"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label29"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Label24"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "Label23"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Label22"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "Label42"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "cmbFacility"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "txtDrivingLicenseNo"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "txtCitizenshipIssue"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "txtPassportNo"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "txtCitizenshipNo"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "cmbBirthPlace"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "cmbSex"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "txtLastName"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "txtMiddleName"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "txtFirstName"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "cmbMaritalStatus"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "txtGrandFathersContact"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "txtFathersContact"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "txtSpouseContact"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "txtGrandFathersName"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).Control(33)=   "txtFathersName"
            Tab(0).Control(33).Enabled=   0   'False
            Tab(0).Control(34)=   "txtSpouseName"
            Tab(0).Control(34).Enabled=   0   'False
            Tab(0).Control(35)=   "txtEmployeeCode"
            Tab(0).Control(35).Enabled=   0   'False
            Tab(0).ControlCount=   36
            TabCaption(1)   =   "Address Information"
            TabPicture(1)   =   "Form1.frx":06D2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label41"
            Tab(1).Control(1)=   "Label37"
            Tab(1).Control(2)=   "Label32"
            Tab(1).Control(3)=   "Label21"
            Tab(1).Control(4)=   "Label13"
            Tab(1).Control(5)=   "Label14"
            Tab(1).Control(6)=   "Label15"
            Tab(1).Control(7)=   "Label16"
            Tab(1).Control(8)=   "Label2"
            Tab(1).Control(9)=   "Label18"
            Tab(1).Control(10)=   "Label19"
            Tab(1).Control(11)=   "Label20"
            Tab(1).Control(12)=   "txtEmailAdd"
            Tab(1).Control(13)=   "txtPager"
            Tab(1).Control(14)=   "txtCellNo"
            Tab(1).Control(15)=   "txtResContact"
            Tab(1).Control(16)=   "txtPermanentAddress"
            Tab(1).Control(17)=   "txtPZone"
            Tab(1).Control(18)=   "txtPDistrict"
            Tab(1).Control(19)=   "txtPCountry"
            Tab(1).Control(20)=   "txtTAddress"
            Tab(1).Control(21)=   "txtTZone"
            Tab(1).Control(22)=   "txtTDistrict"
            Tab(1).Control(23)=   "txtTCountry"
            Tab(1).ControlCount=   24
            TabCaption(2)   =   "Education Details"
            TabPicture(2)   =   "Form1.frx":06EE
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "Work Experience"
            TabPicture(3)   =   "Form1.frx":070A
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            TabCaption(4)   =   "Document Manager"
            TabPicture(4)   =   "Form1.frx":0726
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "lblID"
            Tab(4).ControlCount=   1
            Begin VB.TextBox txtEmployeeCode 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   660
               TabIndex        =   32
               Top             =   1325
               Width           =   2400
            End
            Begin VB.TextBox txtTCountry 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -66675
               TabIndex        =   31
               Top             =   2274
               Width           =   1620
            End
            Begin VB.TextBox txtTDistrict 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -68330
               TabIndex        =   30
               Top             =   2274
               Width           =   1545
            End
            Begin VB.TextBox txtTZone 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -70030
               TabIndex        =   29
               Top             =   2274
               Width           =   1590
            End
            Begin VB.TextBox txtTAddress 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -74310
               TabIndex        =   28
               Top             =   2274
               Width           =   4170
            End
            Begin VB.TextBox txtPCountry 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -66675
               TabIndex        =   27
               Top             =   1328
               Width           =   1620
            End
            Begin VB.TextBox txtPDistrict 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -68330
               TabIndex        =   26
               Top             =   1328
               Width           =   1545
            End
            Begin VB.TextBox txtPZone 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -70030
               TabIndex        =   25
               Top             =   1328
               Width           =   1590
            End
            Begin VB.TextBox txtPermanentAddress 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -74310
               TabIndex        =   24
               Top             =   1328
               Width           =   4170
            End
            Begin VB.TextBox txtResContact 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -74310
               TabIndex        =   23
               Top             =   3220
               Width           =   2175
            End
            Begin VB.TextBox txtCellNo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -72040
               TabIndex        =   22
               Top             =   3220
               Width           =   2175
            End
            Begin VB.TextBox txtPager 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -69770
               TabIndex        =   21
               Top             =   3220
               Width           =   2175
            End
            Begin VB.TextBox txtEmailAdd 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   -67500
               TabIndex        =   20
               Top             =   3220
               Width           =   2175
            End
            Begin VB.TextBox txtSpouseName 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   705
               TabIndex        =   19
               Top             =   4160
               Width           =   2190
            End
            Begin VB.TextBox txtFathersName 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2940
               TabIndex        =   18
               Top             =   4155
               Width           =   2190
            End
            Begin VB.TextBox txtGrandFathersName 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   5175
               TabIndex        =   17
               Top             =   4155
               Width           =   2190
            End
            Begin VB.TextBox txtSpouseContact 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   705
               TabIndex        =   16
               Top             =   5100
               Width           =   2190
            End
            Begin VB.TextBox txtFathersContact 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2940
               TabIndex        =   15
               Top             =   5100
               Width           =   2190
            End
            Begin VB.TextBox txtGrandFathersContact 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   5175
               TabIndex        =   14
               Top             =   5100
               Width           =   2190
            End
            Begin VB.ComboBox cmbMaritalStatus 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "Form1.frx":0742
               Left            =   8385
               List            =   "Form1.frx":0744
               TabIndex        =   13
               Top             =   3220
               Width           =   2055
            End
            Begin VB.TextBox txtFirstName 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3240
               TabIndex        =   12
               Top             =   1335
               Width           =   2325
            End
            Begin VB.TextBox txtMiddleName 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   5625
               TabIndex        =   11
               Top             =   1325
               Width           =   1800
            End
            Begin VB.TextBox txtLastName 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   7470
               TabIndex        =   10
               Top             =   1325
               Width           =   1770
            End
            Begin VB.ComboBox cmbSex 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "Form1.frx":0746
               Left            =   9285
               List            =   "Form1.frx":0756
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1325
               Width           =   1155
            End
            Begin VB.ComboBox cmbBirthPlace 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "Form1.frx":0786
               Left            =   645
               List            =   "Form1.frx":0A66
               TabIndex        =   8
               Top             =   2265
               Width           =   4065
            End
            Begin VB.TextBox txtCitizenshipNo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4785
               TabIndex        =   7
               Top             =   2265
               Width           =   3000
            End
            Begin VB.TextBox txtPassportNo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   6105
               TabIndex        =   6
               Top             =   3220
               Width           =   2070
            End
            Begin VB.TextBox txtCitizenshipIssue 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   7860
               TabIndex        =   5
               Top             =   2265
               Width           =   2520
            End
            Begin VB.TextBox txtDrivingLicenseNo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3930
               TabIndex        =   4
               Top             =   3220
               Width           =   2070
            End
            Begin VB.ComboBox cmbFacility 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "Form1.frx":1618
               Left            =   660
               List            =   "Form1.frx":161A
               TabIndex        =   3
               Top             =   3225
               Width           =   3165
            End
            Begin VB.Label lblID 
               Height          =   225
               Left            =   -68835
               TabIndex        =   63
               Top             =   5910
               Width           =   825
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Employee Code"
               Height          =   195
               Left            =   660
               TabIndex        =   62
               Top             =   870
               Width           =   1110
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Country"
               Height          =   195
               Left            =   -66675
               TabIndex        =   61
               Top             =   1861
               Width           =   585
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "District"
               Height          =   195
               Left            =   -68330
               TabIndex        =   60
               Top             =   1861
               Width           =   495
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Zone"
               Height          =   195
               Left            =   -70030
               TabIndex        =   59
               Top             =   1861
               Width           =   360
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Temporary Address && Contact Information"
               Height          =   195
               Left            =   -74310
               TabIndex        =   58
               Top             =   1861
               Width           =   3060
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Country"
               Height          =   195
               Left            =   -66675
               TabIndex        =   57
               Top             =   915
               Width           =   585
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "District"
               Height          =   195
               Left            =   -68330
               TabIndex        =   56
               Top             =   915
               Width           =   495
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Zone"
               Height          =   195
               Left            =   -70030
               TabIndex        =   55
               Top             =   915
               Width           =   360
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Permanent Address && Contact Information"
               Height          =   195
               Left            =   -74310
               TabIndex        =   54
               Top             =   915
               Width           =   3060
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Residence Contact"
               Height          =   195
               Left            =   -74310
               TabIndex        =   53
               Top             =   2807
               Width           =   1350
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Cell Number"
               Height          =   195
               Left            =   -72040
               TabIndex        =   52
               Top             =   2807
               Width           =   855
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Pager"
               Height          =   195
               Left            =   -69770
               TabIndex        =   51
               Top             =   2807
               Width           =   420
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Email Address"
               Height          =   195
               Left            =   -67500
               TabIndex        =   50
               Top             =   2807
               Width           =   990
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "&Spouse Name"
               Height          =   195
               Left            =   705
               TabIndex        =   49
               Top             =   3750
               Width           =   975
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "&Father's Name"
               Height          =   195
               Left            =   2940
               TabIndex        =   48
               Top             =   3780
               Width           =   1035
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "&Grandfather's Name"
               Height          =   195
               Left            =   5175
               TabIndex        =   47
               Top             =   3780
               Width           =   1440
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "&Contact Number"
               Height          =   195
               Left            =   705
               TabIndex        =   46
               Top             =   4690
               Width           =   1170
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "&Contact Number"
               Height          =   195
               Left            =   2940
               TabIndex        =   45
               Top             =   4690
               Width           =   1170
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "&Contact Number"
               Height          =   195
               Left            =   5175
               TabIndex        =   44
               Top             =   4690
               Width           =   1170
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               Caption         =   "Marital &Status"
               Height          =   195
               Left            =   8385
               TabIndex        =   43
               Top             =   2805
               Width           =   990
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "First Name"
               Height          =   195
               Left            =   3240
               TabIndex        =   42
               Top             =   870
               Width           =   765
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Middle Name"
               Height          =   195
               Left            =   5625
               TabIndex        =   41
               Top             =   870
               Width           =   900
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Last Name"
               Height          =   195
               Left            =   7470
               TabIndex        =   40
               Top             =   870
               Width           =   750
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Sex"
               Height          =   195
               Left            =   9285
               TabIndex        =   39
               Top             =   870
               Width           =   270
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Birth Place"
               Height          =   195
               Left            =   720
               TabIndex        =   38
               Top             =   1845
               Width           =   750
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Citizenship or Social Security Number"
               Height          =   195
               Left            =   4785
               TabIndex        =   37
               Top             =   1855
               Width           =   2640
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Passport Number"
               Height          =   195
               Left            =   6105
               TabIndex        =   36
               Top             =   2805
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Citizenship Issue Place"
               Height          =   195
               Left            =   7860
               TabIndex        =   35
               Top             =   1855
               Width           =   1620
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Driving License Number"
               Height          =   195
               Left            =   3930
               TabIndex        =   34
               Top             =   2805
               Width           =   1665
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Facility"
               Height          =   195
               Left            =   660
               TabIndex        =   33
               Top             =   2820
               Width           =   495
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Abnormal Termination
    End
End Sub

Private Sub Form_Load()
    SubClassSSTAB SSTab1, Picture2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnSubClassSSTAB SSTab1.hwnd
End Sub

Private Sub Label33_Click()

End Sub


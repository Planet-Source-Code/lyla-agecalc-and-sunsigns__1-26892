VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AgeCalc & SunSigns.  By Lyla "
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7800
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Count"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Text            =   "Enter your date"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I hope you like it.
'Enjoy!  :)
'By Lyla. Sept. 2nd, 2001 -->  SaharFx@Yahoo.Com

Function Agecal(myDate As Variant) As Integer
   
   myDate = CDate(Text1)
'* MyDate is the var name for text1.  We used CDate for conversion so you can
'* enter the date as (25/12/1980 Euro dd/mm/yyyy) OR (12/25/1980 American mm/dd/yyyy)
'* OR (1980/12/25 Arabic yyyy/mm/dd)
 
 Dim Totaldays As Long
'*we defined Totaldays as a number(Long) between -2000,000,000 and +2000,000,000
 
 Totaldays = DateDiff("y", myDate, Date)
'* Totaldays = Date Difference("Days", myDate, System's Date)
'* =Total Number of days between Text1's date and Today's date. (Example: 7472 Days)
   
  Numyears = Abs(Totaldays / 365.25)
'* Numyears = Positive(Totaldays / 365.25).  A year actaully has 365.25 Days
'* =positive Number of years between Text1's date and Today's date.(Example: 20.4572 Years)
   
  NumMonths = (Numyears - Int(Numyears)) * 365.25 / 30.4583
'* NumMonths = (Numyears - No fractions(Numyears)) * 365.25 / 30.4583
'*        = (         .4572             ) * 365.25 / 30.4583
'* Number of months between Text1date and Today's date. (Example: 5.4829 Months)
   
  NumDays = CInt((NumMonths - Int(NumMonths)) * 30.4583)
'*   = Round(     .4829            * 30.4583)
'* Number of days.  (Example: 15 Days)


signdate = CInt(Format(myDate, "mmdd")) '* Here we format so Feb. 28th becomes 0228
Select Case signdate
    Case 121 To 219     '*AQUARIUS : (January 21 – February 19)
    sunsign = "AQUARIUS"

    Case 220 To 320     '*PISCES: (February 20 – March 20)
    sunsign = "PISCES"
  
    Case 321 To 420     '*ARIES : (March 21 - April 20 )
    sunsign = "ARIES"
  
    Case 421 To 520     '*TAURUS : (April 21 - May 20)
    sunsign = "TAURUS"
  
    Case 521 To 621     '*GEMINI : (May 21 - June 21)
    sunsign = "GEMINI"
  
    Case 622 To 723     '*CANCER : (June 22 - July 23)
    sunsign = "CANCER"
  
    Case 724 To 823     '*LEO : (July 24 - August 23)
    sunsign = "LEO"
  
    Case 824 To 923     '*VIRGO : (August 24 - September 23)
    sunsign = "VIRGO"
  
    Case 924 To 1022    '*LIBRA : (September 24 - October 22)
    sunsign = "LIBRA"
  
    Case 1023 To 1122   '*SCORPIO : (October 23 - November 22)
    sunsign = "SCORPIO"
  
    Case 1123 To 1222   '*SAGITTARIUS : (November 23 - December 22)
    sunsign = "SAGITTARIUS"
  
    Case Else           '*CAPRICORN : (December 23 - January 20)
    sunsign = "CAPRICORN"
End Select
 
If myDate < Date Then   '* If Our Date is Before Today's Date  ( PAST )
    Label1 = "It has been:    " & Int(Numyears) & "  Year(s),  " & Int(NumMonths) & "  Month(s)  and   " & Int(NumDays) & "  Day(s)" & "  Since that date.  And it falles under: " & sunsign
    
   Else    '* Otherwise, If Our Date is After Today's Date     ( Future )
     
    Label1 = " There are:    " & Int(Numyears) & "  Year(s),  " & Int(NumMonths) & "  Month(s)  and   " & Int(NumDays) & "  Day(s)" & "  To this date.  And it falles under: " & sunsign
End If
Picture1.Picture = LoadPicture(App.Path + "\" & sunsign & ".gif")
End Function

Private Sub Command1_Click()
    If IsDate(Text1) = False Then      '* If the Entered Date is wrong
        MsgBox " Please enter a valid date. "
        Text1.SetFocus                 '* Bring courser back to Text1.
        Text1.SelStart = 0             '* Select first position in Text1.
        Text1.SelLength = Len(Text1)   '* Highlight all Text1.
        Exit Sub
    End If
  Agecal (myDate)                      '* Call the Function
End Sub

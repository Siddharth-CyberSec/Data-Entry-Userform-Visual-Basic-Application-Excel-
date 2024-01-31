INTRODUCTION<br>
Visual Basic for Applications (VBA) is an implementation of Microsoft's event-driven programming language Visual Basic 6, which was declared legacy in 2008, and is an associated integrated development environment (IDE). Although pre-.NET Visual Basic is no longer supported or updated by Microsoft, the VBA programming language was upgraded in 2010 with the introduction of Visual Basic for Applications 7 in Microsoft Office applications. As of 2020, VBA has held its position as "most dreaded" language for developers for 2 years, according to some who participated in surveys undertaken by Stack Overflow. (The most dreaded language for 2018 was Visual Basic 6).<br>
Visual Basic for Applications enables building user-defined functions (UDFs), automating processes and accessing Windows API and other low-level functionality through dynamic-link libraries (DLLs). It supersedes and expands on the abilities of earlier application-specific macro programming languages such as Word's WordBASIC. It can be used to control many aspects of the host application, including manipulating user interface features, such as menus and toolbars, and working with custom user forms or dialog boxes.<br>
As its name suggests, VBA is closely related to Visual Basic and uses the Visual Basic Runtime Library. However, VBA code normally can only run within a host application, rather than as a standalone program. VBA can, however, control one application from another using OLE Automation. For example, VBA can automatically create a Microsoft Word report from Microsoft Excel data that Excel collects automatically from polled sensors. VBA can use, but not create, ActiveX/COM DLLs, and later versions add support for class modules.<br>

USER-FORMS<br>
There are 12 different User-Forms for different parts of the Organization like after the Initial data entry user form there is for Different Regions, Countries, Companies, Businesses, Country Divisions, Country Business Units, Department and Team of Siemens AG.

First User-Form:”frmData”<br>
This user-form initializes and takes in the basic information of the employee like the Id, Name, Gender, Location, E-mail Address, Contact Number and any possible remarks. This user-form does not add just takes in the information to be added later.

Code<br>
'Variable Declaration<br>
Dim BlnVal As Boolean<br>

Private Sub UserForm_Initialize()<br>
    'Variable declaration<br>
    Dim IdVal As Integer<br>
    
    'Finding last row in the Data Sheet<br>
    IdVal = fn_LastRow(Sheets("Data"))<br>
    
    'Update next available id on the userform<br>
    frmData.txtId = IdVal<br>
End Sub<br>
Sub cmdAdd_Click()<br>
    On Error GoTo ErrOccured<br>
    'Boolean Value<br>
    BlnVal = 0<br>
    
    'Data Validation
    Call Data_Validation
    
    'Check validation of all fields are completed are not
    If BlnVal = 0 Then Exit Sub
      
    'TurnOff screen updating
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
     
    'Variable declaration
    Dim txtId, txtName, GenderValue, txtLocation, txtCNum, txtEAddr, txtRemarks
    Dim iCnt As Integer
    
    'find next available row to update data in the data worksheet
    iCnt = fn_LastRow(Sheets("Data")) + 1
    
    'Find Gender value
    If frmData.obMale = True Then
       GenderValue = "Male"
    Else
       GenderValue = "Female"
    End If
    
  
    
    'Display next available Id number on the Userform
    'Variable declaration
    Dim IdVal As Integer
    
    'Finding last row in the Data Sheet
    IdVal = fn_LastRow(Sheets("Data"))
    
    'Update next available id on the userform
    frmData.txtId = IdVal
    
ErrOccured:
    'TurnOn screen updating
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
frmData.Hide
UserForm1.Show

End Sub

'In this example we are finding the last Row of specified Sheet
Function fn_LastRow(ByVal Sht As Worksheet)

    Dim lastRow As Long
    lastRow = Sht.Cells.SpecialCells(xlLastCell).Row
    lRow = Sht.Cells.SpecialCells(xlLastCell).Row
    Do While Application.CountA(Sht.Rows(lRow)) = 0 And lRow <> 1
        lRow = lRow - 1
    Loop
    fn_LastRow = lRow

End Function

'Exit from the Userform
Private Sub cmdCancel_Click()
    Unload Me
End Sub

' Check all the data(except remarks field) has entered are not on the userform
Sub Data_Validation()
     If txtName = "" Then
        MsgBox "Enter Name!", vbInformation, "Name"
        Exit Sub
     ElseIf frmData.obMale = False And frmData.obFMale = False Then
        MsgBox "Select Gender!", vbInformation, "Gender"
        Exit Sub
     ElseIf txtLocation = "" Then
        MsgBox "Enter Location!", vbInformation, "Location"
        Exit Sub
    ElseIf txtEAddr = "" Then
        MsgBox "Enter Address!", vbInformation, "Email Address"
        Exit Sub
    ElseIf txtCNum = "" Then
        MsgBox "Enter Contact Number!", vbInformation, "Contact Number"
        Exit Sub
    Else
        BlnVal = 1
    End If
End Sub

'Clearing data fields of userform
Private Sub cmdClear_Click()
    Application.ScreenUpdating = False
        txtId.Text = ""
        txtName.Text = ""
        obMale.Value = True
        txtLocation = ""
        txtEAddr = ""
        txtCNum = ""
        txtRemarks = ""
    Application.ScreenUpdating = True
End Sub


Second User-Form:”UserForm1”
This User-Form takes the information of the Region Your Country and subsequently the company is situated in that the particular employee works for. As far as I could understand Siemens has divided “FOUR” Regions: Americas, Middle East and Africa, Asia and Oceanic & Europe.

CODE
Private Sub CommandButton2_Click()
    If UserForm1.CheckBox1 = True Then
        UserForm1.Hide
        UserForm2.Show
        Exit Sub
    ElseIf UserForm1.CheckBox2 = True Then
        UserForm1.Hide
        UserForm3.Show
        Exit Sub
    ElseIf UserForm1.CheckBox3 = True Then
        UserForm1.Hide
        UserForm4.Show
        Exit Sub
    ElseIf UserForm1.CheckBox4 = True Then
        UserForm1.Hide
        UserForm5.Show
        Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton1_Click()
    Unload Me
End Sub


Third User-Form:”UserForm2”
This user-form takes the information in which Siemens country company do you work in that particular region. This Particular form shows the Siemens Companies of the “Americas”.

CODE:
'Private Sub CommandButton1_Click()
  '  If UserForm2.CheckBox1 = True Then
  '      UserForm2.Hide
  '     UserForm(n).Show
  '     Exit Sub
  '  ElseIf UserForm2.CheckBox2 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox3 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox4 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox5 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox6 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox7 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox8 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox9 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox10 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox11 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm2.CheckBox12 = True Then
  '      UserForm2.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  End If
'End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub

Fourth User-Form:”UserForm3”
This user-form takes the information in which Siemens country company do you work in that particular region. This Particular form shows the Siemens Companies of the “Middle East and Africa”.

CODE
'Private Sub CommandButton1_Click()
  '  If UserForm3.CheckBox1 = True Then
  '      UserForm3.Hide
  '     UserForm(n).Show
  '     Exit Sub
  '  ElseIf UserForm3.CheckBox2 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox3 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox4 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox5 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox6 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox7 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox8 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox9 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox10 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox11 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox12 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox13 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm3.CheckBox14 = True Then
  '      UserForm3.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  End If
'End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub

Fifth User-Form:”UserForm4”
This user-form takes the information in which Siemens country company do you work in that particular region. This Particular form shows the Siemens Companies of the “Asia and Oceanic”.
Our Country comes in this region “India” (RC-IN), so only the Indian checkbox has been assigned another user form and the others are still comments.

CODE
'Variable Declaration
Dim BlnVal As Boolean

Private Sub CommandButton1_Click()
  '  If UserForm4.CheckBox1 = True Then
  '      UserForm4.Hide
  '     UserForm(n).Show
  '     Exit Sub
  '  ElseIf UserForm4.CheckBox2 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox3 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox4 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox5 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox6 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
    If UserForm4.CheckBox7 = True Then
        UserForm4.Hide
        UserForm6.Show
        Exit Sub
  '  ElseIf UserForm4.CheckBox8 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox9 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox10 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox11 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox12 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox13 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox14 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox15 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox16 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox17 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox18 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox19 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox20 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox21 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox22 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox23 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox24 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox25 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm4.CheckBox26 = True Then
  '      UserForm4.Hide
  '      UserForm(n).Show
  '      Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub


Sixth User-Form:”UserForm5”
This user-form takes the information in which Siemens country company do you work in that particular region. This Particular form shows the Siemens Companies of the “Europe”.

CODE
NOTE: This Code is similar to the previous two, but is exceptionally big because there are 39 countries in Europe with Siemens in them, so it’s a long list.


Seventh User-Form:”UserForm6
In this user form the program ask for which company do you work for in Siemens Ltd. (RC-IN). The Companies are: Energy, Finance, Healthcare, Mobility, Software and Industrial Automation.
As I have only knowledge about the Energy Company (SE) so only the Energy Checkbox is assigned to another user-form, the rest are in the form of comments.

CODE
Private Sub CommandButton1_Click()
   If UserForm6.CheckBox1 = True Then
       UserForm6.Hide
       UserForm7.Show
       Exit Sub
  '  ElseIf UserForm6.CheckBox2 = True Then
  '      UserForm6.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm6.CheckBox3 = True Then
  '      UserForm6.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm6.CheckBox4 = True Then
  '      UserForm6.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm6.CheckBox5 = True Then
  '      UserForm6.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm6.CheckBox6 = True Then
  '      UserForm6.Hide
  '      UserForm(n).Show
  '      Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub

Eighth User-Form:”UserForm7”:
This user form asks for which businesses do you work for under that company, in this scenario only the “Gas and Power” (GP) Checkbox is assigned to another Checkbox.

CODE
Private Sub CommandButton1_Click()
   If UserForm7.CheckBox1 = True Then
       UserForm7.Hide
       UserForm8.Show
       Exit Sub
  '  ElseIf UserForm7.CheckBox2 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox3 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox4 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox5 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox6 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox7 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox8 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox9 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox10 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox11 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm7.CheckBox12 = True Then
  '      UserForm7.Hide
  '      UserForm(n).Show
  '      Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub


Ninth User-Form:”UserForm8”
This user form asks about the country divisions of the particular business you work for like in this scenario in “Gas & Power” the country divisions are: Industrial Application, Transmission or Generation; or either any functions of that business like Finance or Sales, etc. 
	In this user form also not all further divisions are known except for “Generation” (G).

CODE
Private Sub CommandButton1_Click()
  '  If UserForm8.CheckBox1 = True Then
  '     UserForm8.Hide
  '     UserForm(n).Show
  '     Exit Sub
  '  ElseIf UserForm8.CheckBox2 = True Then
  '      UserForm8.Hide
  '      UserForm(n).Show
  '      Exit Sub
    If UserForm8.CheckBox3 = True Then
        UserForm8.Hide
        UserForm9.Show
        Exit Sub
  '  ElseIf UserForm8.CheckBox4 = True Then
  '      UserForm8.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm8.CheckBox5 = True Then
  '      UserForm8.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm8.CheckBox6 = True Then
  '      UserForm8.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm8.CheckBox7 = True Then
  '      UserForm8.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm8.CheckBox8 = True Then
  '      UserForm8.Hide
  '      UserForm(n).Show
  '      Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub



Tenth User-Form:”UserForm9”
This User Form asks about the Countries’ Business Unit you are a part of; in this scenario the only Business Unit which I know of further is “Large Rotating Equipment – Research and Development” (LRE-R&D).


CODE
Private Sub CommandButton1_Click()
  '  If UserForm9.CheckBox1 = True Then
  '     UserForm9.Hide
  '     UserForm(n).Show
  '     Exit Sub
  '  ElseIf UserForm9.CheckBox2 = True Then
  '      UserForm9.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  If UserForm9.CheckBox3 = True Then
  '      UserForm9.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm9.CheckBox4 = True Then
  '      UserForm9.Hide
  '      UserForm(n).Show
  '      Exit Sub
    If UserForm9.CheckBox5 = True Then
        UserForm9.Hide
        UserForm10.Show
        Exit Sub
  '  ElseIf UserForm9.CheckBox6 = True Then
  '      UserForm9.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm9.CheckBox7 = True Then
  '      UserForm9.Hide
  '      UserForm(n).Show
  '      Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub


Eleventh User-Form:”UserForm10”
This Unit asks about the “Department” of the Business Unit you work for like in my case the Department is “Global Customer Operations”, it might as well be a function for that business unit like their Finances or Supply Chain Management.

CODE
Private Sub CommandButton1_Click()
    If UserForm10.CheckBox1 = True Then
       UserForm10.Hide
       UserForm11.Show
       Exit Sub
  '  ElseIf UserForm10.CheckBox2 = True Then
  '      UserForm10.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  If UserForm10.CheckBox3 = True Then
  '      UserForm10.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm10.CheckBox4 = True Then
  '      UserForm10.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm10.CheckBox5 = True Then
  '      UserForm10.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm10.CheckBox6 = True Then
  '      UserForm10.Hide
  '      UserForm(n).Show
  '      Exit Sub
  '  ElseIf UserForm10.CheckBox7 = True Then
  '      UserForm10.Hide
  '      UserForm(n).Show
  '      Exit Sub
    End If
End Sub

'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub


Twelfth User-Form:”UserForm11”
This user form is the final stage before addition of the record to the database, this user form asks for which team of the department you work for, and in my case it is the “Fluid Mechanical Systems”.
This user form will now finally Add the information to the database, the Region, Country, Company, Country Division, Business Unit, Department, Team all of these are shown in the form of their codes, like in my case the entire chain of command code is “RC-IN SE GP G LRE-R&D GCO FMS”


CODE
'Variable Declaration
Dim BlnVal As Boolean

Private Sub UserForm_Initialize()
    'Variable declaration
    Dim IdVal As Integer
    
    'Finding last row in the Data Sheet
    IdVal = fn_LastRow(Sheets("Data"))
    
    'Update next available id on the userform
    frmData.txtId = IdVal
End Sub

Sub cmdAdd_Click()
   On Error GoTo ErrOccured
    'Boolean Value
    BlnVal = 0

    'TurnOff screen updating
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
     
    'Variable declaration
    Dim txtId, txtName, GenderValue, txtLocation, txtCNum, txtEAddr, txtRemarks, CountryVal, CompanyVal, BusinessVal, CountryDivisionVal, CountryBusinessUnitVal, DepartmentVal, TeamVal
    Dim iCnt As Integer
    
    'find next available row to update data in the data worksheet
    iCnt = fn_LastRow(Sheets("Data")) + 1
    
    'Find Gender value
    If frmData.obMale = True Then
       GenderValue = "Male"
    Else
       GenderValue = "Female"
    End If
    
    'Find Country Name
    If UserForm4.CheckBox7 = True Then
        CountryVal = "RC-IN"
    Else
        CountryVal = ""
    End If
    
    'Find Company Name
    If UserForm6.CheckBox1 = True Then
        CompanyVal = "SE"
    Else
        CompanyVal = ""
    End If
    
    'Find Business Name
    If UserForm7.CheckBox1 = True Then
        BusinessVal = "GP"
    Else
        BusinessVal = ""
    End If
    
     'Find Country Division Name
    If UserForm8.CheckBox3 = True Then
        CountryDivisionVal = "G"
    Else
        CountryDivisionVal = ""
    End If
    
     'Find Country Business Unit Name
    If UserForm9.CheckBox5 = True Then
        CountryBusinessUnitVal = "LRE-R&D"
    Else
        CountryBusinessUnitVal = ""
    End If
    
     'Find Department Name
    If UserForm10.CheckBox1 = True Then
        DepartmentVal = "GCO"
    Else
        DepartmentVal = ""
    End If
    
     'Find Team Name
    If UserForm11.CheckBox2 = True Then
        TeamVal = "FMS"
    Else
        TeamVal = ""
    End If
    
    'Update userform data to the Data Worksheet
    With Sheets("Data")
        .Cells(iCnt, 1) = iCnt - 1
        .Cells(iCnt, 2) = frmData.txtName
        .Cells(iCnt, 3) = GenderValue
        .Cells(iCnt, 4) = frmData.txtLocation.Value
        .Cells(iCnt, 5) = frmData.txtEAddr
        .Cells(iCnt, 6) = frmData.txtCNum
        .Cells(iCnt, 7) = frmData.txtRemarks
        .Cells(iCnt, 8) = CountryVal
        .Cells(iCnt, 9) = CompanyVal
        .Cells(iCnt, 10) = BusinessVal
        .Cells(iCnt, 11) = CountryDivisionVal
        .Cells(iCnt, 12) = CountryBusinessUnitVal
        .Cells(iCnt, 13) = DepartmentVal
        .Cells(iCnt, 14) = TeamVal
        
        'Diplay headers on the first row of Data Worksheet
        If .Range("A1") = "" Then
            .Cells(1, 1) = "Id"
            .Cells(1, 2) = "Name               "
            .Cells(1, 3) = "Gender"
            .Cells(1, 4) = "Location"
            .Cells(1, 5) = "Email Addres              "
            .Cells(1, 6) = "Contact Number"
            .Cells(1, 7) = "Remarks"
            .Cells(1, 8) = "Country"
            .Cells(1, 9) = "Company"
            .Cells(1, 10) = "Business"
            .Cells(1, 11) = "Country Div."
            .Cells(1, 12) = "Coun. Bus. Unit"
            .Cells(1, 13) = "Department"
            .Cells(1, 14) = "Team"
            
            'Formatiing Data
            .Columns("A:N").Columns.AutoFit
            .Range("A1:N1").Font.Bold = True
            .Range("A1:N1").LineStyle = xlDash
            
        End If
    End With
    
    'Display next available Id number on the Userform
    'Variable declaration
    Dim IdVal As Integer
    
    'Finding last row in the Data Sheet
    IdVal = fn_LastRow(Sheets("Data"))
    
    'Update next available id on the userform
    frmData.txtId = IdVal
    
ErrOccured:
    'TurnOn screen updating
    Application.ScreenUpdating = True
    Application.EnableEvents = True


End Sub

'In this example we are finding the last Row of specified Sheet
Function fn_LastRow(ByVal Sht As Worksheet)

    Dim lastRow As Long
    lastRow = Sht.Cells.SpecialCells(xlLastCell).Row
    lRow = Sht.Cells.SpecialCells(xlLastCell).Row
    Do While Application.CountA(Sht.Rows(lRow)) = 0 And lRow <> 1
        lRow = lRow - 1
    Loop
    fn_LastRow = lRow

End Function
'Exit from the Userform
Private Sub CommandButton2_Click()
   Unload Me
End Sub



Modules
As this entire project is based on user forms the module portion only take cares of the basic appearance of the workbook.

CODE:
Sub Oval2_Click()
    frmData.Show
End Sub
Sub Clear_DataSheet()
    Sheets("Data").Columns("A:N").Clear
End Sub





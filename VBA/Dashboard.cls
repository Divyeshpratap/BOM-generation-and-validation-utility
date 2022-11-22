VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This VBA utility aims at Generating End User understandable Sub Assembly BOM from the PSV value (refer annexures for more details on what is PSV file type, and how is it generated)
' provided by the user.
' Utility Functionality
' Input Variable -> AutoBOM generated PSV file of Main Assembly, TcM id of assemblies for which sub assembly BOM is required by User
' Output -> BOM of Main assembly, Separate BOM of individual assemblies


' Steps for generating Sub Assembly BOM are:
' 1. User inputs the file location of PSV file in sheet: , cell number:
' 2. User presses command button ( Generate BOM).
' 3. Utility goes into processing mode and successfull BOM generation is informed to user through a msgbox "BOM generation complete".


' Note: VBA Macros should be enabled in excel instance. Minimum 16GB RAM required. MS office version 2010 or higher required. All other excel files should be closed
' before running the utility. Multiple instances of excel shoould not be opened.


' To understand the functionality and troubleshooting the code, it is necessary to get a deep understanding of the input (PSV file) and internal assembly structure
' of a Solid Edge Assembly file.  This section aims at providing a brief overview of the same. A Solid Edge assembly is generated at multiple levels where the Main assembly is given the lowest level of 1 and subsequent parts\ assemblies
' are assigned levels in ascending order. These levels are defined by software to maintain the parent\ child relationship of various components used in Main assembly.
' These levels are automatically assigned by SE, and a thorough understanding of assembly structure can be realized using TcM structure manager.

' To generate BOM in excel format which is a part of ATE workflow; it is necessary that the part list of components is fetched from within Solid Edge and is stored in an intermediate file.
' Here, for this particular purpose a utility developed by Siemens is used which can retreive part list from SE environment using SE API's embedded in visual basic.
' This utility retreives certain parameters (these parameters can be referred from psv file first row) of all parts used in main assembly in a sequential manner as per their sequential arrangement in structure manager.
' After running the Siemens AutoBOM utility, the PSV file is by default generated and stored at "D:\BOM_Report_Files location" in respective PC. PSV is a pipe separated file type which is similar to
' ".csv" file with the only difference being the deliminator in ".psv" file type is a pipe "|" character.

' This utility makes use of the same PSV file and converts it into a user readable and manipulative excel document for further reference and communication.

' Steps internally executed by PSV file are:
' 1. Utility imports psv file into excel into a separate sheet "Psv_values".
' 2. Utility reads the list of sub assembly id provided by user, checks for correctness of assembly id and then generates seperate "Sub_psv" sheet for the first sub assembly.
' 3. For the first sub assembly, code starts segregating the BOM into separate sheets ( "Maintenance", "BOP", "Machined", "Inhouse")  based on their ALB component type.
' 4. The values are populated and all the four sheets are tranferred to a new workbook, which gets placed in the sub directory "Sub" and get renamed as per the TcM id.
' 5. The process is repeated for all sub assemblies for which separate BOM was requested by the user.
' 6. After all sub assembly BOM's are generatd utility moves to generate the BOM of master assembly. The master assembly BOM is the main assembly BOM minus the sum of individual sub assembly BOM.
' 7. This is performed by eliminating the sub assembly values from original "Psv_values" sheet. Master asembly BOM is generated and exported within the same directory.
' 8. After completion of all above sub routines, notification for process complete is shown.
  
' For any iterations in utility contact: ATE design

' -Divyesh
' 28th July 2021


Dim sub_name As String
Dim coun As String
Dim count As String
Dim counts As String

Private Sub screen_update()
    Application.ScreenUpdating = True
End Sub


Private Sub CommandButton19_Click()
    myfilepath = Application.GetOpenFilename()
    Range("C7").Value = myfilepath
End Sub

'This sub routine imports ".psv" file type into excel sheet (Psv_Values) and seperates the columns using "|" variable.

Private Sub CommandButton1_Click()

     
     Dim max_rows As String

     Dim CurSheet As Worksheet
     Set CurSheet = ActiveSheet
     
     psvfilepath = Range("C7").Value
     psvfilepath = "TEXT;" & psvfilepath
     Application.ScreenUpdating = False 'Used to increase code speed
     Worksheets("Template").Copy After:=Worksheets(Worksheets.count) 'Copying from a template increases speed and any further modifications are easier
     ActiveSheet.Name = "Psv_Values"
     With ActiveSheet.QueryTables.Add(Connection:=psvfilepath, _
          Destination:=ActiveSheet.Range("A1")) ' Location of input PSV file
         .TextFileParseType = xlDelimited
         .TextFileOtherDelimiter = "|" 'Delimiter information. Can be updated to any delimiter as per requirement
         .Refresh BackgroundQuery:=False
     End With
     
'      With ActiveSheet.QueryTables.Add(Connection:= _
'          "TEXT;D:\BOM_Report_Files\ATE_BOM_Report.psv", Destination:=ActiveSheet.Range("A1")) ' Location of input PSV file
'         .TextFileParseType = xlDelimited
'         .TextFileOtherDelimiter = "|" 'Delimiter information. Can be updated to any delimiter as per requirement
'         .Refresh BackgroundQuery:=False
'      End With
     
     CurSheet.Activate
     
     Sheets("Psv_Values").Range("A:A").Copy Sheets("Psv_Values").Range("V:V")

     'Following code is required to check if there is some error in the psv file generated by the AutoBOM utility.
     'Two conditions are checked; if there are empty lines in the psv file, and if the starting characters are other than ALB
     'If any errors are found by the code, they have to be manually removed from the psv file by the administrator
          
     max_rows = 2
     Do Until ((Sheets("Psv_Values").Range("A" & max_rows).Value = "") Or (Left(Sheets("Psv_Values").Range("A" & max_rows), 3) <> "ALB"))
     max_rows = max_rows + 1
     Loop
     If Sheets("Psv_Values").Range("A" & max_rows).Value = "" Then
        If ((Sheets("Psv_Values").Range("A" & max_rows + 1).Value = "") And (Sheets("Psv_Values").Range("A" & max_rows + 2).Value = "") And (Sheets("Psv_Values").Range("A" & max_rows + 3).Value = "")) Then
           MsgBox "No error detected in PSV file and total number of components are" & max_rows - 2, vbInformation
        Else
            MsgBox "1st Type Error detected in  PSV file at row number" & max_rows, vbInformation
        End If
     ElseIf (Left(Sheets("Psv_Values").Range("A" & max_rows), 3) <> "ALB") Then
        MsgBox "2nd Type Error detected in  PSV file at row number" & max_rows, vbInformation
        
     Else
           MsgBox "No error detected in PSV file and total number of components are" & max_rows - 2, vbInformation
     End If
     Application.ScreenUpdating = True
         
End Sub







Private Sub CommandButton2_Click()
    
'This sub routine is used to import ALB comp type "Maintenance" as per values in columns "C"

    Dim i, j, total As Integer
    Dim rowb, rowm, rowe, row As Double
    Dim str, str1 As String
    row = 1
    rowb = 2
    rowm = 2
    rowe = 5
'    count = 1
    Sheets("Master_Maintenance").Range("F3").Value = Application.UserName
    Sheets("Master_Maintenance").Range("F2").Value = Date
'    Do Until Sheets("Psv_Values").Range("A" & count).Value = ""
'    count = count + 1
'    Loop
    
    
    Do Until Sheets("Psv_Values").Range("A" & row).Value = ""
    
    If Sheets("Psv_Values").Range("C" & row).Value = "Maintenance" Then                                     'Used for importing only the maintenance comp type
        Sheets("Master_Maintenance").Range("B" & rowe).Value = Sheets("Psv_Values").Range("L" & row).Value
        Sheets("Master_Maintenance").Range("C" & rowe).Value = Sheets("Psv_Values").Range("M" & row).Value
        Sheets("Master_Maintenance").Range("D" & rowe).Value = Sheets("Psv_Values").Range("H" & row).Value
        Sheets("Master_Maintenance").Range("E" & rowe).Value = Sheets("Psv_Values").Range("S" & row).Value
        Sheets("Master_Maintenance").Range("F" & rowe).Value = "=COUNTIF(Psv_Values!A2:A" & count & ",L" & rowe & ")"
        Sheets("Master_Maintenance").Range("G" & rowe).Value = 0
        Sheets("Master_Maintenance").Range("H" & rowe).Formula = "=Sum(F" & rowe & ":G" & rowe & ")"
        Sheets("Master_Maintenance").Range("L" & rowe).Value = Sheets("Psv_Values").Range("V" & row).Value
' Update made in above line in version v4 on 3rd august 2021
        Sheets("Master_Maintenance").Range("A" & rowe).Value = rowe - 4
        
        If Sheets("Master_Maintenance").Range("F" & rowe).Value = "0" Then
            Sheets("Master_Maintenance").Rows(rowe).EntireRow.delete
        End If
' Above If statement added in version v4 on 3rd august 2021

        rowe = rowe + 1
    End If
    ' Next lines of codes for simple formatting and print setup
    row = row + 1
    Loop
    With Sheets("Master_Maintenance").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Maintenance").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Master_Maintenance").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Maintenance").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Maintenance").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Master_Maintenance").Range("A1", "K" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Master_Maintenance").PageSetup.PrintArea = Worksheets("Master_Maintenance").Range("A1", "K" & rowe).Address
    
    

End Sub

Private Sub CommandButton3_Click()


'This sub routine is used to import ALB comp type "Machined" as per values in columns "C"
    Dim i, j, total As Integer
    Dim rowb, rowm, rowe, row As Double
    Dim str, str1 As String

    row = 1
    rowb = 2
    rowm = 2
    rowe = 5

    Sheets("Master_Machined_Component").Range("G3").Value = Application.UserName
    Sheets("Master_Machined_Component").Range("G2").Value = Date


    
    
    Do Until Sheets("Psv_Values").Range("A" & row).Value = ""
    
    If Sheets("Psv_Values").Range("C" & row).Value = "Machined_Component" Then                              'Used for importing only the machined comp type
        Sheets("Master_Machined_Component").Range("B" & rowe).Value = Sheets("Psv_Values").Range("V" & row).Value
        ' Update made in above line in version v4 on 3rd august 2021
        Sheets("Master_Machined_Component").Range("C" & rowe).Value = Sheets("Psv_Values").Range("M" & row).Value
        Sheets("Master_Machined_Component").Range("D" & rowe).Value = Sheets("Psv_Values").Range("G" & row).Value
        Sheets("Master_Machined_Component").Range("E" & rowe).Value = Sheets("Psv_Values").Range("P" & row).Value
        Sheets("Master_Machined_Component").Range("F" & rowe).Value = Sheets("Psv_Values").Range("E" & row).Value
        Sheets("Master_Machined_Component").Range("G" & rowe).Value = Sheets("Psv_Values").Range("N" & row).Value
        Sheets("Master_Machined_Component").Range("H" & rowe).Value = "=COUNTIF(Psv_Values!A2:A" & count & ",B" & rowe & ")"
        Sheets("Master_Machined_Component").Range("I" & rowe).Value = 0
        Sheets("Master_Machined_Component").Range("J" & rowe).Formula = "=Sum(H" & rowe & ":I" & rowe & ")"
        Sheets("Master_Machined_Component").Range("A" & rowe).Value = rowe - 4
        If Sheets("Master_Machined_Component").Range("H" & rowe).Value = "0" Then
            Sheets("Master_Machined_Component").Rows(rowe).EntireRow.delete
        End If
        ' Above If statement added in version v4 on 3rd august 2021
        rowe = rowe + 1
    End If
    
    ' Next line of codes for simple formatting and print setup
    
    row = row + 1
    Loop
    With Sheets("Master_Machined_Component").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("G" & rowe, "J" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("K" & rowe, "M" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Machined_Component").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("G" & rowe, "J" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Master_Machined_Component").Range("K" & rowe, "M" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Machined_Component").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("G" & rowe, "J" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Machined_Component").Range("K" & rowe, "M" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Master_Machined_Component").Range("A1", "M" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Master_Machined_Component").PageSetup.PrintArea = Worksheets("Master_Machined_Component").Range("A1", "M" & rowe).Address
    
    



End Sub

Private Sub CommandButton4_Click()

'This sub routine is used to import ALB comp type "Fabricated" as per values in columns "C"

    Dim i, j, total As Integer
    Dim rowe, row As Double
    Dim str, str1 As String

    row = 1
    rowe = 6

    Sheets("Master_Fabricated").Range("G3").Value = Application.UserName
    Sheets("Master_Fabricated").Range("G2").Value = Date
    
    
    Do Until Sheets("Psv_Values").Range("A" & row).Value = ""
    
    If Sheets("Psv_Values").Range("C" & row).Value = "Fabricated" Then                              'Used for importing only the Fabricated comp type
        Sheets("Master_Fabricated").Range("B" & rowe).Value = Sheets("Psv_Values").Range("V" & row).Value
        ' Update made in above line in version v4 on 3rd august 2021
        Sheets("Master_Fabricated").Range("C" & rowe).Value = Sheets("Psv_Values").Range("M" & row).Value
        Sheets("Master_Fabricated").Range("D" & rowe).Value = Sheets("Psv_Values").Range("G" & row).Value
        Sheets("Master_Fabricated").Range("E" & rowe).Value = "=COUNTIF(Psv_Values!A2:A" & count & ",B" & rowe & ")"
        Sheets("Master_Fabricated").Range("A" & rowe).Value = rowe - 5
        If Sheets("Master_Fabricated").Range("E" & rowe).Value = "0" Then
            Sheets("Master_Fabricated").Rows(rowe).EntireRow.delete
        End If
        ' Above If statement added in version v4 on 3rd august 2021
        rowe = rowe + 1
    End If
    
    row = row + 1
    Loop
    
    ' Next line of codes for simple formatting and print setup
    
    With Sheets("Master_Fabricated").Range("A" & rowe, "B" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("D" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("E" & rowe, "G" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Fabricated").Range("A" & rowe, "B" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("D" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Master_Fabricated").Range("E" & rowe, "G" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Fabricated").Range("A" & rowe, "B" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("D" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Fabricated").Range("E" & rowe, "G" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Master_Fabricated").Range("A1", "G" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Master_Fabricated").PageSetup.PrintArea = Worksheets("Master_Fabricated").Range("A1", "G" & rowe).Address
    
    




End Sub

Private Sub CommandButton5_Click()


'This sub routine is used to import ALB comp type "Bought Out" as per values in columns "C"

    Dim i, j, total As Integer
    Dim rowe, row As Double
    Dim str, str1 As String

    row = 1
    rowe = 5

    Sheets("Master_Bought_Out").Range("F3").Value = Application.UserName
    Sheets("Master_Bought_Out").Range("F2").Value = Date
 
    
    
    Do Until Sheets("Psv_Values").Range("A" & row).Value = ""
    
    If Sheets("Psv_Values").Range("C" & row).Value = "Bought_Out" Then                              'Used for importing only the Bought Out comp type
        Sheets("Master_Bought_Out").Range("B" & rowe).Value = Sheets("Psv_Values").Range("L" & row).Value
        Sheets("Master_Bought_Out").Range("C" & rowe).Value = Sheets("Psv_Values").Range("M" & row).Value
        Sheets("Master_Bought_Out").Range("D" & rowe).Value = Sheets("Psv_Values").Range("H" & row).Value
        Sheets("Master_Bought_Out").Range("E" & rowe).Value = Sheets("Psv_Values").Range("S" & row).Value
        Sheets("Master_Bought_Out").Range("F" & rowe).Value = "=COUNTIF(Psv_Values!A2:A" & count & ",L" & rowe & ")"
        Sheets("Master_Bought_Out").Range("G" & rowe).Value = 0
        Sheets("Master_Bought_Out").Range("H" & rowe).Formula = "=Sum(F" & rowe & ":G" & rowe & ")"
        Sheets("Master_Bought_Out").Range("L" & rowe).Value = Sheets("Psv_Values").Range("V" & row).Value
        ' Update made in above line in version v4 on 3rd august 2021
        Sheets("Master_Bought_Out").Range("A" & rowe).Value = rowe - 4
        If Sheets("Master_Bought_Out").Range("F" & rowe).Value = "0" Then
            Sheets("Master_Bought_Out").Rows(rowe).EntireRow.delete
        End If
        ' Above If statement added in version v4 on 3rd august 2021
        rowe = rowe + 1
    End If
    
    row = row + 1
    Loop
    
    ' Next line of codes for simple formatting and print setup
    
    With Sheets("Master_Bought_Out").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Bought_Out").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Master_Bought_Out").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Master_Bought_Out").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Master_Bought_Out").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Master_Bought_Out").Range("A1", "K" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Master_Bought_Out").PageSetup.PrintArea = Worksheets("Master_Bought_Out").Range("A1", "K" & rowe).Address
    Application.ScreenUpdating = True



End Sub
' This subroutine is used to delete sub assemblies for which BOM has been seperately generated
Private Sub CommandButton6_Click()

Dim counte As String
Dim as_num  As String
counte = 8
Dim counter As String
counter = 2
Dim level As String
level = 1
Dim counter_two As String

Do Until Sheets("Assembly_numbers").Range("B" & counte).Value = ""
    as_num = Sheets("Assembly_numbers").Range("B" & counte).Value
    Sheets("Assembly_numbers").Range("J" & counte).Value = as_num
        Do Until counter = count
            If Sheets("Psv_Values").Range("A" & counter).Value = as_num Then
            level = Sheets("Psv_Values").Range("B" & counter).Value
            Sheets("Psv_Values").Range("A" & counter).Value = "Match found"
            Sheets("Psv_Values").Range("C" & counter).Value = ""
            counter_two = counter + 1
            Do Until level >= Sheets("Psv_Values").Range("B" & counter_two).Value
                Sheets("Psv_Values").Range("A" & counter_two).Value = "Sub assembly deleted"
 '               Sheets("Psv_Values").Range("C" & counter_two).Value = ""
 ' Above line removed in version v4 code on 3rd August, For testing purpose
                counter_two = counter_two + 1
            Loop
            End If
            counter = counter + 1
            level = 1
        Loop
  
        
    counter = 2
    counte = counte + 1

Loop

Sheets("Assembly_numbers").Range("K" & counte).Value = counte - 8

End Sub

Private Sub CommandButton7_Click()

    Call CommandButton1_Click
    Call CommandButton6_Click
    Call CommandButton2_Click
    Call CommandButton3_Click
    Call CommandButton4_Click
    Call CommandButton5_Click
    
  
    
End Sub

Private Sub CommandButton8_Click()

Application.ScreenUpdating = False
'Dim count As String
Dim as_num As String

Dim counter As String
counter = 2
Dim counter_two As String
counter_two = 2
Dim counter_three As String

Dim level As String
level = 1

Dim count_dup As String
count_dup = 2
Dim count_dup2 As String
count_dup2 = 1
'count_dup initial value = 2 has to be checked, Infact for all variables as on version v4 4th August

Dim CurSheet As Worksheet
Set CurSheet = ActiveSheet
Application.ScreenUpdating = False
Worksheets("Template").Copy After:=Worksheets(Worksheets.count)
ActiveSheet.Name = "Psv_Values_Sub"
    ' With ActiveSheet
    

Do Until counter = count
    If Sheets("Psv_Values").Range("A" & counter).Value = sub_name Then
        level = Sheets("Psv_Values").Range("B" & counter).Value
        counter_three = counter + 1
        counter_four = counter + 1
        Do Until level >= Sheets("Psv_Values").Range("B" & counter_three).Value
            Sheets("Psv_Values_Sub").Rows(counter_two).Value = Sheets("Psv_Values").Rows(counter_three).Value
            
            'Sheets("Psv_Values_Sub").Range("W" & counter_two).Value = "= Vlookup(Psv_Values_Sub!A" & counter_two & ",Psv_Values!A2:C" & count & ",3, FALSE)"
            counter_two = counter_two + 1
            counter_three = counter_three + 1
            
        Loop
    End If
counter = counter + 1
Loop
counts = 2
' counts = 1 , To be checked
Do Until Sheets("Psv_Values_Sub").Range("A" & counts).Value = ""
counts = counts + 1
Loop

Dim CurSheet2 As Worksheet
Set CurSheet2 = ActiveSheet
Application.ScreenUpdating = False
Worksheets("Template").Copy After:=Worksheets(Worksheets.count)
ActiveSheet.Name = "Psv_Values_Sub_Temp"
    ' With ActiveSheet

Sheets("Psv_Values_Sub").Range("A2:A" & counts).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Psv_Values_Sub_Temp").Range("A2"), Unique:=True

Do Until count_dup = counts
    Do Until count_dup2 = count
        If Sheets("Psv_Values_Sub_Temp").Range("A" & count_dup).Value = Sheets("Psv_Values").Range("A" & count_dup2).Value Then
        Sheets("Psv_Values_Sub_Temp").Rows(count_dup).Value = Sheets("Psv_Values").Rows(count_dup2).Value
        'Sheets("Psv_Values_Sub").Range("Y" & count_dup).Value = "= Vlookup(Psv_Values_Sub!X" & count_dup & ",Psv_Values!A2:C" & count & ",3, FALSE)"
        End If
        count_dup2 = count_dup2 + 1
    Loop
    count_dup = count_dup + 1
Loop
' Above nested loop added on 4th August version v4
' Added on 4th August version v4 above 2 lines

Sheets("Psv_Values_Sub").Rows(1).EntireRow.delete
'This above line has to be validated before final release of utility as on 3rd August 2021

Application.ScreenUpdating = True
End Sub


Private Sub CommandButton9_Click()


'This sub routine is used to import ALB comp type "Maintenance" as per values in columns "C" from "Psv_Values_Sub"

    Dim i, j, total As Integer
    Dim rowe, row As Double
    Dim str, str1 As String

    row = 2
    'row = 1 previously, updated on 4th august, value output to be checked
    rowe = 5

    Sheets("Sub_Maintenance").Range("F3").Value = Application.UserName
    Sheets("Sub_Maintenance").Range("F2").Value = Date
    

    
    
    Do Until Sheets("Psv_Values_Sub_Temp").Range("A" & row).Value = ""
    
    If Sheets("Psv_Values_Sub_Temp").Range("C" & row).Value = "Maintenance" Then                         'Used for importing only the Maintenance comp type
    Sheets("Sub_Maintenance").Range("B" & rowe).Value = Sheets("Psv_Values_Sub_Temp").Range("L" & row).Value
    Sheets("Sub_Maintenance").Range("C" & rowe).Value = Sheets("Psv_Values_Sub_Temp").Range("M" & row).Value
    Sheets("Sub_Maintenance").Range("D" & rowe).Value = Sheets("Psv_Values_Sub_Temp").Range("H" & row).Value
    Sheets("Sub_Maintenance").Range("E" & rowe).Value = Sheets("Psv_Values_Sub_Temp").Range("S" & row).Value
    Sheets("Sub_Maintenance").Range("F" & rowe).Value = "=COUNTIF(Psv_Values_Sub!A2:A" & counts & ",L" & rowe & ")"
    Sheets("Sub_Maintenance").Range("G" & rowe).Value = 0
    Sheets("Sub_Maintenance").Range("H" & rowe).Formula = "=Sum(F" & rowe & ":G" & rowe & ")"
    Sheets("Sub_Maintenance").Range("L" & rowe).Value = Sheets("Psv_Values_Sub_Temp").Range("A" & row).Value
    Sheets("Sub_Maintenance").Range("A" & rowe).Value = rowe - 4
    
    rowe = rowe + 1
    End If
    
    row = row + 1
    Loop
    
    ' Next line of codes for simple formatting and print setup
    
    With Sheets("Sub_Maintenance").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Maintenance").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Sub_Maintenance").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Maintenance").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Maintenance").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Sub_Maintenance").Range("A1", "K" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Sub_Maintenance").PageSetup.PrintArea = Worksheets("Sub_Maintenance").Range("A1", "K" & rowe).Address
    
  
End Sub

Private Sub CommandButton10_Click()

'This sub routine is used to import ALB comp type "Machined" as per values in columns "C" from "Psv_Values_Sub"


    Dim i, j, total As Integer
    Dim rowe, row As Double
    Dim str, str1 As String

    row = 1
    rowe = 5

    Sheets("Sub_Machined_Component").Range("G3").Value = Application.UserName
    Sheets("Sub_Machined_Component").Range("G2").Value = Date

    
    
    Do Until Sheets("Psv_Values_Sub").Range("X" & row).Value = ""
    
    If Sheets("Psv_Values_Sub").Range("Y" & row).Value = "Machined_Component" Then                          'Used for importing only the Machined comp type
    Sheets("Sub_Machined_Component").Range("B" & rowe).Value = Sheets("Psv_Values_Sub").Range("A" & row).Value
    Sheets("Sub_Machined_Component").Range("C" & rowe).Value = Sheets("Psv_Values_Sub").Range("M" & row).Value
    Sheets("Sub_Machined_Component").Range("D" & rowe).Value = Sheets("Psv_Values_Sub").Range("G" & row).Value
    Sheets("Sub_Machined_Component").Range("E" & rowe).Value = Sheets("Psv_Values_Sub").Range("P" & row).Value
    Sheets("Sub_Machined_Component").Range("F" & rowe).Value = Sheets("Psv_Values_Sub").Range("E" & row).Value
    Sheets("Sub_Machined_Component").Range("G" & rowe).Value = Sheets("Psv_Values_Sub").Range("N" & row).Value
    Sheets("Sub_Machined_Component").Range("H" & rowe).Value = "=COUNTIF(Psv_Values_Sub!A2:A" & counts & ",B" & rowe & ")"
    Sheets("Sub_Machined_Component").Range("I" & rowe).Value = 0
    Sheets("Sub_Machined_Component").Range("J" & rowe).Formula = "=Sum(H" & rowe & ":I" & rowe & ")"
    Sheets("Sub_Machined_Component").Range("A" & rowe).Value = rowe - 4
    
    rowe = rowe + 1
    End If
    
    row = row + 1
    Loop
    With Sheets("Sub_Machined_Component").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("G" & rowe, "J" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("K" & rowe, "M" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Machined_Component").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("G" & rowe, "J" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Sub_Machined_Component").Range("K" & rowe, "M" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Machined_Component").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("G" & rowe, "J" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Machined_Component").Range("K" & rowe, "M" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Sub_Machined_Component").Range("A1", "M" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Sub_Machined_Component").PageSetup.PrintArea = Worksheets("Sub_Machined_Component").Range("A1", "M" & rowe).Address
    
    
    
    
    



End Sub

Private Sub CommandButton11_Click()

'This sub routine is used to import ALB comp type "Fabricated" as per values in columns "C" from "Psv_Values_Sub"


    Dim i, j, total As Integer
    Dim rowe, row As Double
    Dim str, str1 As String

    row = 1
    rowe = 6

    Sheets("Sub_Fabricated").Range("G3").Value = Application.UserName
    Sheets("Sub_Fabricated").Range("G2").Value = Date

    
    
    Do Until Sheets("Psv_Values_Sub").Range("X" & row).Value = ""
    
    If Sheets("Psv_Values_Sub").Range("Y" & row).Value = "Fabricated" Then                          'Used for importing only the Fabricated comp type
    Sheets("Sub_Fabricated").Range("B" & rowe).Value = Sheets("Psv_Values_Sub").Range("A" & row).Value
    Sheets("Sub_Fabricated").Range("C" & rowe).Value = Sheets("Psv_Values_Sub").Range("M" & row).Value
    Sheets("Sub_Fabricated").Range("D" & rowe).Value = Sheets("Psv_Values_Sub").Range("G" & row).Value
    Sheets("Sub_Fabricated").Range("E" & rowe).Value = "=COUNTIF(Psv_Values_Sub!A2:A" & counts & ",B" & rowe & ")"
    Sheets("Sub_Fabricated").Range("A" & rowe).Value = rowe - 5
    
    rowe = rowe + 1
    End If
    
    row = row + 1
    Loop
    With Sheets("Sub_Fabricated").Range("A" & rowe, "B" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("D" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("E" & rowe, "G" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Fabricated").Range("A" & rowe, "B" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("D" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Sub_Fabricated").Range("E" & rowe, "G" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Fabricated").Range("A" & rowe, "B" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("D" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Fabricated").Range("E" & rowe, "G" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Sub_Fabricated").Range("A1", "G" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Sub_Fabricated").PageSetup.PrintArea = Worksheets("Sub_Fabricated").Range("A1", "G" & rowe).Address
    
    






End Sub

Private Sub CommandButton12_Click()

'This sub routine is used to import ALB comp type "Bought Out" as per values in columns "C" from "Psv_Values_Sub"


    Dim i, j, total As Integer
    Dim rowe, row As Double
    Dim str, str1 As String

    row = 1
    rowe = 5
 
    Sheets("Sub_Bought_Out").Range("F3").Value = Application.UserName
    Sheets("Sub_Bought_Out").Range("F2").Value = Date

    
    Do Until Sheets("Psv_Values_Sub").Range("X" & row).Value = ""
    
    If Sheets("Psv_Values_Sub").Range("Y" & row).Value = "Bought_Out" Then                          'Used for importing only the Bought Out comp type
    Sheets("Sub_Bought_Out").Range("B" & rowe).Value = Sheets("Psv_Values_Sub").Range("L" & row).Value
    Sheets("Sub_Bought_Out").Range("C" & rowe).Value = Sheets("Psv_Values_Sub").Range("M" & row).Value
    Sheets("Sub_Bought_Out").Range("D" & rowe).Value = Sheets("Psv_Values_Sub").Range("H" & row).Value
    Sheets("Sub_Bought_Out").Range("E" & rowe).Value = Sheets("Psv_Values_Sub").Range("S" & row).Value
    Sheets("Sub_Bought_Out").Range("F" & rowe).Value = "=COUNTIF(Psv_Values_Sub!A2:A" & counts & ",L" & rowe & ")"
    Sheets("Sub_Bought_Out").Range("G" & rowe).Value = 0
    Sheets("Sub_Bought_Out").Range("H" & rowe).Formula = "=Sum(F" & rowe & ":G" & rowe & ")"
    Sheets("Sub_Bought_Out").Range("L" & rowe).Value = Sheets("Psv_Values_Sub").Range("A" & row).Value
    Sheets("Sub_Bought_Out").Range("A" & rowe).Value = rowe - 4
    
    rowe = rowe + 1
    End If
    
    row = row + 1
    Loop
    
    ' Next line of codes for simple formatting and print setup
    
    With Sheets("Sub_Bought_Out").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Prepared by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Checked by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Verified by"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "Approved by"
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Bought_Out").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = Application.UserName
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    
    End With
    
    With Sheets("Sub_Bought_Out").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    rowe = rowe + 1
    With Sheets("Sub_Bought_Out").Range("A" & rowe, "C" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Designer)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("D" & rowe, "F" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Group Head)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("G" & rowe, "I" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(Design Incharge)"
        .Borders.Weight = xlMedium
    End With
    
    With Sheets("Sub_Bought_Out").Range("J" & rowe, "K" & rowe)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Value = "(D.P.M)"
        .Borders.Weight = xlMedium
    End With
    
    Sheets("Sub_Bought_Out").Range("A1", "K" & rowe).BorderAround _
    Weight:=xlMedium
    Worksheets("Sub_Bought_Out").PageSetup.PrintArea = Worksheets("Sub_Bought_Out").Range("A1", "K" & rowe).Address
    
   



End Sub


Private Sub CommandButton13_Click()

    Call CommandButton8_Click
    Call CommandButton9_Click
    Call CommandButton10_Click
    Call CommandButton11_Click
    Call CommandButton12_Click
    Call CommandButton14_Click
    
    
End Sub


Private Sub CommandButton14_Click()

'This code is responsible for copying the sub assembly BOM generated to a separate excel workbook

Dim loc As String
loc = ThisWorkbook.Path        'Location on sub assembly BOM's can be changed here
Dim file_name As String
file_name = Sheets("Assembly_numbers").Range("B" & coun).Value              'Filename of Workbook can be changed here
Dim file_path As String
file_path = loc & "\Sub\" & file_name & ".xlsx"
Dim wbNew As Workbook
ThisWorkbook.Sheets(Array("Sub_Maintenance", "Sub_Bought_Out", "Sub_Fabricated", "Sub_Machined_Component")).Copy        'Copies all sheets. If only particular sheets required, can be updated from here
Set wbNew = ActiveWorkbook
With wbNew
    With .Worksheets(1).UsedRange
        .Value = .Value
    End With
    With .Worksheets(2).UsedRange
        .Value = .Value
    End With
    With .Worksheets(3).UsedRange
        .Value = .Value
    End With
    With .Worksheets(4).UsedRange
        .Value = .Value
    End With
    .SaveAs file_path
     .Close SaveChanges:=False
End With

'After the new workbook has been generated, values previously input in the BOM sheets has to be deleted,
'so that values from next sub assembly can be placed.

Sheets("Sub_Maintenance").Rows("5:1000").EntireRow.delete
Sheets("Sub_Bought_Out").Rows("5:1000").EntireRow.delete
Sheets("Sub_Fabricated").Rows("6:1000").EntireRow.delete
Sheets("Sub_Machined_Component").Rows("5:1000").EntireRow.delete
Application.DisplayAlerts = False
Sheets("Psv_Values_Sub").delete
Application.DisplayAlerts = True


End Sub
Private Sub CommandButton18_Click()

Dim locm As String
locm = ThisWorkbook.Path         'Location on sub assembly BOM's can be changed here
Dim file_name As String
file_namem = Sheets("Psv_Values").Range("A2").Value              'Filename of Workbook can be changed here
Dim file_path As String
file_path = locm & "\" & file_namem & ".xlsx"
Dim wbNew As Workbook
ThisWorkbook.Sheets(Array("Master_Maintenance", "Master_Bought_Out", "Master_Fabricated", "Master_Machined_Component")).Copy        'Copies all sheets. If only particular sheets required, can be updated from here
Set wbNew = ActiveWorkbook
With wbNew
    With .Worksheets(1).UsedRange
        .Value = .Value
    End With
    With .Worksheets(2).UsedRange
        .Value = .Value
    End With
    With .Worksheets(3).UsedRange
        .Value = .Value
    End With
    With .Worksheets(4).UsedRange
        .Value = .Value
    End With
    .SaveAs file_path
     .Close SaveChanges:=False
End With

'After the new workbook has been generated, values previously input in the BOM sheets has to be deleted,
'so that values from next sub assembly can be placed.

Sheets("Master_Maintenance").Rows("5:1000").EntireRow.delete
Sheets("Master_Bought_Out").Rows("5:1000").EntireRow.delete
Sheets("Master_Fabricated").Rows("6:1000").EntireRow.delete
Sheets("Master_Machined_Component").Rows("5:1000").EntireRow.delete
Application.DisplayAlerts = False
Application.DisplayAlerts = True



End Sub
Private Sub CommandButton15_Click()
 
coun = 8
Do Until Sheets("Assembly_numbers").Range("B" & coun).Value = ""
    sub_name = Sheets("Assembly_numbers").Range("B" & coun).Value
    Call CommandButton8_Click
    Call CommandButton9_Click
    Call CommandButton10_Click
    Call CommandButton11_Click
    Call CommandButton12_Click
    Call CommandButton14_Click
    coun = coun + 1
Loop
    
    
 
  
   MsgBox "Entire SubAssembly Operation complete"
  
End Sub

Private Sub CommandButton16_Click()
  'Call CommandButton1_Click
  
count = 1
Do Until Sheets("Psv_Values").Range("A" & count).Value = ""
count = count + 1
Loop
coun = 8
Do Until Sheets("Assembly_numbers").Range("B" & coun).Value = ""
    sub_name = Sheets("Assembly_numbers").Range("B" & coun).Value
    Call CommandButton8_Click
    Call CommandButton9_Click
'    Call CommandButton10_Click
'    Call CommandButton11_Click
'    Call CommandButton12_Click
'    Call CommandButton14_Click
'above comment Temporarily added on 3rd August 2021
    coun = coun + 1
Loop
    
    
 
    Call CommandButton6_Click
    Call CommandButton2_Click
    Call CommandButton3_Click
    Call CommandButton4_Click
    Call CommandButton5_Click
    Call CommandButton18_Click
   MsgBox "Entire Operation complete GODSpeed"
   
End Sub

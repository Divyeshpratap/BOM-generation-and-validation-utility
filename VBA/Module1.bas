Attribute VB_Name = "Module1"

Sub delete()

Dim count As String
Dim as_num  As String
count = 8
Dim counter As String
counter = 2
Dim level As String
level = 1
Dim counter_two As String

Do Until Sheets("Assembly_numbers").Range("B" & count).Value = ""
    as_num = Sheets("Assembly_numbers").Range("B" & count).Value
    Sheets("Assembly_numbers").Range("J" & count).Value = as_num
        Do Until counter = 96000
            If Sheets("Psv_Values").Range("A" & counter).Value = as_num Then
            level = Sheets("Psv_Values").Range("B" & counter).Value
            Sheets("Psv_Values").Range("A" & counter).Value = "Match found"
            Sheets("Psv_Values").Range("C" & counter).Value = ""
            counter_two = counter + 1
            Do Until level >= Sheets("Psv_Values").Range("B" & counter_two).Value
                Sheets("Psv_Values").Range("A" & counter_two).Value = "Sub assembly deleted"
                Sheets("Psv_Values").Range("C" & counter_two).Value = ""
                counter_two = counter_two + 1
            Loop
            End If
            counter = counter + 1
            level = 1
        Loop
  
        
    counter = 2
    count = count + 1

Loop

Sheets("Assembly_numbers").Range("K" & count).Value = count - 8


End Sub

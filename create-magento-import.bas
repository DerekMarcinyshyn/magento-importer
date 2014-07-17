Option Explicit
 
Sub CreateMagentoImport()
'
' @author   Derek Marcinyshyn <derek@marcinyshyn.com>
' @date     July 9, 2014
' @updated  July 15, 2014
' @version  1.0.4
 
' To enable
' Developer Tab
' AddIns => Check Create Magento Import
' AddIns => Check Menu
 
' Check to see if references were added
' Developer Tab
' Tools => References
' Microsoft Visual Basic for Applications Extensibility 5.3
' Microsoft Forms 2.0 Object Library
' 
' 

Call AddReference
Call CheckForErrors
 
End Sub

Private Function CheckForErrors()

    Dim WorkSheet1 As Worksheet
    Dim WorkSheetMagentoImport As Worksheet

    On Error Resume Next
    Set WorkSheet1 = Sheets("Sheet1")
    
    If WorkSheet1 Is Nothing Then
        MsgBox "Sheet1 must exist with the data! Rename your data worksheet.", vbCritical, "Can find Sheet1"
        Set WorkSheet1 = Nothing
        On Error GoTo 0
    Else
        On Error Resume Next
        Set WorkSheetMagentoImport = Sheets("Magento Import")
        
        If WorkSheetMagentoImport Is Nothing Then
            Call CreateInterface
        Else
            MsgBox "You have already run it once. Please delete the Worksheet called 'Magento Import'!", vbCritical, "Magento Import already exists"
            Set WorkSheetMagentoImport = Nothing
            On Error GoTo 0
        End If
    End If

End Function


Private Function AddReference()
 
 Dim i As Long
 Dim guidArr() As String
 ReDim guidArr(1 To 2, 1 To 3)
 
 
 ' reference details
 ' Name: MSForms
 ' Description: Microsoft Forms 2.0 Object Library
 ' GUID: {0D452EE1-E08F-101A-852E-02608C4D0BB4}
 ' Major: 2
 ' Minor: 0
 ' FullPath: C:\Windows\SysWOW64\FM20.DLL
 guidArr(1, 1) = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
 guidArr(1, 2) = "2"
 guidArr(1, 3) = "0"

 ' reference details
 ' Name: VBIDE
 ' Description: Microsoft Visual Basic for Applications Extensibility 5.3
 ' GUID: {0002E157-0000-0000-C000-000000000046}
 ' Major: 5
 ' Minor: 3
 ' FullPath: C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
 guidArr(2, 1) = "{0002E157-0000-0000-C000-000000000046}"
 guidArr(2, 2) = "5"
 guidArr(2, 3) = "3"
 
 On Error Resume Next
 
 For i = 1 To 2
     ThisWorkbook.VBProject.References.AddFromGuid GUID:=guidArr(i, 1), Major:=guidArr(i, 2), Minor:=guidArr(i, 3)
 Next i
 
End Function

 
Private Function CreateInterface()
 
    Dim MyForm As Object
    Dim MyCommandButton As MSForms.CommandButton
    Dim ModeLabel As MSForms.Label
    Dim ModeStandardButton As MSForms.OptionButton
    Dim ModeBoardsButton As MSForms.OptionButton
    Dim ModeBoardsSizesButton As MSForms.OptionButton
    Dim ImageLabel As MSForms.Label
    Dim ImageTextbox As MSForms.TextBox
    Dim CategoriesLabel As MSForms.Label
    Dim CategoriesTextbox As MSForms.TextBox
    Dim ManufacturerLabel As MSForms.Label
    Dim ManufacturerTextBox As MSForms.TextBox
    Dim StartDateLabel As MSForms.Label
    Dim StartDateTextbox As MSForms.TextBox
    Dim EndDateLabel As MSForms.Label
    Dim EndDateTextbox As MSForms.TextBox
    
    
    ' Stop screen from flashing
    Application.VBE.MainWindow.Visible = False
    
    Set MyForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    
    With MyForm
        .Properties("Caption") = "Chaindrive to Magento"
        .Properties("Width") = 420
        .Properties("Height") = 400
    End With
    
    Set ModeLabel = MyForm.Designer.Controls.Add("Forms.label.1")
    With ModeLabel
        .Top = 20
        .Left = 20
        .Width = 300
        .Caption = "Select one of the import modes"
        .Font.Name = "Tahoma"
        .Font.Size = 12
    End With
    
    Set ModeStandardButton = MyForm.Designer.Controls.Add("Forms.optionbutton.1")
    With ModeStandardButton
        .Name = "ModeStandard"
        .Value = True
        .Top = 40
        .Left = 20
        .Caption = "Standard"
        .Font.Name = "Tahoma"
        .Font.Size = 10
    End With
    
    Set ModeBoardsButton = MyForm.Designer.Controls.Add("Forms.optionbutton.1")
    With ModeBoardsButton
        .Name = "ModeBoards"
        .Top = 40
        .Left = 100
        .Caption = "Boards"
        .Font.Name = "Tahoma"
        .Font.Size = 10
    End With
    
    Set ModeBoardsSizesButton = MyForm.Designer.Controls.Add("Forms.optionbutton.1")
    With ModeBoardsSizesButton
        .Name = "ModeBoardsSizes"
        .Top = 40
        .Left = 160
        .Width = 300
        .Caption = "Boards with individual image sizes"
        .Font.Name = "Tahoma"
        .Font.Size = 10
    End With
    
    Set ImageLabel = MyForm.Designer.Controls.Add("Forms.label.1")
    With ImageLabel
        .Top = 90
        .Left = 20
        .Width = 350
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .Caption = "Image Path eg. W15/BURTON/SOFTGOODS/"
    End With
    
    Set ImageTextbox = MyForm.Designer.Controls.Add("Forms.textbox.1")
    With ImageTextbox
        .Name = "ImagePath"
        .Top = 110
        .Left = 20
        .Width = 380
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .BackColor = RGB(220, 220, 220)
        .Height = 20
    End With
    
    Set CategoriesLabel = MyForm.Designer.Controls.Add("Forms.label.1")
    With CategoriesLabel
        .Top = 140
        .Left = 20
        .Width = 350
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .Caption = "Categories Path eg. Mens/Hardgoods/Snowboards"
    End With
    
    Set CategoriesTextbox = MyForm.Designer.Controls.Add("Forms.textbox.1")
    With CategoriesTextbox
        .Name = "Categories"
        .Top = 160
        .Left = 20
        .Width = 380
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .BackColor = RGB(220, 220, 220)
        .Height = 20
    End With
    
    Set ManufacturerLabel = MyForm.Designer.Controls.Add("Forms.label.1")
    With ManufacturerLabel
        .Top = 190
        .Left = 20
        .Width = 350
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .Caption = "Manufacturer as it appears in Magento"
    End With
    
    Set ManufacturerTextBox = MyForm.Designer.Controls.Add("Forms.textbox.1")
    With ManufacturerTextBox
        .Name = "Manufacturer"
        .Top = 210
        .Left = 20
        .Width = 380
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .BackColor = RGB(220, 220, 220)
        .Height = 20
    End With
    
    Set StartDateLabel = MyForm.Designer.Controls.Add("Forms.label.1")
    With StartDateLabel
        .Top = 240
        .Left = 20
        .Width = 350
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .Caption = "New Arrivals Start Date eg. 2014-07-01"
    End With
    
    Set StartDateTextbox = MyForm.Designer.Controls.Add("Forms.textbox.1")
    With StartDateTextbox
        .Name = "StartDate"
        .Top = 260
        .Left = 20
        .Width = 380
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .BackColor = RGB(220, 220, 220)
        .Height = 20
    End With
    
    Set EndDateLabel = MyForm.Designer.Controls.Add("Forms.label.1")
    With EndDateLabel
        .Top = 290
        .Left = 20
        .Width = 350
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .Caption = "New Arrivals End Date eg. 2014-09-01"
    End With
    
    Set EndDateTextbox = MyForm.Designer.Controls.Add("Forms.textbox.1")
    With EndDateTextbox
        .Name = "EndDate"
        .Top = 310
        .Left = 20
        .Width = 380
        .Font.Size = 11
        .Font.Name = "Tahoma"
        .BackColor = RGB(220, 220, 220)
        .Height = 20
    End With
    
    Set MyCommandButton = MyForm.Designer.Controls.Add("Forms.commandbutton.1")
    With MyCommandButton
        .Name = "cmd_1"
        .Caption = "Unleash"
        .Top = 350
        .Left = 20
        .Accelerator = "M"
        .Font.Size = 12
        .Font.Name = "Tahoma"
        .BackColor = RGB(160, 0, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Inject some code for the button
    MyForm.CodeModule.InsertLines 1, "Private Sub cmd_1_Click()"
    MyForm.CodeModule.InsertLines 2, "   Call Unleash(ModeStandard.Value, ModeBoards.Value, ModeBoardsSizes.Value, ImagePath.Value, Manufacturer.Value, Categories.Value, StartDate.Value, EndDate.Value)"
    MyForm.CodeModule.InsertLines 3, "End Sub"
    
    VBA.UserForms.Add(MyForm.Name).Show
 
End Function
 
Public Function Unleash(ModeStandard, ModeBoards, ModeBoardsSizes, ImagePath, Manufacturer, Categories, StartDate, EndDate)
 
    ' Check if fields are filled in
    If ImagePath <> "" And Manufacturer <> "" And Categories <> "" And StartDate <> "" And EndDate <> "" Then
        ' Duplicate sheets
        Call Duplicate
        
        ' Check for mode
        If ModeStandard Then
            Call Standard(ImagePath, Manufacturer, Categories, StartDate, EndDate)
        End If
        
        If ModeBoards Then
            Call Boards(ImagePath, Manufacturer, Categories, StartDate, EndDate)
        End If
        
        If ModeBoardsSizes Then
            Call BoardsSizes(ImagePath, Manufacturer, Categories, StartDate, EndDate)
        End If
            
    Else
        MsgBox "You need to fill in all fields!", vbCritical, "You trying to blow this up?"
    End If
    
End Function
 
Private Function Standard(ImagePath, Manufacturer, Categories, StartDate, EndDate)
    
    Dim numberOfRowsImage As Long
    Dim rowNumberImage As Long
    Dim sourceSku As Long
    
    numberOfRowsImage = Sheets("Magento Import").Cells(Rows.Count, 1).End(xlUp).Row
 
    For rowNumberImage = 2 To numberOfRowsImage
        Cells(rowNumberImage, 19).Value = ImagePath & Cells(rowNumberImage, 4) & ".jpg"
        Cells(rowNumberImage, 20).Value = ImagePath & Cells(rowNumberImage, 4) & ".jpg"
        Cells(rowNumberImage, 21).Value = ImagePath & Cells(rowNumberImage, 4) & ".jpg"
        Cells(rowNumberImage, 3).Value = Categories
        Cells(rowNumberImage, 26).NumberFormat = "YYYY/MM/DD"
        Cells(rowNumberImage, 26).Value = StartDate
        Cells(rowNumberImage, 27).NumberFormat = "YYYY/MM/DD"
        Cells(rowNumberImage, 27).Value = EndDate
        Cells(rowNumberImage, 10).Value = Manufacturer
    Next rowNumberImage
    
    ' Create new row if source_sku changes
    For sourceSku = Cells(Cells.Rows.Count, "D").End(xlUp).Row To 3 Step -1
        If Cells(sourceSku, "D") <> Cells(sourceSku - 1, "D") Then Rows(sourceSku).EntireRow.Insert
    Next sourceSku
        
    ' Get the next empty row and copy data from one line above
    Dim numberOfRows As Long
    Dim rowNumber As Long
    
    numberOfRows = Sheets("Magento Import").Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    For rowNumber = 2 To numberOfRows
        
        ' A attribute_set
        If Cells(rowNumber, 1).Value = "" Then
            Cells(rowNumber, 1).Value = Cells(rowNumber - 1, 1).Value
        
        ' Wrap it around just checking the first cell of each row
        
           ' B type
            If Cells(rowNumber, 2).Value = "" Then
                Cells(rowNumber, 2).Value = "configurable"
            End If
            
            ' C categories
            If Cells(rowNumber, 3).Value = "" Then
                Cells(rowNumber, 3).Value = Categories
            End If
            
            ' D source_sku -- leave blank?
            
            ' E name
            If Cells(rowNumber, 5).Value = "" Then
                Cells(rowNumber, 5).Value = Left(Cells(rowNumber - 1, 5).Value, InStrRev(Cells(rowNumber - 1, 5).Value, " ") - 1)
            End If
            
            ' F sku
            If Cells(rowNumber, 6).Value = "" Then
                Cells(rowNumber, 6).Value = Cells(rowNumber - 1, 4).Value
            End If
            
            ' G boot_size -- should be empty
            
            ' H price
            If Cells(rowNumber, 8).Value = "" Then
                Cells(rowNumber, 8).Value = Cells(rowNumber - 1, 8).Value
            End If
            
            ' I qty -- blank
                        
            ' J manufacturer
            If Cells(rowNumber, 10).Value = "" Then
                Cells(rowNumber, 10).Value = Manufacturer
            End If
            
            ' K weight
            If Cells(rowNumber, 11).Value = "" Then
                Cells(rowNumber, 11).Value = Cells(rowNumber - 1, 11).Value
            End If
            
            ' L length
            If Cells(rowNumber, 12).Value = "" Then
                Cells(rowNumber, 12).Value = Cells(rowNumber - 1, 12).Value
            End If
            
            ' M width
            If Cells(rowNumber, 13).Value = "" Then
                Cells(rowNumber, 13).Value = Cells(rowNumber - 1, 13).Value
            End If
            
            ' N height
            If Cells(rowNumber, 14).Value = "" Then
                Cells(rowNumber, 14).Value = Cells(rowNumber - 1, 14).Value
            End If
            
            ' O simple_skus
            If Cells(rowNumber, 15).Value = "" Then
                ' check all rows above to see if they numerical if so then must be simple sku
                Dim simpleSku As String
                                
                Dim skuCounter As Integer
                skuCounter = 0
                
                Dim currentRowNumber As Integer
                currentRowNumber = rowNumber
                
                ' Loop through and append string
                While (IsNumeric(Cells(currentRowNumber - 1, 6)))
                    If (skuCounter = 0) Then simpleSku = Cells(currentRowNumber - 1, 6).Value
                    If (skuCounter > 0) Then simpleSku = simpleSku & ", " & Cells(currentRowNumber - 1, 6).Value
                    
                    skuCounter = skuCounter + 1
                    currentRowNumber = currentRowNumber - 1
                Wend
                
                Cells(rowNumber, 15).Value = CStr(simpleSku)
            End If
            
            ' P configurable_attributes
            If Cells(rowNumber, 16).Value = "" Then
                Cells(rowNumber, 16).Value = Cells(1, 7).Value
            End If
            
            ' Q visibility
            If Cells(rowNumber, 17).Value = "" Then
                Cells(rowNumber, 17).Value = "Catalog, Search"
            End If
            
            ' R status
            If Cells(rowNumber, 18).Value = "" Then
                Cells(rowNumber, 18).Value = "Enabled"
            End If
            
            ' S image
            If Cells(rowNumber, 19).Value = "" Then
                Cells(rowNumber, 19).Value = Cells(rowNumber - 1, 19).Value
            End If
            
            ' T small_image
            If Cells(rowNumber, 20).Value = "" Then
                Cells(rowNumber, 20).Value = Cells(rowNumber - 1, 20).Value
            End If
            
            ' U thumbnail
            If Cells(rowNumber, 21).Value = "" Then
                Cells(rowNumber, 21).Value = Cells(rowNumber - 1, 21).Value
            End If
            
            ' V media_gallery -- skip
            
            ' W tax_class_id
            If Cells(rowNumber, 23).Value = "" Then
                Cells(rowNumber, 23).Value = Cells(rowNumber - 1, 23).Value
            End If
            
            ' X is_in_stock
            If Cells(rowNumber, 24).Value = "" Then
                Cells(rowNumber, 24).Value = Cells(rowNumber - 1, 24).Value
            End If
            
            ' Y season
            If Cells(rowNumber, 25).Value = "" Then
                Cells(rowNumber, 25).Value = Cells(rowNumber - 1, 25).Value
            End If
            
            ' Z news_from_date
            If Cells(rowNumber, 26).Value = "" Then
                Cells(rowNumber, 26).NumberFormat = "YYYY/MM/DD"
                Cells(rowNumber, 26).Value = StartDate
            End If
            
            ' AA news_to_date
            If Cells(rowNumber, 27).Value = "" Then
                Cells(rowNumber, 27).NumberFormat = "YYYY/MM/DD"
                Cells(rowNumber, 27).Value = EndDate
            End If
            
            ' AB short_description
            If Cells(rowNumber, 28).Value = "" Then
                Cells(rowNumber, 28).Value = Cells(rowNumber - 1, 28).Value
            End If
            
            ' AC description
            If Cells(rowNumber, 29).Value = "" Then
                Cells(rowNumber, 29).Value = Cells(rowNumber - 1, 29).Value
            End If
  
        End If
        
    Next rowNumber
    
    Call ResizeColumns
End Function
 
Private Function Boards(ImagePath, Manufacturer, Categories, StartDate, EndDate)
 
    Dim numberOfRowsImage As Long
    Dim rowNumberImage As Long
    Dim sourceSku As Long
    
    numberOfRowsImage = Sheets("Magento Import").Cells(Rows.Count, 1).End(xlUp).Row
 
    For rowNumberImage = 2 To numberOfRowsImage
        Cells(rowNumberImage, 22).Value = ImagePath & Cells(rowNumberImage, 4) & ".jpg"
        Cells(rowNumberImage, 23).Value = ImagePath & Cells(rowNumberImage, 4) & ".jpg"
        Cells(rowNumberImage, 24).Value = ImagePath & Cells(rowNumberImage, 4) & ".jpg"
        Cells(rowNumberImage, 3).Value = Categories
        Cells(rowNumberImage, 29).NumberFormat = "YYYY/MM/DD"
        Cells(rowNumberImage, 29).Value = StartDate
        Cells(rowNumberImage, 30).NumberFormat = "YYYY/MM/DD"
        Cells(rowNumberImage, 30).Value = EndDate
        Cells(rowNumberImage, 13).Value = Manufacturer
    Next rowNumberImage
 
    ' Create new row if source_sku changes
    For sourceSku = Cells(Cells.Rows.Count, "D").End(xlUp).Row To 3 Step -1
        If Cells(sourceSku, "D") <> Cells(sourceSku - 1, "D") Then Rows(sourceSku).EntireRow.Insert
    Next sourceSku
 
' Get the next empty row and copy data from one line above
    Dim numberOfRows As Long
    Dim rowNumber As Long
    
    numberOfRows = Sheets("Magento Import").Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    For rowNumber = 2 To numberOfRows
        
        ' A attribute_set
        If Cells(rowNumber, 1).Value = "" Then
            Cells(rowNumber, 1).Value = Cells(rowNumber - 1, 1).Value
        
        ' Wrap it around just checking the first cell of each row
        
           ' B type
            If Cells(rowNumber, 2).Value = "" Then
                Cells(rowNumber, 2).Value = "configurable"
            End If
            
            ' C categories
            If Cells(rowNumber, 3).Value = "" Then
                Cells(rowNumber, 3).Value = Categories
            End If
            
            ' D source_sku -- leave blank?
            
            ' E camber
            If Cells(rowNumber, 5).Value = "" Then
                Cells(rowNumber, 5).Value = Cells(rowNumber - 1, 5).Value
            End If
            
            ' F terrain
            If Cells(rowNumber, 6).Value = "" Then
                Cells(rowNumber, 6).Value = Cells(rowNumber - 1, 6).Value
            End If
            
            ' G board_width
            If Cells(rowNumber, 7).Value = "" Then
                Cells(rowNumber, 7).Value = Cells(rowNumber - 1, 7).Value
            End If
            
            ' H name -- remove last word
            If Cells(rowNumber, 8).Value = "" Then
                Cells(rowNumber, 8).Value = Left(Cells(rowNumber - 1, 8).Value, InStrRev(Cells(rowNumber - 1, 8).Value, " ") - 1)
            End If
            
            ' I sku -- copy source_sku
            If Cells(rowNumber, 9).Value = "" Then
                Cells(rowNumber, 9).Value = Cells(rowNumber - 1, 4).Value
            End If
            
            ' J snowboard_size -- leave blank?
            If Cells(rowNumber, 10).Value = "" Then
                Cells(rowNumber, 10).Value = ""
            End If
            
            ' K price
            If Cells(rowNumber, 11).Value = "" Then
                Cells(rowNumber, 11).Value = Cells(rowNumber - 1, 11).Value
            End If
            
            ' L qty -- blank
            
            ' M manufacturer
            If Cells(rowNumber, 13).Value = "" Then
                Cells(rowNumber, 13).Value = Manufacturer
            End If
            
            ' N weight
            If Cells(rowNumber, 14).Value = "" Then
                Cells(rowNumber, 14).Value = Cells(rowNumber - 1, 14).Value
            End If
            
            ' O length
            If Cells(rowNumber, 15).Value = "" Then
                Cells(rowNumber, 15).Value = Cells(rowNumber - 1, 15).Value
            End If
            
            ' P width
            If Cells(rowNumber, 16).Value = "" Then
                Cells(rowNumber, 16).Value = Cells(rowNumber - 1, 16).Value
            End If
            
            ' Q height
            If Cells(rowNumber, 17).Value = "" Then
                Cells(rowNumber, 17).Value = Cells(rowNumber - 1, 17).Value
            End If
            
            ' R simple_skus
            If Cells(rowNumber, 18).Value = "" Then
                ' check all rows above to see if they numerical if so then must be simple sku
                Dim simpleSku As String
                                
                Dim skuCounter As Integer
                skuCounter = 0
                
                Dim currentRowNumber As Integer
                currentRowNumber = rowNumber
                
                ' Loop through and append string
                While (IsNumeric(Cells(currentRowNumber - 1, 9)))
                    If (skuCounter = 0) Then simpleSku = Cells(currentRowNumber - 1, 9).Value
                    If (skuCounter > 0) Then simpleSku = simpleSku & ", " & Cells(currentRowNumber - 1, 9).Value
                    
                    skuCounter = skuCounter + 1
                    currentRowNumber = currentRowNumber - 1
                Wend
                
                Cells(rowNumber, 18).Value = CStr(simpleSku)
            End If
            
            ' S configurable_attributes
            If Cells(rowNumber, 19).Value = "" Then
                Cells(rowNumber, 19).Value = Cells(1, 10).Value
            End If
            
            ' T visibility
            If Cells(rowNumber, 20).Value = "" Then
                Cells(rowNumber, 20).Value = "Catalog, Search"
            End If
            
            ' U status
            If Cells(rowNumber, 21).Value = "" Then
                Cells(rowNumber, 21).Value = "Enabled"
            End If
            
            ' V image
            If Cells(rowNumber, 22).Value = "" Then
                Cells(rowNumber, 22).Value = ImagePath & Cells(rowNumber - 1, 4).Value & ".jpg"
            End If
            
            ' W small_image
            If Cells(rowNumber, 23).Value = "" Then
                Cells(rowNumber, 23).Value = ImagePath & Cells(rowNumber - 1, 4).Value & ".jpg"
            End If
            
            ' X thumbnail
            If Cells(rowNumber, 24).Value = "" Then
                Cells(rowNumber, 24).Value = ImagePath & Cells(rowNumber - 1, 4).Value & ".jpg"
            End If
            
            ' Y media_gallery -- skip?
            
            ' Z tax_class_id
            If Cells(rowNumber, 26).Value = "" Then
                Cells(rowNumber, 26).Value = Cells(rowNumber - 1, 26).Value
            End If
            
            ' AA is_in_stock
            If Cells(rowNumber, 27).Value = "" Then
                Cells(rowNumber, 27).Value = Cells(rowNumber - 1, 27).Value
            End If
            
            ' AB season
            If Cells(rowNumber, 28).Value = "" Then
                Cells(rowNumber, 28).Value = Cells(rowNumber - 1, 28).Value
            End If
            
            ' AC news_from_date
            If Cells(rowNumber, 29).Value = "" Then
                Cells(rowNumber, 29).NumberFormat = "YYYY/MM/DD"
                Cells(rowNumber, 29).Value = StartDate
            End If
            
            ' AD news_to_date
            If Cells(rowNumber, 30).Value = "" Then
                Cells(rowNumber, 30).NumberFormat = "YYYY/MM/DD"
                Cells(rowNumber, 30).Value = EndDate
            End If
        End If
        
    Next rowNumber
 
    Call ResizeColumns
 
End Function
 
Private Function BoardsSizes(ImagePath, Manufacturer, Categories, StartDate, EndDate)
    
    Dim numberOfRowsImage As Long
    Dim rowNumberImage As Long
    Dim sourceSku As Long
    
    numberOfRowsImage = Sheets("Magento Import").Cells(Rows.Count, 1).End(xlUp).Row
    
    For rowNumberImage = 2 To numberOfRowsImage
        Cells(rowNumberImage, 22).Value = ImagePath & Cells(rowNumberImage, 4) & "_" & Cells(rowNumberImage, 10) & ".jpg"
        Cells(rowNumberImage, 23).Value = ImagePath & Cells(rowNumberImage, 4) & "_" & Cells(rowNumberImage, 10) & ".jpg"
        Cells(rowNumberImage, 24).Value = ImagePath & Cells(rowNumberImage, 4) & "_" & Cells(rowNumberImage, 10) & ".jpg"
        Cells(rowNumberImage, 3).Value = Categories
        Cells(rowNumberImage, 13).Value = Manufacturer
        Cells(rowNumberImage, 29).NumberFormat = "YYYY/MM/DD"
        Cells(rowNumberImage, 29).Value = StartDate
        Cells(rowNumberImage, 30).NumberFormat = "YYYY/MM/DD"
        Cells(rowNumberImage, 30).Value = EndDate
    Next rowNumberImage
 
    ' Create new row if source_sku changes
    For sourceSku = Cells(Cells.Rows.Count, "D").End(xlUp).Row To 3 Step -1
        If Cells(sourceSku, "D") <> Cells(sourceSku - 1, "D") Then Rows(sourceSku).EntireRow.Insert
    Next sourceSku
    
    
    ' Get the next empty row and copy data from one line above
    Dim numberOfRows As Long
    Dim rowNumber As Long
    
    numberOfRows = Sheets("Magento Import").Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    For rowNumber = 2 To numberOfRows
        
        ' A attribute_set
        If Cells(rowNumber, 1).Value = "" Then
            Cells(rowNumber, 1).Value = Cells(rowNumber - 1, 1).Value
        
        ' Wrap it around just checking the first cell of each row
        
           ' B type
            If Cells(rowNumber, 2).Value = "" Then
                Cells(rowNumber, 2).Value = "configurable"
            End If
            
            ' C categories
            If Cells(rowNumber, 3).Value = "" Then
                Cells(rowNumber, 3).Value = Categories
            End If
            
            ' D source_sku -- leave blank?
            
            ' E camber
            If Cells(rowNumber, 5).Value = "" Then
                Cells(rowNumber, 5).Value = Cells(rowNumber - 1, 5).Value
            End If
            
            ' F terrain
            If Cells(rowNumber, 6).Value = "" Then
                Cells(rowNumber, 6).Value = Cells(rowNumber - 1, 6).Value
            End If
            
            ' G board_width
            If Cells(rowNumber, 7).Value = "" Then
                Cells(rowNumber, 7).Value = Cells(rowNumber - 1, 7).Value
            End If
            
            ' H name -- remove last word
            If Cells(rowNumber, 8).Value = "" Then
                Cells(rowNumber, 8).Value = Left(Cells(rowNumber - 1, 8).Value, InStrRev(Cells(rowNumber - 1, 8).Value, " ") - 1)
            End If
            
            ' I sku -- copy source_sku
            If Cells(rowNumber, 9).Value = "" Then
                Cells(rowNumber, 9).Value = Cells(rowNumber - 1, 4).Value
            End If
            
            ' J snowboard_size -- leave blank?
            If Cells(rowNumber, 10).Value = "" Then
                Cells(rowNumber, 10).Value = ""
            End If
            
            ' K price
            If Cells(rowNumber, 11).Value = "" Then
                Cells(rowNumber, 11).Value = Cells(rowNumber - 1, 11).Value
            End If
            
            ' L qty -- blank
            
            ' M manufacturer
            If Cells(rowNumber, 13).Value = "" Then
                Cells(rowNumber, 13).Value = Manufacturer
            End If
            
            ' N weight
            If Cells(rowNumber, 14).Value = "" Then
                Cells(rowNumber, 14).Value = Cells(rowNumber - 1, 14).Value
            End If
            
            ' O length
            If Cells(rowNumber, 15).Value = "" Then
                Cells(rowNumber, 15).Value = Cells(rowNumber - 1, 15).Value
            End If
            
            ' P width
            If Cells(rowNumber, 16).Value = "" Then
                Cells(rowNumber, 16).Value = Cells(rowNumber - 1, 16).Value
            End If
            
            ' Q height
            If Cells(rowNumber, 17).Value = "" Then
                Cells(rowNumber, 17).Value = Cells(rowNumber - 1, 17).Value
            End If
            
            ' R simple_skus
            If Cells(rowNumber, 18).Value = "" Then
                ' check all rows above to see if they numerical if so then must be simple sku
                Dim simpleSku As String
                                
                Dim skuCounter As Integer
                skuCounter = 0
                
                Dim currentRowNumber As Integer
                currentRowNumber = rowNumber
                
                ' Loop through and append string
                While (IsNumeric(Cells(currentRowNumber - 1, 9)))
                    If (skuCounter = 0) Then simpleSku = Cells(currentRowNumber - 1, 9).Value
                    If (skuCounter > 0) Then simpleSku = simpleSku & ", " & Cells(currentRowNumber - 1, 9).Value
                    
                    skuCounter = skuCounter + 1
                    currentRowNumber = currentRowNumber - 1
                Wend
                
                Cells(rowNumber, 18).Value = CStr(simpleSku)
            End If
            
            ' S configurable_attributes
            If Cells(rowNumber, 19).Value = "" Then
                Cells(rowNumber, 19).Value = Cells(1, 10).Value
            End If
            
            ' T visibility
            If Cells(rowNumber, 20).Value = "" Then
                Cells(rowNumber, 20).Value = "Catalog, Search"
            End If
            
            ' U status
            If Cells(rowNumber, 21).Value = "" Then
                Cells(rowNumber, 21).Value = "Enabled"
            End If
            
            ' V image
            If Cells(rowNumber, 22).Value = "" Then
                Cells(rowNumber, 22).Value = ImagePath & Cells(rowNumber - 1, 4).Value & ".jpg"
            End If
            
            ' W small_image
            If Cells(rowNumber, 23).Value = "" Then
                Cells(rowNumber, 23).Value = ImagePath & Cells(rowNumber - 1, 4).Value & ".jpg"
            End If
            
            ' X thumbnail
            If Cells(rowNumber, 24).Value = "" Then
                Cells(rowNumber, 24).Value = ImagePath & Cells(rowNumber - 1, 4).Value & ".jpg"
            End If
            
            ' Y media_gallery -- skip?
            
            ' Z tax_class_id
            If Cells(rowNumber, 26).Value = "" Then
                Cells(rowNumber, 26).Value = Cells(rowNumber - 1, 26).Value
            End If
            
            ' AA is_in_stock
            If Cells(rowNumber, 27).Value = "" Then
                Cells(rowNumber, 27).Value = Cells(rowNumber - 1, 27).Value
            End If
            
            ' AB season
            If Cells(rowNumber, 28).Value = "" Then
                Cells(rowNumber, 28).Value = Cells(rowNumber - 1, 28).Value
            End If
            
            ' AC news_from_date
            If Cells(rowNumber, 29).Value = "" Then
                Cells(rowNumber, 29).NumberFormat = "YYYY/MM/DD"
                Cells(rowNumber, 29).Value = StartDate
            End If
            
            ' AD news_to_date
            If Cells(rowNumber, 30).Value = "" Then
                Cells(rowNumber, 30).NumberFormat = "YYYY/MM/DD"
                Cells(rowNumber, 30).Value = Cells(rowNumber - 1, 30).Value
            End If
            
            ' AE short_description
            
            ' AF description
            
            ' burton_p2p
            
        End If
        
    Next rowNumber
    
    Call ResizeColumns
 
End Function
 
Private Function Duplicate()
    Sheets("Sheet1").Copy after:=ActiveSheet
    ActiveSheet.Name = "Magento Import"
End Function
 
Private Function ResizeColumns()
    
    ' Resize columns
    Dim numberOfColumns As Long
    Dim columnNumber As Long
    
    numberOfColumns = Sheets("Magento Import").Cells(1, Columns.Count).End(xlToLeft).Column
    
    For columnNumber = 1 To numberOfColumns
        Columns(columnNumber).AutoFit
    Next columnNumber
 
End Function




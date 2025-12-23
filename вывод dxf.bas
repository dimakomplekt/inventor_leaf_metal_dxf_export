Option Explicit

Sub Main()

    ' Error handler setup
    Dim assembly_error As Boolean: assembly_error = False
    Dim has_error As Boolean: has_error = False
    Dim export_error As Boolean: export_error = False

    On Error GoTo handle_common_error

    ' Assembly check
    check_assembly(assembly_error)
    If assembly_error Then GoTo handle_assembly_error


    ' INI-file name for DXF-export setup
    Dim dxf_ini_name As String
    dxf_ini_name = "Под резку.ini"

    Dim current_assembly_path As String ' Path to the current assebly
    current_assembly_path = ThisApplication.ActiveDocument.FullFileName

    Dim current_assembly_path_tmp As String ' tmp
    current_assembly_path_tmp = current_assembly_path

    ' Basic folder of the constructive project
    Dim base_folder_path As String
    base_folder_path = find_base_folder_path(current_assembly_path)

    ' Project structure check
    If (base_folder_path  = "") Then

        has_error = True
        GoTo handle_common_error

    End If


    ' DXF-storage folder by the adress of the basic folder
    Dim dxf_folder_path As String
    dxf_folder_path = find_dxf_folder_path(base_folder_path)

    ' Project structure check
    If (dxf_folder_path = "") Then

        has_error = True
        GoTo handle_common_error

    End If


    ' INI-file path by the DXF-storage folder and INI-file name
    Dim ini_file_path As String 
    ini_file_path = find_ini_file_path(dxf_folder_path, dxf_ini_name)

    ' Project structure check
    If (ini_file_path = "") Then

        has_error = True
        GoTo handle_common_error

    End If


    ' Initialization of the assembly - sheet metal details list
    ' Format [[Assembly name, [SHD_path_1, ... , SHD_path_n], ...]
    Dim assemblies_with_sheet_metal As Object
    assemblies_with_sheet_metal = CreateObject("System.Collections.ArrayList")

    ' Root assembly flag
    Dim is_root As Boolean
    is_root =  True

    ' Dictionary for the assemblies repeat check
    Dim added_assemblies As Object
    added_assemblies = CreateObject("Scripting.Dictionary")


    ' Fill the list with assemblies and sheet metal details by the function
    Call collect_assemblies_and_sheet_metal(ThisApplication.ActiveDocument.ComponentDefinition.Occurrences, assemblies_with_sheet_metal, is_root, ThisApplication.ActiveDocument, added_assemblies)

    ' Log the obtained list
    Call log_assemblies_with_sheet_metal(assemblies_with_sheet_metal)


    ' DXF export by the information from the assemblies_with_sheet_metal
    export_dxf_for_curerrent_assembly(assemblies_with_sheet_metal, dxf_folder_path, ini_file_path, export_error)


    ' Error handler
    If export_error Then GoTo handle_export_error


    ' Success
    MsgBox("Программа выполнена успешно")


    ' Reopen the current assembly  
    Dim doc As Document
    doc = ThisApplication.Documents.Open(current_assembly_path_tmp, True) ' True → открыть и активировать
    doc.Activate


    ' Standart exit
    Exit Sub

' Error handlers
handle_common_error:
    MsgBox("Произошла системная ошибка: " & Err.Description, vbCritical, "Error")
    Exit Sub

handle_assembly_error:
    MsgBox("Open an assembly before running this rule.", vbCritical, "Error")
    Exit Sub

handle_export_error:
    MsgBox("Export error!", vbCritical, "Error")
    Exit Sub
End Sub


' Helper-sub to check if the current document is assembly
Sub check_assembly(ByRef assembly_error As Boolean)

    If ThisApplication.ActiveDocument Is Nothing Or _
       ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then

        assembly_error = True
        Exit Sub

    End If

End Sub


' Helper-function to find the root folder of the project
Function find_base_folder_path(current_assembly_path As String) As String

    ' Adress for return
    Dim project_root As String
    project_root = ""

    ' Directory as object for GetParentFolderName methods use
    Dim fso As Object
    fso = CreateObject("Scripting.FileSystemObject")

    ' Current assembly path
    Dim cur As String
    cur = current_assembly_path

    ' Iterative root folder search
    Do While (cur <> "")

        ' Project root attributes check
        If (Dir(cur & "\3_Модели", vbDirectory) <> "") And (Dir(cur & "\4_Чертежи", vbDirectory) <> "") Then

            project_root = cur
            Exit Do

        End If

        ' Exit if we fail to the disk root
        If (fso.GetParentFolderName(cur) = "") Then Exit Do
        
        ' New iteration setting
        cur = fso.GetParentFolderName(cur)
    Loop

    ' If the project root not found
    If (project_root = "") Then

        MsgBox("Project root not found (expected folders: 3_Модели and 4_Чертежи).", vbCritical, "Error")
        find_base_folder_path = ""  ' Empty string return
        Exit Function

    End If

    find_base_folder_path = project_root ' Return of the project root adress
End Function


' Helper-function for DXF-folder search
Function find_dxf_folder_path(base_folder_path As String) As String

    ' Variable to return
    Dim dxf_folder As String
    dxf_folder = base_folder_path & "\4_Чертежи\3_DXF"

    ' Check the DXF folder existence 
    If Dir(dxf_folder, vbDirectory) = "" Then
        ' Log allert and return empty string 
        MsgBox("Project folders not found (expected folders: 4_Чертежи and 3_DXF).", vbCritical, "Error")
        find_dxf_folder_path = ""

    Else
        ' Return the path if the folder exists
        find_dxf_folder_path = dxf_folder

    End If

End Function


' Helper-function to find the ini-file for DXF-export
Function find_ini_file_path(dxf_folder_path As String, dxf_ini_name As String) As String

    ' Variable for return
    Dim ini_path As String
    ini_path = dxf_folder_path & "\" & dxf_ini_name

    ' Check the existence of the 
    If Dir(ini_path) = "" Then
        ' Log allert and return empty string 
        MsgBox("Project file not found (expected file: " & dxf_ini_name & ").", vbCritical, "Error")
        find_ini_file_path = ""
    Else
        ' Return the path
        find_ini_file_path = ini_path
    End If

End Function


' Recursive sub for the filling of the assembly - sheet_metal_details list 
Sub collect_assemblies_and_sheet_metal(occurrences As Object, assemblies As Object, is_root As Boolean, root_assembly As Object, added_assemblies As Object)
    
    
    Dim assembly_path As String
    Dim assembly_name As String
    
    ' Get the name and adress of the root assembly
    If is_root And Not root_assembly Is Nothing Then

        assembly_path = root_assembly.FullFileName
        assembly_name = root_assembly.DisplayName

    ' Fill the current element of the output list information variables
    Else
        If occurrences.Count > 0 Then
            assembly_path = occurrences(1).Parent.Document.FullFileName
            assembly_name = occurrences(1).Parent.Document.DisplayName
        Else
            assembly_path = ""
            assembly_name = "Unnamed"
        End If
    End If
    
    ' Dictionary check for the assemblies repeats preventing
    ' Exit works fine, cause the recursive subsub calls by the for loop,
    ' and after exit we call the collect_assemblies_and_sheet_metal for the next iteration
    If added_assemblies.Exists(assembly_path) Then Exit Sub

    ' Add the element to dictionary for the assemblies repeats preventing
    Call added_assemblies.Add(assembly_path, True)
    
    ' Create the assembly entry
    Dim current_entry(1) As Object
    current_entry(0) = assembly_name
    current_entry(1) = CreateObject("System.Collections.ArrayList")
    
    ' Fill the sheet metal details list of the current assembly
    Dim occ As Object
    For Each occ In occurrences

        ' Sheet metal check
        If TypeName(occ.Definition) = "SheetMetalComponentDefinition" Then

            ' Check the path to the detail
            Dim full_path As String
            full_path = occ.Definition.Document.FullFileName

            ' Details repeats preventing
            If Not current_entry(1).Contains(full_path) Then
                Call current_entry(1).Add(full_path)

            End If
        End If
    Next
    
    ' Add the assebly to the output list if there are sheet metal details in assembly
    If current_entry(1).Count > 0 Then
        Call assemblies.Add(current_entry)
    End If
    
    ' Recursive step forward
    For Each occ In occurrences
        If TypeName(occ.Definition) = "AssemblyComponentDefinition" Then
            Call collect_assemblies_and_sheet_metal(occ.Definition.Occurrences, assemblies, False, Nothing, added_assemblies)
        End If
    Next

End Sub


' Sub for the DXF export by the assemblies_with_sheet_metal list
Sub export_dxf_for_curerrent_assembly(assemblies_with_sheet_metal As Object, dxf_folder_path As String, ini_file_path As String, ByRef export_error As Boolean)

    ' Current assebly variable for iteration
    Dim asm_entry As Object
    ' Dim i As Long

    ' Iteration by the assebly
    For Each asm_entry In assemblies_with_sheet_metal

        ' Output folder setup
        Dim clean_subassembly_name As String
        clean_subassembly_name = asm_entry(0)


        ' Drop out the bad sybols logic 1
        If InStrRev(clean_subassembly_name, ".") > 0 Then
            clean_subassembly_name = Left(clean_subassembly_name, InStrRev(clean_subassembly_name, ".") - 1)
        End If

        ' Drop out the bad sybols logic 2
        Dim invalidChars As String
        invalidChars = "/\[]*?<>|:"

        For k = 1 To Len(invalidChars)
            clean_subassembly_name = Replace(clean_subassembly_name, Mid(invalidChars, k, 1), "_")
        Next

        ' Drop out the bad sybols logic 3
        Dim arr() As String
        arr = Split(clean_subassembly_name, "_")

        ' Drop out the bad sybols logic 4
        If UBound(arr) > 0 Then
            If IsNumeric(arr(UBound(arr))) Then
                ReDim Preserve arr(UBound(arr) - 1)
                clean_subassembly_name = Join(arr, "_")
            End If
        End If

        
        ' Subfolder path generation by the clean name
        Dim subfolder_path As String
        subfolder_path = dxf_folder_path & "\" & clean_subassembly_name

        ' Subfolder creation
        If Dir(subfolder_path, vbDirectory) = "" Then MkDir(subfolder_path)


        ' Preparation of the details list for the current assembly
        Dim sheet_parts As Object
        sheet_parts = asm_entry(1)

        Dim part_as_str As Object
        Dim part_index As Integer
        part_index = 1


        ' Nested iteration by the current detail in current assembly
        For Each part_as_str In sheet_parts

            ' Error handler
            On Error Resume Next

            ' Open the part document as an app object
            Dim part_as_component As Object
            part_as_component = ThisApplication.Documents.Open(part_as_str, False)

            ' If the part is opened
            If Not part_as_component Is Nothing Then

                ' Get the properties of the current detail file
                Dim part_definition As Object
                part_definition = part_as_component.ComponentDefinitions(1)

                ' If the detail is sheet metal
                If TypeName(part_definition) = "SheetMetalComponentDefinition" Then

                    ' Another open of the doc (wht can i do, i'm not the inventor developer...)
                    Dim sm_detail_document As Object
                    sm_detail_document = part_definition.Document

                    ' Unfold the sheet metal detail
                    If Not part_definition.HasFlatPattern Then

                        part_definition.Unfold()
                        part_definition.FlatPattern.ExitEdit()

                    End If

                    ' Create the drawing which will be used for the DXF export
                    Dim sm_drawing As Object
                    sm_drawing = ThisApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, , False)

                    ' Set the measurement units as the millimeters
                    sm_drawing.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits

                    ' Setup the view
                    Dim detail_tr_geom As Object
                    detail_tr_geom = ThisApplication.TransientGeometry

                    Dim o_X As Double
                    o_X = 0

                    Dim o_Y As Double
                    o_Y = 0

                    ' Set the base view
                    Dim detail_base_view_options As Object
                    detail_base_view_options = ThisApplication.TransientObjects.CreateNameValueMap()

                    ' Add the base view
                    Call detail_base_view_options.Add("SheetMetalFoldedModel", False)

                    ' Sheet variable
                    Dim o_sheet As Object
                    o_sheet = sm_drawing.ActiveSheet

                    ' View setup variable by the AddBaseView call
                    Dim o_view As Object

                    o_view = o_sheet.DrawingViews.AddBaseView(sm_detail_document, detail_tr_geom.CreatePoint2d(o_X, o_Y), 1, _
                        ViewOrientationTypeEnum.kDefaultViewOrientation, _
                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                        "FlatPattern", , detail_base_view_options)


                    ' Thickness of the detail check routine (wht can i do, i'm not the inventor developer...)
                    Dim thickness_mm As Double
                    Dim thickness_str As String
                    Dim dot_position As Long
                    Dim s As String

                    thickness_mm = part_definition.Thickness.Value * 10
                    s = CStr(thickness_mm)
                    dot_position = InStr(s, ".")
                    If dot_position = 0 Then dot_position = InStr(s, ",")
                    If dot_position > 0 Then
                        thickness_str = Left(s, dot_position + 2)
                    Else
                        thickness_str = s
                    End If


                    ' Add the thichness to the label
                    o_view.Name = sm_detail_document.DisplayName & " - " & thickness_str & " мм"
                    ' Show the label
                    o_view.ShowLabel = True


                    ' Remove the bend lines from the view by the helper-sub
                    Call remove_bend_lines(o_view, part_definition.FlatPattern)


                    ' Setup the part name for the export 

                    ' Clean the part name
                    Dim part_file_name As String
                    part_file_name = Mid(part_as_str, InStrRev(part_as_str, "\") + 1)
                    
                    If InStrRev(part_file_name, ".") > 0 Then
                        part_file_name = Left(part_file_name, InStrRev(part_file_name, ".") - 1)
                    End If

                    For n = 1 To Len(invalidChars)
                        part_file_name = Replace(part_file_name, Mid(invalidChars, n, 1), "_")
                    Next


                    ' Fill the DXF name by the cleaned detail name and some text with format, like "_text.dxf" 
                    Dim detail_DXF_name As String
                    detail_DXF_name = subfolder_path & "\" & part_file_name & "_развертка.dxf"


                    ' Save the DXF file of the detail to the folder for it's assebly
                    Call SaveDXF(sm_drawing, detail_DXF_name, ini_file_path)


                    ' Close the drawing for the current iteration
                    Call sm_drawing.Close() 


                    ' Change the part index (increment iterator)
                    part_index = part_index + 1
                End If

                ' Close the current detail
                Call part_as_component.Close(True)
            End If

            ' Error handler
            On Error GoTo 0
        Next
    Next

    MsgBox("DXF Export Completed", vbInformation)
End Sub



' Helper-sub to remove bend lines 
Sub remove_bend_lines(o_view As DrawingView, o_flat_pattern As FlatPattern)

    ' Upward bend edges variable by the FlatPatternEdgeTypeEnum.kBendUpFlatPatternEdg
    Dim o_bend_edges_up As Edges
    o_bend_edges_up = o_flat_pattern.GetEdgesOfType(FlatPatternEdgeTypeEnum.kBendUpFlatPatternEdge)

    ' Downward bend edges variable by the FlatPatternEdgeTypeEnum.kBendDownFlatPatternEdg
    Dim o_bend_edges_down As Edges
    o_bend_edges_down = o_flat_pattern.GetEdgesOfType(FlatPatternEdgeTypeEnum.kBendDownFlatPatternEdge)

    Dim o_edge As Edge
    Dim o_curve As DrawingCurve
    Dim o_segment As DrawingCurveSegment
    
    ' Change the visibility of the curves segments
    For Each o_edge In o_bend_edges_up
        For Each o_curve In o_view.DrawingCurves(o_edge)
            For Each o_segment In o_curve.Segments
                o_segment.Visible = False
            Next
        Next
    Next

    ' Change the visibility of the curves segments
    For Each o_edge In o_bend_edges_down
        For Each o_curve In o_view.DrawingCurves(o_edge)
            For Each o_segment In o_curve.Segments
                o_segment.Visible = False
            Next
        Next
    Next
End Sub


' Helper-sub for DXF-saving by drawing, file name and INI-file path
Sub SaveDXF(sm_drawing As DrawingDocument, o_file_name As String, o_ini_file As String)

	' A reference to the DFX translator
	Dim DXF_add_in As TranslatorAddIn
	DXF_add_in = ThisApplication.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")

	' Create translation context
	Dim o_context As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext
	o_context.Type = IOMechanismEnum.kFileBrowseIOMechanism

	' Create options for the translation
	Dim o_options As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap

	'Create a o_data_medium object
	Dim o_data_medium As DataMedium = ThisApplication.TransientObjects.CreateDataMedium

	' The options (which .ini-file to use)
	If DXF_add_in.HasSaveCopyAsOptions(sm_drawing, o_context, o_options) Then
		o_options.Value("Export_Acad_IniFile") = o_ini_file
	End If

	' The filename property of the o_data_medium object
	o_data_medium.FileName = o_file_name
    
    ' SaveCopyAs call
    DXF_add_in.SaveCopyAs(sm_drawing, o_context, o_options, o_data_medium)
End Sub


' Text file creation sub
Sub CreateTXT(oText As String, o_file_name As String)
    Dim oTxtWriter As Object
    oTxtWriter = CreateObject("Scripting.FileSystemObject").CreateTextFile(o_file_name, True)
    oTxtWriter.WriteLine(oText)
    oTxtWriter.Close
End Sub


' Log the assemblies_with_sheet_metal list
Sub log_assemblies_with_sheet_metal(assemblies_with_sheet_metal As Object)

    Dim msg As String
    msg = "=== assemblies_with_sheet_metal ===" & vbCrLf

    Dim entry As Object
    Dim sheetPart As Object

    For Each entry In assemblies_with_sheet_metal
        msg = msg & "Subassembly: " & entry(0) & vbCrLf
        For Each sheetPart In entry(1)
            msg = msg & "  sheet part: " & sheetPart & vbCrLf
        Next
    Next

    MsgBox(msg, vbInformation, "assemblies_with_sheet_metal log")
End Sub

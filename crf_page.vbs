' Created By Kai Zhou
' tips : Range.FindNext does not work in Sub? 
' tips : Parentheses should not be used when calling sub with parameters

Option Explicit
' Global variables
Dim app, workbook, workbook2, sheet
Dim row,col,filename, filename2, msg
Dim fso, current_directory
Dim domains(100)
Dim dm_count
Dim logfile

Dim current_dt, current_time

current_dt = formatdatetime(date(), 2)
current_time = formatdatetime(time(), 4)

'msgbox current_dt & "T" & current_time

Set app = WScript.CreateObject("Excel.Application")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
current_directory = fso.GetAbsolutePathName(".")' dot means current path

'Set logfile = fso.CreateTextFile(current_directory & "\compareResult" & msgbox current_dt & "T" & current_time&".log", True)


'filename  = current_directory+"\"+ InputBox("Please input spec name (include file extension)")
'filename2 = current_directory+"\"+ InputBox("Please input crf page file name (include file extension)")

filename  = current_directory+"\"+  "229288 sdtm mapping specifications v0.1_20160913.xlsx"
filename2 = current_directory+"\"+  "test.xlsx"

call main()

Sub  main()
	msgbox(filename)
	dm_count = 0
	Set workbook  = get_file(filename)
    Set workbook2 = get_file(filename2)

    'obtain all domains
    Dim define_sheet, dm, i
    Set define_sheet = workbook.Sheets("Define_DATADEF")
    for i = 6 to 100
    	dm = Trim(define_sheet.Range("C"&i).value)
    	if dm <> "" then
    		domains(i-5) = dm
    		dm_count = dm_count + 1
    	else 
    		exit For
     	end if 
    next

    dm_count = 1

    ' Search every domain and every variable which's origin in CRF
    for i = 1 to dm_count
    	Dim rg1, rg2, spec_crf_page
    	Dim domain, var
    	Dim firstaddress

    	domain = domains(i)
    	domain = "CO"

    	if domain <> "" then
    		Set rg1 = workbook.Sheets(domain).Range("J:J").find("CRF", ,-4163)

    		if not rg1 is nothing then
    			firstaddress = rg1.Address
	    		do 
	    			Set var = workbook.Sheets(domain).range("D" & rg1.row) 'get variable from spec
	    			Set spec_crf_page = workbook.Sheets(domain).range("S" & rg1.row)
	    			'msgbox var
	    			Set rg2 = find_page_range(workbook2, domain, Trim(var.value)) ' Search this variable in page file
	    			if rg2 is nothing then ' if cannot find in page file
	    				spec_crf_page.AddComment("Cannot find " & domain & "." & Trim(var.value) & " in page file")
	    				spec_crf_page.Interior.ColorIndex = 3 ' red
	    			else ' if found it
	    				if Trim(spec_crf_page.value) <> Trim(rg2.value) then 
                            spec_crf_page.AddComment("Updated, please double check, former value is " & spec_crf_page.value)
                            'logfile.write()
	    					spec_crf_page.value = rg2.value
	    					spec_crf_page.Interior.ColorIndex = 6 'yellow 
	    				end if

	    				'msgbox "find " & domain & "." & var & " in page file"
	    			end if 
	    			Set rg1 = workbook.Sheets(domain).Range("J:J").find("CRF", rg1 ,-4163) ' next variable

	    			if rg1 is nothing then msgbox "here---"
	    		loop while not rg1 is nothing and rg1.address <> firstaddress
	    	end if 
    	else 
    		exit for 
    	end if
    next

    'reverse search to check if variable exists in page file but does not exist in spec
    Dim page_rg, mycell
    'Set cell = Nothing
    Set page_rg = workbook2.Sheets(1).Range("A:A")

    'if page_rg is Nothing then msgbox "Nothing"

    for each mycell in page_rg
    	if mycell.value = "" then exit for
    	Dim t_sheet, t_rg1, t_rg2, t_rg3
    	Dim sheet_exist

    	Set t_sheet = Nothing
    	Set t_rg1 = Nothing
    	Set t_rg2 = Nothing
    	Set t_rg3 = Nothing

    	Set t_rg3 = workbook2.Sheets(1).Range("B" & mycell.row)

    	' test if domain is in the spec
    	sheet_exist = sheetExists(Trim(mycell.value), workbook.Worksheets)
    	if Not sheet_exist then 
    		mycell.AddComment("Cannot find the domain " & mycell.value & " in spec")
    		mycell.Interior.ColorIndex = 37
    		'msgbox "cannot find the domain " & mycell.value & " in spec"
    	else
    		Set t_sheet = workbook.Sheets(Trim(mycell.value))
    		Set t_rg1 = t_sheet.Range("D:D").find(Trim(t_rg3.value), ,-4163, 1)
    		if Not t_rg1 is Nothing then 
    			Set t_rg2 = t_sheet.Range("J" & t_rg1.row)
    			if Instr(t_rg2.value, "CRF") = 0 then Set t_rg2 = Nothing
    		end if

    	'if Not t_rg1 is Nothing and Not t_rg2 is Nothing then msgbox "find" & t_rg1.value & t_rg3.value
	    	if t_rg1 is nothing or (Not t_rg1 is nothing and t_rg2 is Nothing) then 
	    		t_rg3.AddComment("Cannot find this variable in spec")
	    		t_rg3.Interior.ColorIndex = 3
	    	end if 
    	end if
    next

    ' VALDEF sheet check
    Dim spec_valdef_sheet, spec_valdef_rg, spec_valdef_row
    Dim page_cell

    Set spec_valdef_sheet = workbook.sheets("VALDEF")
    Set spec_valdef_rg = spec_valdef_sheet.Range("A2:O600")

    for each spec_valdef_row in spec_valdef_rg.rows
        if spec_valdef_row.cells(1).value = "" then exit for

        if mid(spec_valdef_row.cells(1).value, 1, 2) <> "DS" and mid(spec_valdef_row.cells(1).value, 1, 2) <> "LB" _
        and mid(spec_valdef_row.cells(1).value, 1, 2) <> "QS" then

            'if trim(spec_valdef_row.cells(1).value) = "FA.FATESTCD" then 
                Set page_cell = page_autofilter(workbook2.Sheets(2), spec_valdef_row.cells(1).value, spec_valdef_row.cells(2).value)
                if page_cell is Nothing then 
                    spec_valdef_row.cells(14).Interior.ColorIndex = 3
                    spec_valdef_row.cells(14).AddComment("Cannot find this variable in page file")
                elseif Not page_cell is Nothing then
                    if Trim(page_cell.text) <> Trim(spec_valdef_row.cells(14).text) then 
                        spec_valdef_row.cells(14).AddComment("Updated, please double check, former value is " & _
                            spec_valdef_row.cells(14).value)
                        spec_valdef_row.cells(14).value = page_cell.value
                        spec_valdef_row.cells(14).Interior.ColorIndex = 6
                    end if 
                end if 
            'end if
        end if 
    next

    ' reverse VALDEF sheet check
    Dim page_valdef_sheet, page_valdef_rg, page_valdef_row
    Dim spec_cell

    Set page_valdef_sheet = workbook2.Sheets(2)
    Set page_valdef_rg = page_valdef_sheet.Range("A2:O600")

    for each page_valdef_row in page_valdef_rg.rows
        if page_valdef_row.cells(1).value = "" then exit for

        if mid(page_valdef_row.cells(1).value, 1, 2) <> "DS" and mid(page_valdef_row.cells(1).value, 1, 2) <> "LB" _
        and mid(page_valdef_row.cells(1).value, 1, 2) <> "QS" then
            Set spec_cell = spec_autofilter(workbook.Sheets("VALDEF"), page_valdef_row.cells(1).value, page_valdef_row.cells(2).value)
            if spec_cell is Nothing then 
                page_valdef_row.cells(3).Interior.ColorIndex = 3
                page_valdef_row.cells(3).AddComment("Cannot find this variable in spec file")
            end if
        end if
    next

    'logfile.Close
	workbook.Close(true)
	workbook2.Close(true)
End Sub


Function get_file(path)
	Set get_file = app.WorkBooks.Open(path)
End Function


Function find_range(wb, domain, var)
	Dim sheet
	Dim rg1, rg2, rg3
	Set sheet = wb.Sheets(domain)

	Set rg1 = sheet.Range("D:D").Find(var, , -4163, 1 )

	If Not rg1 is Nothing Then
		'msgbox rg1.Cells(1,1).row
		Set rg2 = sheet.Cells(rg1.Cells(1,1).row, 19)
		msgbox(rg2.value)
	End if
End Function

Function find_page_range(wb, domain, var)
	Dim sheet
	Dim rg1, rg2, rg3
	Dim varname
	Dim firstaddress
	Set sheet = wb.Sheets(1)

	Set rg1 = sheet.Range("A:A").Find(domain, , -4163, 1)
	if Not rg1 is nothing then 
		varname = Trim(sheet.Range("B" & rg1.row).value)
		firstaddress = rg1.Address

		if varname <> var then 
			Do
				varname = ""
				Set rg1 = sheet.Range("A:A").Find(domain, rg1 , -4163, 1)
				if Not rg1 is nothing and rg1.Address <> firstaddress then 
					varname = Trim(sheet.Range("B" & rg1.row).value)
				end if
			loop while Not rg1 is nothing and rg1.Address <> firstaddress and varname <> var

			if rg1.Address = firstaddress then Set rg1 = Nothing
		end if
	end if

	if Not rg1 is nothing then 
		Set find_page_range = sheet.Range("C" & rg1.row)
	else 
		Set find_page_range = Nothing
	end if
End Function

Function sheetExists(sheetToFind, Worksheets)
    sheetExists = False
    Dim sheet
    For Each sheet In Worksheets
        If sheetToFind = sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next
End Function

Function page_autofilter(page_sheet, var1, var2)
    Dim rng 

    Set rng = Nothing

    'disable autofilter in case it's already enabled'
    page_sheet.AutoFilterMode = False 

    With page_sheet.Range("A1:C1")
        'set autofilter'
        .AutoFilter 1, "="&var1
        .AutoFilter 2, "="&var2
    End With

    With page_sheet

        'On Error Resume Next
        Set rng = .Range("C2:C600").SpecialCells(12)
        'On Error GoTo 0
        'msgbox rng.value
    End With

    if Not rng is Nothing then 
        Set page_autofilter = rng
        if trim(rng.text) = "" then Set page_autofilter = Nothing
    else 
        Set page_autofilter = Nothing
    end if

    ' close again
    page_sheet.AutoFilterMode = False 
End Function


Function spec_autofilter(page_sheet, var1, var2)
    Dim rng 
    Set rng = Nothing

    'disable autofilter in case it's already enabled'
    page_sheet.AutoFilterMode = False 

    With page_sheet.Range("A1:N1")
        'set autofilter'
        .AutoFilter 1, "="&var1
        .AutoFilter 2, "="&var2
    End With

    With page_sheet

        'On Error Resume Next
        Set rng = .Range("N2:N600").SpecialCells(12)
        'On Error GoTo 0
        'msgbox rng.value
    End With

    if Not rng is Nothing then 
        Set spec_autofilter = rng
        if trim(rng.text) = "" then Set spec_autofilter = Nothing
    else 
        Set spec_autofilter = Nothing
    end if

    ' close again
    page_sheet.AutoFilterMode = False 
End Function
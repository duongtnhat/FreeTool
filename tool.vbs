Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = "[A-Z]{5}[0-9]{3}"
Set WshShell = CreateObject("WScript.Shell")
strCurDir = WshShell.CurrentDirectory
Set objFSO=CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(strCurDir & "\out.csv", True)
Set allDate = CreateObject("System.Collections.ArrayList")
Set dataList = CreateObject("Scripting.Dictionary")

For Each fileName In objFSO.GetFolder(strCurDir).Files
  If LCase(objFSO.GetExtensionName(fileName.Name)) = "xls" Then
    Set objExcel = CreateObject("Excel.Application")
    'fileName = "bien_ban_doi_soat_S18759785_Cong_ty_Co_phan_Duoc_pham_Pharmacity_1629705602_856.xls"
    Set objWorkbook = objExcel.Workbooks.Open(fileName)

    intRow = 27
    ngay = objExcel.Cells(10, 5).Value
    tienCod = objExcel.Cells(12, 8).Value
    phiDV = objExcel.Cells(13, 8).Value
    soDon = 0
    allDate.add ngay

    Do Until objExcel.Cells(intRow,1).Value = ""
      row = objExcel.Cells(intRow, 8).Value
      Set colMatches = objRegEx.Execute(row)
      If colMatches.Count > 0 Then
        For Each strCode in colMatches
          add ngay, strCode
        Next
      Else
        Wscript.Echo fileName & " not found code at line " & intRow
      End If

      intRow = intRow + 1
      soDon = soDon + 1
    Loop

    Wscript.Echo "File: " & fileName
    Wscript.Echo "Ngay: " & ngay
    Wscript.Echo "Tien Cod: " & tienCod
    Wscript.Echo "Phi DV: " & phiDV
    Wscript.Echo "So don: " & soDon
    Wscript.Echo "================================================================================"
    addValue ngay, ".Tien Cod", tienCod
    addValue ngay, ".PhiDV", phiDv
    addValue ngay, ".So don", soDon

    objExcel.Quit
  end if
next

allDate.sort

rowWrite = "Ngay,"
For i = 0 To allDate.Count-1
    rowWrite = rowWrite & allDate(i) & ","
next
objFile.Write rowWrite & vbCrLf

for each code in BubbleSort(dataList.keys)
  Set data = dataList.Item(code)
  rowWrite = code & ","
  For i = 0 To allDate.Count-1
    dateShow = allDate(i)
    if data.exists(dateShow) then
      rowWrite = rowWrite & data(dateShow) & ","
    else
      rowWrite = rowWrite & ","
    end if
  next
  objFile.Write rowWrite & vbCrLf
next

objFile.Close

function add(date, codeRaw)
  code = UCase(codeRaw)
  if not dataList.Exists(code) then
    set data = CreateObject("Scripting.Dictionary")
    data.Add date, 1
    dataList.add code, data
  else
    Set data = dataList.Item(code)
    if not data.Exists(date) then
      data.add date, 1
    else
      numberNow = data(date)
      data.remove date
      data.add date, (numberNow + 1)
    end if
  end if
end function

function addValue(date, codeRaw, value)
  code = UCase(codeRaw)
  if not dataList.Exists(code) then
    set data = CreateObject("Scripting.Dictionary")
    data.Add date, value
    dataList.add code, data
  else
    Set data = dataList.Item(code)
    if not data.Exists(date) then
      data.add date, value
    end if
  end if
end function

Function BubbleSort(arrValues)
  Dim j, k, Temp
  For j = 0 To UBound(arrValues) - 1
    For k = j + 1 To UBound(arrValues)
      If strComp(arrValues(j),arrValues(k)) < 0 Then
        Temp = arrValues(j)
        arrValues(j) = arrValues(k)
        arrValues(k) = Temp
      End If
    Next
  Next
  BubbleSort = arrValues
End Function















Attribute VB_Name = "prj_Form"
Option Explicit
Sub GetFormCode_ReasonForAbsence()

    On Error Resume Next
    Dim strID As String, strName As String
    strID = celDC.Offset(, intOffset(celDC, "코드"))
    strName = celDC.Offset(, intOffset(celDC, "성명"))
    On Error GoTo 0

    If strID = "" Or strName = "" Then Exit Sub

    Dim dv As DBvar
    With dv
        .Alias = "EDU_LMS"
        .Table = "lms_edu_exempt"

        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)

        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .db = arrDB(4)

        .Query = "SELECT EXEMPT_CD FROM " & .Table & " WHERE WATV_ID = '" & strID & "' ORDER BY EXEMPT_CD;"
        .arrQuery = Array(.Query)

        Dim rt As Integer
        Dim arr As Variant
        Dim blnExist As Boolean
        blnExist = Application.WorksheetFunction.CountIf(GetTwbRange("DB_교육면제[코드]"), strID) > 0

ReTry:
        If blnExist Then
            If rt < 3 Then
                .arrData = SelectQueryArr(.arrQuery, .Type, .File)
                rt = rt + 1
                If IsArrayEmpty(.arrData) Then GoTo ReTry
            Else
                Exit Sub
            End If
            arr = .arrData(0)
        End If

    End With

    With frm_AddReasonForAbsence
        .cmb_Code.Clear

        If Not IsArrayEmpty(arr) Then
            Dim i As Long
            For i = 0 To UBound(arr, 2)
                .cmb_Code.AddItem arr(0, i)
            Next
            .cmb_Code.ListIndex = 0
        End If

pass:
        .txt_ID.value = strID
        .txt_Name.value = strName
        .Show
    End With

End Sub
Sub GetIndividualCourseStatus()
    Dim dv As DBvar
    Dim rng As Range, rngSelect As Range, celTarget As Range
    
    Dim strID As String
    strID = celDC.Offset(, intOffset(celDC, "코드"))
    
    With dv
        .Alias = "현재워크북"
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .Query = "SELECT T2.명칭, T1.명칭, T3.이수일 FROM ([DB_챕터$] AS T1 LEFT JOIN [DB_시즌$] AS T2 ON T2.코드 = T1.시즌코드) LEFT JOIN (SELECT * FROM [DB_수강현황$] WHERE 코드 = '" & strID & "') AS T3 ON T3.시즌코드 = T1.시즌코드 AND T3.챕터 = T1.챕터;"
        
        .arrQuery = Array(.Query)
        
        Dim rt As Integer
        Dim arr As Variant
ReTry:
        If rt < 3 Then
            .arrData = SelectQueryArr(.arrQuery, .Type, .File, , , , , , , True)
            rt = rt + 1
            If IsArrayEmpty(.arrData) Then GoTo ReTry
        Else
            Exit Sub
        End If
        
        arr = .arrData(0)

    End With
    
    With frm_IndividualCourseStatus
        
        .txt_ID = strID
        .txt_Name = celDC.Offset(, intOffset(celDC, "성명"))
        .txt_Position = celDC.Offset(, intOffset(celDC, "직책"))
        .txt_Duty = celDC.Offset(, intOffset(celDC, "직분"))
        .txt_Season = celDC.Offset(, intOffset(celDC, "시즌코드"))
        .txt_tSeasonPct = celDC.Offset(, intOffset(celDC, "이수"))
        .txt_aSeasonPct = celDC.Offset(, intOffset(celDC, "이수2"))
        
        .lst_CourseStatus.ColumnCount = UBound(arr, 2) + 1
        .lst_CourseStatus.List = arr
        .lst_CourseStatus.ColumnWidths = GetAutoColumnWidths(arr)
        .Show
    End With
 End Sub


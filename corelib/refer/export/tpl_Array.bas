Attribute VB_Name = "tpl_Array"
Option Explicit
Function ConvertToArrData(arr As Variant) As Variant
    Dim arrData() As Variant  ' 괄호만 있고 크기 지정 안 함
    Dim dataIndex As Long
    dataIndex = 0
    
    ' 먼저 총 요소 개수를 계산
    Dim totalCount As Long
    totalCount = CountElements(arr)
    
    ' totalCount가 0보다 클 때만 배열 크기 재설정
    If totalCount > 0 Then
        ReDim arrData(0 To totalCount - 1)
        
        ' 재귀적으로 배열 평면화
        FlattenArray arr, arrData, dataIndex
    End If
    
    ConvertToArrData = arrData
End Function

' 배열의 모든 요소를 평면화하는 서브루틴
Private Sub FlattenArray(arr As Variant, ByRef arrData() As Variant, ByRef dataIndex As Long)
    Dim i As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    
    ' 배열인지 확인
    If Not IsArray(arr) Then
        arrData(dataIndex) = arr
        dataIndex = dataIndex + 1
        Exit Sub
    End If
    
    On Error GoTo NotAnArray
    lowerBound = LBound(arr)
    upperBound = UBound(arr)
    On Error GoTo 0
    
    ' 배열인 경우 각 요소를 재귀적으로 처리
    For i = lowerBound To upperBound
        If IsArray(arr(i)) Then
            FlattenArray arr(i), arrData, dataIndex
        Else
            arrData(dataIndex) = arr(i)
            dataIndex = dataIndex + 1
        End If
    Next i
    
    Exit Sub
    
NotAnArray:
    arrData(dataIndex) = arr
    dataIndex = dataIndex + 1
End Sub

' 배열의 총 요소 개수를 계산하는 함수
Private Function CountElements(arr As Variant) As Long
    Dim count As Long
    Dim i As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    
    count = 0
    
    ' 배열인지 확인
    If Not IsArray(arr) Then
        CountElements = 1
        Exit Function
    End If
    
    On Error GoTo NotAnArray
    lowerBound = LBound(arr)
    upperBound = UBound(arr)
    On Error GoTo 0
    
    ' 배열인 경우
    For i = lowerBound To upperBound
        If IsArray(arr(i)) Then
            count = count + CountElements(arr(i))
        Else
            count = count + 1
        End If
    Next i
    
    CountElements = count
    Exit Function
    
NotAnArray:
    CountElements = 1
End Function

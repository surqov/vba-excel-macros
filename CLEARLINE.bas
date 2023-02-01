''Attribute VB_Name = "CLEARLINE"
''https://github.com/surqov
''https://excelium.ru/
'Кодировка = UTF-8'

Private Function RepSymb(ByVal line_ As String, Optional dic_ As Variant, Optional check_doubles As Boolean) As String
    Dim IsDict As Boolean, CheckDoubles As Boolean, ResultLine As String
    'Checking parameters
    IsDict = Not IsMissing(dic_) And TypeName(dic_) = "String"
    If (Not IsMissing(dic_) And Not TypeName(dic_) = "String") Then
        CheckDoubles = CBool(dic_)
    ElseIf (Not IsMissing(check_doubles)) Then
        CheckDoubles = check_doubles
    Else
        CheckDoubles = False
    End If
    
    'Creating RegExp object
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True: objRegExp.IgnoreCase = True
    
    'Checing if there's additional dictonary got from user
    If (IsDict) Then
        Dim EscapedDic As String
        objRegExp.Pattern = "[-/\\^$*+?()|[\]{}]"
        EscapedDic = objRegExp.Replace(CStr(dic_), "\$&")
        objRegExp.Pattern = "[^a-zA-Z0-9а-яА-Я" & EscapedDic & "]*"
    Else
    'Using standart symbols ranges if there no dictonary got from user
        objRegExp.Pattern = "[^a-zA-Z0-9а-яА-Я]*"
    End If
    
    'Implementing RegExp
    ResultLine = objRegExp.Replace(line_, "")
    
    'Checking if we need to remove repeated symbols and doing it
    If (CheckDoubles) Then
        objRegExp.Pattern = "(.)(?=\1)"
        ResultLine = objRegExp.Replace(ResultLine, "")
    End If
    
    'Returning clear string
    RepSymb = ResultLine
End Function

Public Function CLEARLINE(ByVal Rng As Range, Optional dic_ As Variant, Optional check_doubles As Boolean) As Variant
    Dim avArr(), lr As Long, lc As Long
    If Rng.Count = 1 Then
        ReDim avArr(1, 1): avArr(1, 1) = Rng.Value
    Else
        avArr = Rng.Value
    End If
    For lr = 1 To UBound(avArr, 1)
        For lc = 1 To UBound(avArr, 2)
            avArr(lr, lc) = RepSymb(avArr(lr, lc), dic_, check_doubles)
        Next lc
    Next lr
    If Rng.Count = 1 Then
        CLEARLINE = avArr(1, 1)
    Else
        CLEARLINE = avArr
    End If
End Function

Public Function УБРЛИШН(ByVal Rng As Range, Optional dic_ As Variant, Optional check_doubles As Boolean) As Variant
    УБРЛИШН = CLEARLINE(Rng, dic_, check_doubles)
End Function
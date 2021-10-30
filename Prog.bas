Attribute VB_Name = "Prog"
Type Student
        idStr As String
        schoolStr As String
        classStr As String
        nameStr As String
        genderStr As String
        gradeStr As String
        heightStr As String
        weightStr As String
        fhlStr As String
        m50Str As String
        zwtStr As String
        tsStr As String
        tyStr As String
        ywqz0Str As String
        ywqz1Str As String
        nlp0Str As String
        nlp1Str As String

        gender As Long
        grade As Long
        ' 10倍
        height As Long
        ' 1000000倍
        weight As Long

        bmiValid As Long
        bmiScore As Long
        ' 10倍
        bmi As Long
        bmiLow As Long

        fhlValid As Long
        fhl As Long
        fhlScore As Long

        m50Valid As Long
        ' * 10
        m50 As Long
        m50Score As Long

        zwtValid As Long
        ' 10倍
        zwt As Long
        zwtScore As Long

        tsValid As Long
        ts As Long
        tsScore As Long
        tsJfScore As Long

        tyValid As Long
        ty As Long
        tyScore As Long
        tyJfScore As Long

        ywqzValid As Long
        ywqz0 As Long
        ywqz1 As Long
        ywqzScore As Long
        ywqzJfScore As Long

        nlpValid As Long
        ' 秒
        nlp0 As Long
        nlp1 As Long
        nlpScore As Long
        nlpJfScore As Long

        totalValid As Long
        ' 100倍
        totalScore As Long
        ' 100倍
        totalJfScore As Long

        ' 存储报表数据
        arr() As Variant
End Type

Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

' table header
Private rptHdrTbl
Private gradeNameTbl

Private BMIData
Private fhlData0
Private fhlData1
Private M50Data0
Private M50Data1
Private tsData0
Private tsData1
Private tyData0
Private tyData1
Private zwtData0
Private zwtData1
Private ywqzData0
Private ywqzData1
Private nlpData0
Private nlpData1

Sub start()
        Dim st As Student
        Dim row As Long
        Dim offset As Long
        Dim t0 As Long
        Dim t1 As Long
        Dim ret As Long
        Dim dbg As Long

        dbg = 0

        row = ActiveCell.row()
        offset = 0
        If dbg = 0 Then
                row = 1
                offset = 1 - ActiveCell.row()
        End If

        ReDim st.arr(64)

        initScoreTable
        initStringTable

        t0 = timeGetTime()
        Debug.Print "[" & t0 & "]: "; "创建表头 ..."

        createStudentReportHeader ActiveCell.Column()
        'Exit Sub

        Debug.Print "[" & timeGetTime() & "]: "; "开始生成 ..."

        Application.ScreenUpdating = False
        Do While True
                ' end of data
                If Range("A" & row) = "" Then Exit Do

                ' invalid grade
                If Not Range("B" & row) < 100 Then GoTo rowComplete
                If Not Range("B" & row) > 10 Then GoTo rowComplete

                initStudent st

                ret = readStudent(st, row)
                If ret <> 0 Then GoTo rowComplete

                ret = calcStudentScore(st)
                If ret <> 0 Then
                        Debug.Print "failed to calculate for row " & row
                        GoTo rowComplete
                End If

                createStudentReport st, offset

                If dbg = 1 Then Exit Do
rowComplete:
                If row Mod 200 = 0 Then
                        't1 = timeGetTime()
                        'Debug.Print "[" & t1 & "]: "; "已经生成到" & row & "行, 耗时 " & t1 - t0 & " ms ..."
                End If
                row = row + 1
                offset = offset + 1
        Loop
        t1 = timeGetTime()
        Debug.Print "[" & t1 & "]: "; "已经生成到" & row & "行, 耗时 " & t1 - t0 & " ms ..."
        Debug.Print "[" & timeGetTime() & "]: "; "生成结束 ..."
        Debug.Print ""
        Application.ScreenUpdating = True
End Sub

Function initStudent(ByRef st As Student)
        st.idStr = ""
        st.schoolStr = ""
        st.classStr = ""
        st.nameStr = ""
        st.genderStr = ""
        st.gradeStr = ""
        st.heightStr = ""
        st.weightStr = ""
        st.fhlStr = ""
        st.m50Str = ""
        st.zwtStr = ""
        st.tsStr = ""
        st.tyStr = ""
        st.ywqz0Str = ""
        st.ywqz1Str = ""
        st.nlp0Str = ""
        st.nlp1Str = ""

        st.gender = 0
        st.grade = 0

        st.height = 0
        st.weight = 0
        st.bmiValid = 0
        st.bmi = 0
        st.bmiLow = 0
        st.bmiScore = 0

        st.fhlValid = 0
        st.fhl = 0
        st.fhlScore = 0

        st.m50Valid = 0
        st.m50 = 0
        st.m50Score = 0

        st.zwtValid = 0
        st.zwt = 0
        st.zwtScore = 0

        st.tsValid = 0
        st.ts = 0
        st.tsScore = 0
        st.tsJfScore = 0

        st.tyValid = 0
        st.ty = 0
        st.tyScore = 0
        st.tyJfScore = 0

        st.ywqzValid = 0
        st.ywqz0 = 0
        st.ywqz1 = 0
        st.ywqzScore = 0
        st.ywqzJfScore = 0

        st.nlpValid = 0
        st.nlp0 = 0
        st.nlp1 = 0
        st.nlpScore = 0
        st.nlpJfScore = 0

        st.totalValid = 0
        st.totalScore = 0
        st.totalJfScore = 0
End Function

Function timeToSeconds(t As String)
        Dim num As Long
        Dim tmp As Long
        Dim i As Long

        num = 0
        tmp = Len(t)
        For i = 1 To tmp
                If Not IsNumeric(Mid(t, i, 1)) Then Exit For
                num = num + 1
        Next i

        If num <> 0 Then
                timeToSeconds = Int(val(Left(t, num))) * 60 + Int(val(Mid(t, num + 2)))
        End If
End Function

Function validateInput(val As Long, str As String)
        If val <> 0 Or str = "0" Then validateInput = 1
End Function

Function readStudent(ByRef st As Student, ByVal row As Long)
        If Range("B" & row) < 20 Then
                st.idStr = Range("E" & row)
                st.schoolStr = Range("A" & row)
                st.classStr = Range("D" & row)
                st.nameStr = Range("G" & row)
                st.genderStr = Range("H" & row).Value
                st.gender = Range("H" & row)
                st.gradeStr = Range("B" & row)
                st.heightStr = Range("M" & row)
                st.weightStr = Range("N" & row)
                st.fhlStr = Range("O" & row)
                st.m50Str = Range("P" & row)
                st.zwtStr = Range("Q" & row)
                st.tsStr = Range("R" & row)
                st.tyStr = Range("R" & row)
                st.ywqz0Str = Range("S" & row)
                st.ywqz1Str = Range("S" & row)
                st.nlp0Str = Range("T" & row)
                st.nlp1Str = Range("T" & row)
        Else
                st.idStr = Range("E" & row)
                st.schoolStr = Range("A" & row)
                st.classStr = Range("D" & row)
                st.nameStr = Range("G" & row)
                st.genderStr = Range("H" & row)
                st.gender = Range("H" & row)
                st.gradeStr = Range("B" & row)
                st.heightStr = Range("M" & row)
                st.weightStr = Range("N" & row)
                st.fhlStr = Range("O" & row)
                st.m50Str = Range("P" & row)
                st.zwtStr = Range("Q" & row)
                ' ty
                st.tsStr = Range("R" & row)
                st.tyStr = Range("R" & row)
                If Range("U" & row) <> "" Then
                        st.ywqz0Str = Range("U" & row)
                        'st.ywqz1Str = Range("u" & row)
                Else
                        st.ywqz0Str = Range("v" & row)
                        'st.ywqz1Str = Range("V" & row)
                End If
                If Range("S" & row) <> "" Then
                        st.nlp0Str = Range("S" & row)
                        'st.nlp1Str = Range("S" & row)
                Else
                        st.nlp0Str = Range("T" & row)
                        'st.nlp1Str = Range("T" & row)
                End If
        End If

        'st.gender = Int(Val(st.genderStr))
        st.grade = calcGradeIdx(val(st.gradeStr))

        st.height = stringToInt(st.heightStr, 1)
        st.weight = stringToInt(st.weightStr, 6)
        If st.height > 0 And st.weight > 0 Then st.bmiValid = 1

        st.fhl = Int(val(st.fhlStr))
        If st.fhl > 0 Then st.fhlValid = 1

        st.m50 = stringToInt(st.m50Str, 1)
        If st.m50 > 0 Then st.m50Valid = 1

        st.zwt = stringToInt(st.zwtStr, 1)
        st.zwtValid = validateInput(st.zwt, st.zwtStr)

        st.ts = Int(val(st.tsStr))
        If st.ts >= 0 Then st.tsValid = validateInput(st.ts, st.tsStr)
        st.ty = st.ts
        st.tyValid = st.tsValid

        st.ywqz0 = Int(val(st.ywqz0Str))
        st.ywqz1 = st.ywqz0
        If st.ywqz0 >= 0 Then st.ywqzValid = validateInput(st.ywqz0, st.ywqz0Str)

        st.nlp0 = timeToSeconds(st.nlp0Str)
        st.nlp1 = st.nlp0
        If st.nlp0 >= 0 Then st.nlpValid = validateInput(st.nlp0, st.nlp0Str)

        If st.bmiValid = 1 Or st.fhlValid = 1 Or st.m50Valid = 1 Or st.tsValid = 1 Or st.ywqzValid = 1 Or st.nlpValid = 1 Then
                st.totalValid = 1
        End If
End Function

Function calcStudentScore(ByRef st As Student)
        ' 共性指标
        If st.bmiValid = 1 Then calcBMIScore st
        If st.fhlValid = 1 Then calcFhlScore st
        If st.m50Valid = 1 Then calcM50Score st
        If st.zwtValid = 1 Then calcZwtScore st

        ' 跳绳或者跳远
        If st.grade < 6 Then
                If st.tsValid = 1 Then calcTsScore st
        Else
                If st.tyValid = 1 Then calcTyScore st
        End If

        ' 引体向上或者仰卧起坐
        If st.grade > 1 Then
                If st.ywqzValid = 1 Then calcYwqzScore st
        End If

        ' 耐力跑: 50x8或者1000米或者800米
        If st.grade > 3 Then
                If st.nlpValid = 1 Then calcNlpScore st
        End If

        If st.totalValid = 1 Then calcTotalScore st
End Function

Function createStudentReportHeader(ByVal col As Long)
        Dim cols As Long
        cols = UBound(rptHdrTbl)
        Cells(1, col).Resize(1, cols) = rptHdrTbl
End Function

Function createStudentReport(ByRef st As Student, ByVal row As Long)
        Dim offset As Long
        offset = row
        Dim col As Long
        col = 0
        Dim has As Long
        Dim i As Long

        With st
        ' 姓名
        .arr(col) = st.nameStr
        col = col + 1
        ' ID
        .arr(col) = st.idStr
        col = col + 1
        ' 学校
        .arr(col) = st.schoolStr
        col = col + 1
        ' 年级
        .arr(col) = getGradeName(st.grade)
        col = col + 1
        ' 班级
        .arr(col) = st.classStr
        col = col + 1
        ' 性别
        If st.gender = 1 Then
                .arr(col) = "男"
        Else
                .arr(col) = "女"
        End If
        col = col + 1

        If st.totalValid = 1 Then
                ' 综合成绩
                .arr(col + 0) = Round(Int((st.totalScore + st.totalJfScore + 5) / 10) / 10, 1)
                ' 综合评定
                .arr(col + 1) = getTotalLevel(st.totalScore + st.totalJfScore)
                ' 测试成绩
                .arr(col + 2) = Round(Int((st.totalScore + 5) / 10) / 10, 1)
                ' 测试评定
                .arr(col + 3) = getTotalLevel(st.totalScore)
                ' 附加分
                .arr(col + 4) = st.totalJfScore / 100
        Else
                .arr(col + 0) = "X"
                .arr(col + 1) = "X"
                .arr(col + 2) = "X"
                .arr(col + 3) = "X"
                .arr(col + 4) = "X"
        End If
        col = col + 5

        ' 身高成绩
        .arr(col) = st.heightStr
        col = col + 1
        ' 体重成绩
        .arr(col) = st.weightStr
        col = col + 1
        If st.bmiValid Then
                ' 身高体重指数
                .arr(col + 0) = st.bmi / 10
                ' 身高体重成绩
                .arr(col + 1) = st.bmiScore
                ' 身高体重等级
                .arr(col + 2) = getBmiLevel(st.bmiScore, st.bmiLow)
        Else
                .arr(col + 0) = "X"
                .arr(col + 1) = "X"
                .arr(col + 2) = "X"
        End If
        col = col + 3

        ' 肺活量成绩
        .arr(col) = st.fhlStr
        col = col + 1
        If st.fhlValid = 1 Then
                ' 肺活量得分
                .arr(col + 0) = st.fhlScore
                ' 肺活量等级
                .arr(col + 1) = getFhlLevel(st.fhlScore)
        Else
                .arr(col + 0) = "X"
                .arr(col + 1) = "X"
        End If
        col = col + 2

        ' 50米跑成绩
        .arr(col) = st.m50Str
        col = col + 1
        If st.m50Valid = 1 Then
                ' 50米跑得分
                .arr(col + 0) = st.m50Score
                ' 50米跑等级
                .arr(col + 1) = getM50Level(st.m50Score)
        Else
                .arr(col + 0) = "X"
                .arr(col + 1) = "X"
        End If
        col = col + 2

        ' 坐位体前屈成绩
        .arr(col) = st.zwtStr
        col = col + 1
        If st.zwtValid = 1 Then
                ' 坐位体前屈得分
                .arr(col + 0) = st.zwtScore
                ' 坐位体前屈等级
                .arr(col + 1) = getZwtLevel(st.zwtScore)
        Else
                .arr(col + 0) = "X"
                .arr(col + 1) = "X"
        End If
        col = col + 2

        If st.grade < 6 Then
                ' 一分钟跳绳成绩
                .arr(col + 0) = st.tsStr
                If st.tsValid = 1 Then
                        ' 一分钟跳绳得分
                        .arr(col + 1) = st.tsScore
                        ' 一分钟跳绳等级
                        .arr(col + 2) = getTsLevel(st.tsScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 一分钟跳绳成绩
                .arr(col + 0) = ""
                ' 一分钟跳绳得分
                .arr(col + 1) = ""
                ' 一分钟跳绳等级
                .arr(col + 2) = ""
        End If
        col = col + 3

        has = 0
        If st.grade >= 2 Then
                ' 小学二年级以上
                If st.grade < 6 Or st.gender <> 1 Then
                        has = 1
                End If
        End If

        If has = 1 Then
                ' 一分钟仰卧起坐成绩
                .arr(col + 0) = st.ywqz0Str
                If st.ywqzValid = 1 Then
                        ' 一分钟仰卧起坐得分
                        .arr(col + 1) = st.ywqzScore
                        ' 一分钟仰卧起坐等级
                        .arr(col + 2) = getYwqzLevel(st.ywqzScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 一分钟仰卧起坐成绩
                .arr(col + 0) = ""
                ' 一分钟仰卧起坐得分
                .arr(col + 1) = ""
                ' 一分钟仰卧起坐等级
                .arr(col + 2) = ""
        End If
        col = col + 3

        If st.grade = 4 Or st.grade = 5 Then
                ' 50米×8往返跑成绩
                .arr(col + 0) = st.nlp0Str
                If st.nlpValid = 1 Then
                        ' 50米×8往返跑得分
                        .arr(col + 1) = st.nlpScore
                        ' 50米×8往返跑等级
                        .arr(col + 2) = getNlpLevel(st.nlpScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 50米×8往返跑成绩
                .arr(col + 0) = ""
                ' 50米×8往返跑得分
                .arr(col + 1) = ""
                ' 50米×8往返跑等级
                .arr(col + 2) = ""
        End If
        col = col + 3

        If st.grade >= 6 Then
                ' 立定跳远成绩
                .arr(col + 0) = st.tyStr
                If st.tyValid = 1 Then
                        ' 立定跳远得分
                        .arr(col + 1) = st.tyScore
                        ' 立定跳远等级
                        .arr(col + 2) = getTyLevel(st.tyScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 立定跳远成绩
                .arr(col + 0) = ""
                ' 立定跳远得分
                .arr(col + 1) = ""
                ' 立定跳远等级
                .arr(col + 2) = ""
        End If
        col = col + 3

        If st.grade >= 6 And st.gender <> 1 Then
                ' 800米跑成绩
                .arr(col + 0) = st.nlp0Str
                If st.nlpValid = 1 Then
                        ' 800米跑得分
                        .arr(col + 1) = st.nlpScore
                        ' 800米跑等级
                        .arr(col + 2) = getNlpLevel(st.nlpScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 800米跑成绩
                .arr(col + 0) = ""
                ' 800米跑得分
                .arr(col + 1) = ""
                ' 800米跑等级
                .arr(col + 2) = ""
        End If
        col = col + 3

        If st.grade >= 6 And st.gender = 1 Then
                ' 1000米跑成绩
                .arr(col + 0) = st.nlp0Str
                If st.nlpValid = 1 Then
                        ' 1000米跑得分
                        .arr(col + 1) = st.nlpScore
                        ' 1000米跑等级
                        .arr(col + 2) = getNlpLevel(st.nlpScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 1000米跑成绩
                .arr(col + 0) = ""
                ' 1000米跑得分
                .arr(col + 1) = ""
                ' 1000米跑等级
                .arr(col + 2) = ""
        End If
        col = col + 3

        If st.grade >= 6 And st.gender = 1 Then
                ' 引体向上成绩
                .arr(col + 0) = st.ywqz0Str
                If st.ywqzValid = 1 Then
                        ' 引体向上得分
                        .arr(col + 1) = st.ywqzScore
                        ' 引体向上等级
                        .arr(col + 2) = getYwqzLevel(st.ywqzScore)
                Else
                        .arr(col + 1) = "X"
                        .arr(col + 2) = "X"
                End If
        Else
                ' 引体向上成绩
                .arr(col + 0) = ""
                ' 引体向上得分
                .arr(col + 1) = ""
                ' 引体向上等级
                .arr(col + 2) = ""
        End If
        col = col + 3
        End With

        ActiveCell.offset(offset, 0).Resize(1, col) = st.arr
End Function

Function getGradeName(ByVal grade As Long)
        Dim i, j
        i = LBound(gradeNameTbl)
        j = UBound(gradeNameTbl)
        If grade >= i And grade <= j Then
                getGradeName = gradeNameTbl(grade)
        Else
                getGradeName = "未知"
        End If
End Function

Function getTotalLevel(ByVal score As Long)
        If score >= 8995 Then
                getTotalLevel = "优秀"
        ElseIf score >= 7995 Then
                getTotalLevel = "良好"
        ElseIf score >= 5995 Then
                getTotalLevel = "及格"
        Else
                getTotalLevel = "不及格"
        End If
End Function

Function getBmiLevel(ByVal bmi As Long, ByVal low As Long)
        If bmi <= 60 Then
                getBmiLevel = "肥胖"
        ElseIf bmi <= 80 Then
                If low = 1 Then
                        getBmiLevel = "低体重"
                Else
                        getBmiLevel = "超重"
                End If
        Else
                        getBmiLevel = "正常"
        End If
End Function

Function getGeneralLevel(ByVal score As Long)
        If score >= 90 Then
                getGeneralLevel = "优秀"
        ElseIf score >= 80 Then
                getGeneralLevel = "良好"
        ElseIf score >= 60 Then
                getGeneralLevel = "及格"
        Else
                getGeneralLevel = "不及格"
        End If
End Function

Function getFhlLevel(ByVal fhl As Long)
        getFhlLevel = getGeneralLevel(fhl)
End Function

Function getM50Level(ByVal m50 As Long)
        getM50Level = getGeneralLevel(m50)
End Function

Function getZwtLevel(ByVal zwt As Long)
        getZwtLevel = getGeneralLevel(zwt)
End Function

Function getTsLevel(ByVal ts As Long)
        getTsLevel = getGeneralLevel(ts)
End Function

Function getTyLevel(ByVal ty As Long)
        getTyLevel = getGeneralLevel(ty)
End Function

Function getYwqzLevel(ByVal ywqz As Long)
        getYwqzLevel = getGeneralLevel(ywqz)
End Function

Function getNlpLevel(ByVal nlp As Long)
        getNlpLevel = getGeneralLevel(nlp)
End Function

' 0 1 2 3 4 5
' 6 7 8
' 9 10 11
' 12 13 14 15
Function calcGradeIdx(grade As Integer)
        If grade < 20 Then
                calcGradeIdx = grade - 11
        ElseIf grade < 30 Then
                calcGradeIdx = grade - 21 + 6
        ElseIf grade < 40 Then
                calcGradeIdx = grade - 31 + 9
        Else
                calcGradeIdx = grade - 41 + 12
        End If
End Function

Function stringToInt(s As String, rank As Integer)
        Dim v As Long
        Dim i As Integer
        Dim arr

        v = 0
        arr = Split(s, ".")
        If UBound(arr) < 0 Then
        ElseIf UBound(arr) < 1 Then
                For i = 1 To rank
                        arr(0) = arr(0) + "0"
                Next i
                v = val(arr(0))
        Else
                For i = 1 To rank
                        arr(1) = arr(1) + "0"
                Next i
                v = val(arr(0) & Mid(arr(1), 1, rank))
        End If

        stringToInt = v
End Function

Function doubleToInt(d As Double, rank As Integer)
        Dim s As String
        s = d
        doubleToInt = Int(stringToInt(s, rank))
End Function

Function calcBMI(ByRef st As Student)
        Dim h As Long
        Dim w As Long
        Dim tmp As Double
        h = st.height
        w = st.weight
        If h = 0 Then Exit Function
        tmp = w / (h * h)
        st.bmi = doubleToInt(tmp, 2)
        st.bmi = Int((st.bmi + 5) / 10)
End Function

Function calcBMIScore(ByRef st As Student)
        Dim idx As Integer

        ' 计算BMI
        calcBMI st

        idx = st.grade
        ' 大学计分规则一样
        If idx > 12 Then idx = 12
        ' 女生的分数在后面
        If st.gender <> 1 Then idx = idx + 13 * 4

        st.bmiLow = 0
        If st.bmi <= BMIData(idx + 13) Then
                st.bmiScore = 80
                ' 超重
                st.bmiLow = 1
        ElseIf st.bmi < BMIData(idx + 26) Then
                st.bmiScore = 100
        ElseIf st.bmi < BMIData(idx + 39) Then
                st.bmiScore = 80
        Else
                st.bmiScore = 60
        End If
End Function

Function calcFhlScoreImpl(ByRef st As Student, fhlData As Variant)
        ' 数据位置
        Dim offset As Integer
        Dim col As Integer
        Dim i As Integer

        col = st.grade
        ' 大1和大2一样
        If col = 13 Then col = 12
        ' 大3和大4一样
        If col > 13 Then col = 13

        ' 第一列是分数
        col = col + 1
        offset = col

        st.fhlScore = 0
        For i = 0 To 19
                If st.fhl >= fhlData(offset) Then
                        st.fhlScore = fhlData(offset - col)
                        Exit For
                End If
                offset = offset + 15
        Next i
End Function

Function calcFhlScore(ByRef st As Student)
        If st.gender = 1 Then
                calcFhlScoreImpl st, fhlData0
        Else
                calcFhlScoreImpl st, fhlData1
        End If
End Function

Function calcM50ScoreImpl(ByRef st As Student, m50Data As Variant)
        ' data position
        Dim offset As Long
        ' column
        Dim col As Long
        Dim i As Long

        col = st.grade
        If col = 13 Then col = 12
        If col > 13 Then col = 13

        ' 第一列是分数
        col = col + 1
        offset = col

        st.m50Score = 0
        For i = 0 To 19
                If st.m50 <= m50Data(offset) Then
                        st.m50Score = m50Data(offset - col)
                        Exit For
                End If
                offset = offset + 15
        Next i
End Function

Function calcM50Score(ByRef st As Student)
        If st.gender = 1 Then
                calcM50ScoreImpl st, M50Data0
        Else
                calcM50ScoreImpl st, M50Data1
        End If
End Function

Function calcZwtScoreImpl(ByRef st As Student, zwtData As Variant)
        ' data position
        Dim offset As Integer
        ' column
        Dim col As Integer
        Dim i As Integer

        col = st.grade
        If col = 13 Then col = 12
        If col > 13 Then col = 13

        col = col + 1
        offset = col

        st.zwtScore = 0
        For i = 0 To 19
                If st.zwt >= zwtData(offset) Then
                        st.zwtScore = zwtData(offset - col)
                        Exit For
                End If
                offset = offset + 15
        Next i
End Function

Function calcZwtScore(ByRef st As Student)
        If st.gender = 1 Then
                calcZwtScoreImpl st, zwtData0
        Else
                calcZwtScoreImpl st, zwtData1
        End If
End Function

Function calcTsScoreImpl(ByRef st As Student, tsData As Variant)
        ' data position
        Dim offset As Integer
        ' column
        Dim col As Integer
        Dim idx As Integer
        Dim i As Integer

        idx = st.grade
        If idx > 5 Then
                st.tsScore = 0
                GoTo found
        End If

        ' the scores are in the first column
        col = idx + 1
        ' 20 rows and 7 columns
        offset = col

        st.tsScore = 0
        For i = 0 To 19
                If st.ts >= tsData(offset) Then
                        st.tsScore = tsData(offset - col)
                        Exit For
                End If
                offset = offset + 7
        Next i

found:
        st.tsJfScore = 0
        If st.tsScore = 100 And idx < 6 Then
                Dim tmp As Long
                tmp = Int((st.ts - tsData(offset)) / 2)
                If tmp > 20 Then tmp = 20
                st.tsJfScore = tmp
        End If
End Function

Function calcTsScore(ByRef st As Student)
        If st.gender = 1 Then
                calcTsScoreImpl st, tsData0
        Else
                calcTsScoreImpl st, tsData1
        End If
End Function

Function calcTyScoreImpl(ByRef st As Student, tyData As Variant)
        ' data position
        Dim offset As Integer
        ' column
        Dim col As Integer
        Dim i As Integer

        col = st.grade
        If col < 6 Then
                st.tyScore = 0
                Exit Function
        ElseIf col < 12 Then
                col = col - 6
        ElseIf col < 14 Then
                col = 6
        Else
                col = 7
        End If

        col = col + 1
        offset = col

        st.tyScore = 0
        For i = 0 To 19
                If st.ty >= tyData(offset) Then
                        st.tyScore = tyData(offset - col)
                        Exit For
                End If
                offset = offset + 9
        Next i
End Function

Function calcTyScore(ByRef st As Student)
        If st.gender = 1 Then
                calcTyScoreImpl st, tyData0
        Else
                calcTyScoreImpl st, tyData1
        End If
End Function

Function calcYwqzScoreImpl(ByRef st As Student, ywqzData As Variant)
        ' data position
        Dim offset As Integer
        ' column
        Dim col As Integer
        Dim idx As Integer
        Dim i As Integer

        idx = st.grade
        If idx < 2 Then
                st.ywqzScore = 0
                GoTo found
        ElseIf idx < 12 Then
                idx = idx - 2
        ElseIf idx < 14 Then
                idx = 10
        Else
                idx = 11
        End If

        col = idx + 1
        offset = col

        st.ywqzScore = 0
        For i = 0 To 19
                If st.ywqz0 >= ywqzData(offset) Then
                        st.ywqzScore = ywqzData(offset - col)
                        Exit For
                End If
                offset = offset + 13
        Next i

found:
        ' 计算附加分
        st.ywqzJfScore = 0
        If st.ywqzScore = 100 Then
                Dim tmp As Long
                tmp = st.ywqz0 - ywqzData(offset)
                If st.gender <> 1 Then
                        If tmp < 7 Then
                                tmp = Int(tmp / 2)
                        Else
                                tmp = tmp - 3
                        End If
                End If
                If st.grade < 6 Then tmp = 0
                If tmp > 10 Then tmp = 10
                st.ywqzJfScore = tmp
        End If
End Function

Function calcYwqzScore(ByRef st As Student)
        If st.gender = 1 Then
                calcYwqzScoreImpl st, ywqzData0
        Else
                calcYwqzScoreImpl st, ywqzData1
        End If
End Function

Function calcNlpScoreImpl(ByRef st As Student, nlpData As Variant)
        ' data position
        Dim offset As Integer
        ' column
        Dim col As Integer
        Dim idx As Integer
        Dim i As Integer

        idx = st.grade
        If idx < 4 Then
                st.nlpScore = 0
                GoTo found
        ElseIf idx < 12 Then
                idx = idx - 4
        ElseIf idx < 14 Then
                idx = 8
        Else
                idx = 9
        End If

        col = idx + 1
        offset = col

        st.nlpScore = 0
        For i = 0 To 19
                If st.nlp0 <= nlpData(offset) Then
                        st.nlpScore = nlpData(offset - col)
                        ' MsgBox "nlp " & i & " " & nlp & " " & idx & " " & calcNlpScoreImpl
                        Exit For
                End If
                offset = offset + 11
        Next i

found:
        st.nlpJfScore = 0
        If st.nlpScore = 100 Then
                Dim x As Long
                Dim y As Long
                'x = 100 * nlpData(offset)
                'y = 100 * st.nlp0
                x = nlpData(offset)
                y = st.nlp0
                'Debug.Print x, y
                Dim tmp As Long
                'tmp = (Int(x / 100) - Int(y / 100)) * 60
                'tmp = tmp + Int(x Mod 100 - y Mod 100)
                tmp = x - y
                If st.gender = 1 Then
                        If tmp < 23 Then
                                tmp = Int(tmp / 4)
                        Else
                                tmp = 6 + Int((tmp - 23) / 3)
                        End If
                Else
                        tmp = Int(tmp / 5)
                End If
                If st.grade < 6 Then tmp = 0
                If tmp > 10 Then tmp = 10
                st.nlpJfScore = tmp
                'Debug.Print tmp
        End If
End Function

Function calcNlpScore(ByRef st As Student)
        If st.gender = 1 Then
                calcNlpScoreImpl st, nlpData0
        Else
                calcNlpScoreImpl st, nlpData1
        End If
End Function

Function calcTotalScore(ByRef st As Student)
        Dim idx As Long
        idx = st.grade

        ' Debug.Print st.bmiScore, st.fhlScore, st.m50Score, st.zwtScore; st.tsScore, st.tyScore, st.nlpScore, st.nlpJfScore, st.ywqzScore, st.ywqzJfScore

        If idx < 2 Then
                st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 30 + st.tsScore * 20
        ElseIf idx < 4 Then
                st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 20 + st.tsScore * 20 + st.ywqzScore * 10
        ElseIf idx < 6 Then
                st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 10 + st.tsScore * 10 + st.ywqzScore * 20 + st.nlpScore * 10
        Else
                st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 10 + st.tyScore * 10 + st.ywqzScore * 10 + st.nlpScore * 20
        End If

        ' 附加分
        st.totalJfScore = (st.tsJfScore + st.ywqzJfScore + st.nlpJfScore) * 100
End Function

Function initStringTable()
        rptHdrTbl = Array("姓名", "ID", "学校", "年级", "班级", "性别", _
                "综合成绩", "综合评定", "测试成绩", "测试成绩评定", "附加分", _
                "身高成绩", "体重成绩", "身高体重指数", "身高体重成绩", "身高体重等级", _
                "肺活量成绩", "肺活量得分", "肺活量等级", _
                "50米跑成绩", "50米跑得分", "50米跑等级", _
                "坐位体前屈成绩", "坐位体前屈得分", "坐位体前屈等级", _
                "一分钟跳绳成绩", "一分钟跳绳得分", "一分钟跳绳等级", _
                "一分钟仰卧起坐成绩", "一分钟仰卧起坐得分", "一分钟仰卧起坐等级", _
                "50米×8往返跑成绩", "50米×8往返跑得分", "50米×8往返跑等级", _
                "立定跳远成绩", "立定跳远得分", "立定跳远等级", _
                "800米跑成绩", "800米跑得分", "800米跑等级", _
                "1000米跑成绩", "1000米跑得分", "1000米跑等级", _
                "引体向上成绩", "引体向上得分", "引体向上等级")

        gradeNameTbl = Array("一年级", "二年级", "三年级", "四年级", "五年级", "六年级", _
                "初一", "初二", "初三", _
                "高一", "高二", "高三", _
                "大一", "大二", "大三", "大四")
End Function

Function initScoreTable()
        BMIData = Array(135, 137, 139, 142, 144, 147, 155, 157, 158, 165, 168, 173, 179, _
              134, 136, 138, 141, 143, 146, 154, 156, 157, 164, 167, 172, 178, _
              182, 185, 195, 202, 215, 219, 222, 226, 229, 233, 238, 239, 240, _
              204, 205, 222, 227, 242, 246, 250, 253, 261, 264, 266, 274, 280, _
              133, 135, 136, 137, 138, 142, 148, 153, 160, 165, 169, 171, 172, _
              132, 134, 135, 136, 137, 141, 147, 152, 159, 164, 168, 170, 171, _
              174, 179, 187, 195, 206, 209, 218, 223, 227, 228, 233, 234, 240, _
              193, 203, 212, 221, 230, 237, 245, 249, 252, 253, 255, 258, 280)

        fhlData0 = Array(100, 1700, 2000, 2300, 2600, 2900, 3200, 3640, 3940, 4240, 4540, 4740, 4940, 5040, 5140, _
                95, 1600, 1900, 2200, 2500, 2800, 3100, 3520, 3820, 4120, 4420, 4620, 4820, 4920, 5020, _
                90, 1500, 1800, 2100, 2400, 2700, 3000, 3400, 3700, 4000, 4300, 4500, 4700, 4800, 4900, _
                85, 1400, 1650, 1900, 2150, 2450, 2750, 3150, 3450, 3750, 4050, 4250, 4450, 4550, 4650, _
                80, 1300, 1500, 1700, 1900, 2200, 2500, 2900, 3200, 3500, 3800, 4000, 4200, 4300, 4400, _
                78, 1240, 1430, 1620, 1820, 2110, 2400, 2780, 3080, 3380, 3680, 3880, 4080, 4180, 4280, _
                76, 1180, 1360, 1540, 1740, 2020, 2300, 2660, 2960, 3260, 3560, 3760, 3960, 4060, 4160, _
                74, 1120, 1290, 1460, 1660, 1930, 2200, 2540, 2840, 3140, 3440, 3640, 3840, 3940, 4040, _
                72, 1060, 1220, 1380, 1580, 1840, 2100, 2420, 2720, 3020, 3320, 3520, 3720, 3820, 3920, _
                70, 1000, 1150, 1300, 1500, 1750, 2000, 2300, 2600, 2900, 3200, 3400, 3600, 3700, 3800, _
                68, 940, 1080, 1220, 1420, 1660, 1900, 2180, 2480, 2780, 3080, 3280, 3480, 3580, 3680, _
                66, 880, 1010, 1140, 1340, 1570, 1800, 2060, 2360, 2660, 2960, 3160, 3360, 3460, 3560, _
                64, 820, 940, 1060, 1260, 1480, 1700, 1940, 2240, 2540, 2840, 3040, 3240, 3340, 3440, _
                62, 760, 870, 980, 1180, 1390, 1600, 1820, 2120, 2420, 2720, 2920, 3120, 3220, 3320, _
                60, 700, 800, 900, 1100, 1300, 1500, 1700, 2000, 2300, 2600, 2800, 3000, 3100, 3200, _
                50, 660, 750, 840, 1030, 1220, 1410, 1600, 1890, 2180, 2470, 2660, 2850, 2940, 3030, _
                40, 620, 700, 780, 960, 1140, 1320, 1500, 1780, 2060, 2340, 2520, 2700, 2780, 2860, _
                30, 580, 650, 720, 890, 1060, 1230, 1400, 1670, 1940, 2210, 2380, 2550, 2620, 2690, _
                20, 540, 600, 660, 820, 980, 1140, 1300, 1560, 1820, 2080, 2240, 2400, 2460, 2520, _
                10, 500, 550, 600, 750, 900, 1050, 1200, 1450, 1700, 1950, 2100, 2250, 2300, 2350)
        fhlData1 = Array(100, 1400, 1600, 1800, 2000, 2250, 2500, 2750, 2900, 3050, 3150, 3250, 3350, 3400, 3450, _
                95, 1300, 1500, 1700, 1900, 2150, 2400, 2650, 2850, 3000, 3100, 3200, 3300, 3350, 3400, _
                90, 1200, 1400, 1600, 1800, 2050, 2300, 2550, 2800, 2950, 3050, 3150, 3250, 3300, 3350, _
                85, 1100, 1300, 1500, 1700, 1950, 2200, 2450, 2650, 2800, 2900, 3000, 3100, 3150, 3200, _
                80, 1000, 1200, 1400, 1600, 1850, 2100, 2350, 2500, 2650, 2750, 2850, 2950, 3000, 3050, _
                78, 960, 1150, 1340, 1530, 1770, 2010, 2250, 2400, 2550, 2650, 2750, 2850, 2900, 2950, _
                76, 920, 1100, 1280, 1460, 1690, 1920, 2150, 2300, 2450, 2550, 2650, 2750, 2800, 2850, _
                74, 880, 1050, 1220, 1390, 1610, 1830, 2050, 2200, 2350, 2450, 2550, 2650, 2700, 2750, _
                72, 840, 1000, 1160, 1320, 1530, 1740, 1950, 2100, 2250, 2350, 2450, 2550, 2600, 2650, _
                70, 800, 950, 1100, 1250, 1450, 1650, 1850, 2000, 2150, 2250, 2350, 2450, 2500, 2550, _
                68, 760, 900, 1040, 1180, 1370, 1560, 1750, 1900, 2050, 2150, 2250, 2350, 2400, 2450, _
                66, 720, 850, 980, 1110, 1290, 1470, 1650, 1800, 1950, 2050, 2150, 2250, 2300, 2350, _
                64, 680, 800, 920, 1040, 1210, 1380, 1550, 1700, 1850, 1950, 2050, 2150, 2200, 2250, _
                62, 640, 750, 860, 970, 1130, 1290, 1450, 1600, 1750, 1850, 1950, 2050, 2100, 2150, _
                60, 600, 700, 800, 900, 1050, 1200, 1350, 1500, 1650, 1750, 1850, 1950, 2000, 2050, _
                50, 580, 680, 780, 880, 1020, 1170, 1310, 1460, 1610, 1710, 1810, 1910, 1960, 2010, _
                40, 560, 660, 760, 860, 990, 1140, 1270, 1420, 1570, 1670, 1770, 1870, 1920, 1970, _
                30, 540, 640, 740, 840, 960, 1110, 1230, 1380, 1530, 1630, 1730, 1830, 1880, 1930, _
                20, 520, 620, 720, 820, 930, 1080, 1190, 1340, 1490, 1590, 1690, 1790, 1840, 1890, _
                10, 500, 600, 700, 800, 900, 1050, 1150, 1300, 1450, 1550, 1650, 1750, 1800, 1850)

        M50Data0 = Array(100, 102, 96, 91, 87, 84, 82, 78, 75, 73, 71, 70, 68, 67, 66, _
                95, 103, 97, 92, 88, 85, 83, 79, 76, 74, 72, 71, 69, 68, 67, _
                90, 104, 98, 93, 89, 86, 84, 80, 77, 75, 73, 72, 70, 69, 68, _
                85, 105, 99, 94, 90, 87, 85, 81, 78, 76, 74, 73, 71, 70, 69, _
                80, 106, 100, 95, 91, 88, 86, 82, 79, 77, 75, 74, 72, 71, 70, _
                78, 108, 102, 97, 93, 90, 88, 84, 81, 79, 77, 76, 74, 73, 72, _
                76, 110, 104, 99, 95, 92, 90, 86, 83, 81, 79, 78, 76, 75, 74, _
                74, 112, 106, 101, 97, 94, 92, 88, 85, 83, 81, 80, 78, 77, 76, _
                72, 114, 108, 103, 99, 96, 94, 90, 87, 85, 83, 82, 80, 79, 78, _
                70, 116, 110, 105, 101, 98, 96, 92, 89, 87, 85, 84, 82, 81, 80, _
                68, 118, 112, 107, 103, 100, 98, 94, 91, 89, 87, 86, 84, 83, 82, _
                66, 120, 114, 109, 105, 102, 100, 96, 93, 91, 89, 88, 86, 85, 84, _
                64, 122, 116, 111, 107, 104, 102, 98, 95, 93, 91, 90, 88, 87, 86, _
                62, 124, 118, 113, 109, 106, 104, 100, 97, 95, 93, 92, 90, 89, 88, _
                60, 126, 120, 115, 111, 108, 106, 102, 99, 97, 95, 94, 92, 91, 90, _
                50, 128, 122, 117, 113, 110, 108, 104, 101, 99, 97, 96, 94, 93, 92, _
                40, 130, 124, 119, 115, 112, 110, 106, 103, 101, 99, 98, 96, 95, 94, _
                30, 132, 126, 121, 117, 114, 112, 108, 105, 103, 101, 100, 98, 97, 96, _
                20, 134, 128, 123, 119, 116, 114, 110, 107, 105, 103, 102, 100, 99, 98, _
                10, 136, 130, 125, 121, 118, 116, 112, 109, 107, 105, 104, 102, 101, 100)
        M50Data1 = Array(100, 110, 100, 92, 87, 83, 82, 81, 80, 79, 78, 77, 76, 75, 74, _
                95, 111, 101, 93, 88, 84, 83, 82, 81, 80, 79, 78, 77, 76, 75, _
                90, 112, 102, 94, 89, 85, 84, 83, 82, 81, 80, 79, 78, 77, 76, _
                85, 115, 105, 97, 92, 88, 87, 86, 85, 84, 83, 82, 81, 80, 79, _
                80, 118, 108, 100, 95, 91, 90, 89, 88, 87, 86, 85, 84, 83, 82, _
                78, 120, 110, 102, 97, 93, 92, 91, 90, 89, 88, 87, 86, 85, 84, _
                76, 122, 112, 104, 99, 95, 94, 93, 92, 91, 90, 89, 88, 87, 86, _
                74, 124, 114, 106, 101, 97, 96, 95, 94, 93, 92, 91, 90, 89, 88, _
                72, 126, 116, 108, 103, 99, 98, 97, 96, 95, 94, 93, 92, 91, 90, _
                70, 128, 118, 110, 105, 101, 100, 99, 98, 97, 96, 95, 94, 93, 92, _
                68, 130, 120, 112, 107, 103, 102, 101, 100, 99, 98, 97, 96, 95, 94, _
                66, 132, 122, 114, 109, 105, 104, 103, 102, 101, 100, 99, 98, 97, 96, _
                64, 134, 124, 116, 111, 107, 106, 105, 104, 103, 102, 101, 100, 99, 98, _
                62, 136, 126, 118, 113, 109, 108, 107, 106, 105, 104, 103, 102, 101, 100, _
                60, 138, 128, 120, 115, 111, 110, 109, 108, 107, 106, 105, 104, 103, 102, _
                50, 140, 130, 122, 117, 113, 112, 111, 110, 109, 108, 107, 106, 105, 104, _
                40, 142, 132, 124, 119, 115, 114, 113, 112, 111, 110, 109, 108, 107, 106, _
                30, 144, 134, 126, 121, 117, 116, 115, 114, 113, 112, 111, 110, 109, 108, _
                20, 146, 136, 128, 123, 119, 118, 117, 116, 115, 114, 113, 112, 111, 110, _
                10, 148, 138, 130, 125, 121, 120, 119, 118, 117, 116, 115, 114, 113, 112)

        tsData0 = Array(100, 109, 117, 126, 137, 148, 157, _
                95, 104, 112, 121, 132, 143, 152, _
                90, 99, 107, 116, 127, 138, 147, _
                85, 93, 101, 110, 121, 132, 141, _
                80, 87, 95, 104, 115, 126, 135, _
                78, 80, 88, 97, 108, 119, 128, _
                76, 73, 81, 90, 101, 112, 121, _
                74, 66, 74, 83, 94, 105, 114, _
                72, 59, 67, 76, 87, 98, 107, _
                70, 52, 60, 69, 80, 91, 100, _
                68, 45, 53, 62, 73, 84, 93, _
                66, 38, 46, 55, 66, 77, 86, _
                64, 31, 39, 48, 59, 70, 79, _
                62, 24, 32, 41, 52, 63, 72, _
                60, 17, 25, 34, 45, 56, 65, _
                50, 14, 22, 31, 42, 53, 62, _
                40, 11, 19, 28, 39, 50, 59, _
                30, 8, 16, 25, 36, 47, 56, _
                20, 5, 13, 22, 33, 44, 53, _
                10, 2, 10, 19, 30, 41, 50)
        tsData1 = Array(100, 117, 127, 139, 149, 158, 166, _
                95, 110, 120, 132, 142, 151, 159, _
                90, 103, 113, 125, 135, 144, 152, _
                85, 95, 105, 117, 127, 136, 144, _
                80, 87, 97, 109, 119, 128, 136, _
                78, 80, 90, 102, 112, 121, 129, _
                76, 73, 83, 95, 105, 114, 122, _
                74, 66, 76, 88, 98, 107, 115, _
                72, 59, 69, 81, 91, 100, 108, _
                70, 52, 62, 74, 84, 93, 101, _
                68, 45, 55, 67, 77, 86, 94, _
                66, 38, 48, 60, 70, 79, 87, _
                64, 31, 41, 53, 63, 72, 80, _
                62, 24, 34, 46, 56, 65, 73, _
                60, 17, 27, 39, 49, 58, 66, _
                50, 14, 24, 36, 46, 55, 63, _
                40, 11, 21, 33, 43, 52, 60, _
                30, 8, 18, 30, 40, 49, 57, _
                20, 5, 15, 27, 37, 46, 54, _
                10, 2, 12, 24, 34, 43, 51)

        tyData0 = Array(100, 225, 240, 250, 260, 265, 270, 273, 275, _
                95, 218, 233, 245, 255, 260, 265, 268, 270, _
                90, 211, 226, 240, 250, 255, 260, 263, 265, _
                85, 203, 218, 233, 243, 248, 253, 256, 258, _
                80, 195, 210, 225, 235, 240, 245, 248, 250, _
                78, 191, 206, 221, 231, 236, 241, 244, 246, _
                76, 187, 202, 217, 227, 232, 237, 240, 242, _
                74, 183, 198, 213, 223, 228, 233, 236, 238, _
                72, 179, 194, 209, 219, 224, 229, 232, 234, _
                70, 175, 190, 205, 215, 220, 225, 228, 230, _
                68, 171, 186, 201, 211, 216, 221, 224, 226, _
                66, 167, 182, 197, 207, 212, 217, 220, 222, _
                64, 163, 178, 193, 203, 208, 213, 216, 218, _
                62, 159, 174, 189, 199, 204, 209, 212, 214, _
                60, 155, 170, 185, 195, 200, 205, 208, 210, _
                50, 150, 165, 180, 190, 195, 200, 203, 205, _
                40, 145, 160, 175, 185, 190, 195, 198, 200, _
                30, 140, 155, 170, 180, 185, 190, 193, 195, _
                20, 135, 150, 165, 175, 180, 185, 188, 190, _
                10, 130, 145, 160, 170, 175, 180, 183, 185)
        tyData1 = Array(100, 196, 200, 202, 204, 205, 206, 207, 208, _
                95, 190, 194, 196, 198, 199, 200, 201, 202, _
                90, 184, 188, 190, 192, 193, 194, 195, 196, _
                85, 177, 181, 183, 185, 186, 187, 188, 189, _
                80, 170, 174, 176, 178, 179, 180, 181, 182, _
                78, 167, 171, 173, 175, 176, 177, 178, 179, _
                76, 164, 168, 170, 172, 173, 174, 175, 176, _
                74, 161, 165, 167, 169, 170, 171, 172, 173, _
                72, 158, 162, 164, 166, 167, 168, 169, 170, _
                70, 155, 159, 161, 163, 164, 165, 166, 167, _
                68, 152, 156, 158, 160, 161, 162, 163, 164, _
                66, 149, 153, 155, 157, 158, 159, 160, 161, _
                64, 146, 150, 152, 154, 155, 156, 157, 158, _
                62, 143, 147, 149, 151, 152, 153, 154, 155, _
                60, 140, 144, 146, 148, 149, 150, 151, 152, _
                50, 135, 139, 141, 143, 144, 145, 146, 147, _
                40, 130, 134, 136, 138, 139, 140, 141, 142, _
                30, 125, 129, 131, 133, 134, 135, 136, 137, _
                20, 120, 124, 126, 128, 129, 130, 131, 132, _
                10, 115, 119, 121, 123, 124, 125, 126, 127)

        zwtData0 = Array(100, 161, 162, 163, 164, 165, 166, 176, 196, 216, 236, 243, 246, 249, 251, _
                95, 146, 147, 149, 150, 152, 153, 159, 177, 197, 215, 224, 228, 231, 233, _
                90, 130, 132, 134, 136, 138, 140, 142, 158, 178, 194, 205, 210, 213, 215, _
                85, 120, 119, 118, 117, 116, 115, 123, 137, 158, 172, 183, 191, 195, 199, _
                80, 110, 106, 102, 98, 94, 90, 104, 116, 138, 150, 161, 172, 177, 182, _
                78, 99, 95, 91, 86, 82, 77, 91, 103, 124, 136, 147, 158, 163, 168, _
                76, 88, 84, 80, 74, 70, 64, 78, 90, 110, 122, 133, 144, 149, 154, _
                74, 77, 73, 69, 62, 58, 51, 65, 77, 96, 108, 119, 130, 135, 140, _
                72, 66, 62, 58, 50, 46, 38, 52, 64, 82, 94, 105, 116, 121, 126, _
                70, 55, 51, 47, 38, 34, 25, 39, 51, 68, 80, 91, 102, 107, 112, _
                68, 44, 40, 36, 26, 22, 12, 26, 38, 54, 66, 77, 88, 93, 98, _
                66, 33, 29, 25, 14, 10, -1, 13, 25, 40, 52, 63, 74, 79, 84, _
                64, 22, 18, 14, 2, -2, -14, 0, 12, 26, 38, 49, 60, 65, 70, _
                62, 11, 7, 3, -10, -14, -27, -13, -1, 12, 24, 35, 46, 51, 56, _
                60, 0, -4, -8, -22, -26, -40, -26, -14, -2, 10, 21, 32, 37, 42, _
                50, -8, -12, -16, -32, -36, -50, -38, -26, -14, 0, 11, 22, 27, 32, _
                40, -16, -20, -24, -42, -46, -60, -50, -38, -26, -10, 1, 12, 17, 22, _
                30, -24, -28, -32, -52, -56, -70, -62, -50, -38, -20, -9, 2, 7, 12, _
                20, -32, -36, -40, -62, -66, -80, -74, -62, -50, -30, -19, -8, -3, 2, _
                10, -40, -44, -48, -72, -76, -90, -86, -74, -62, -40, -29, -18, -13, -8)
        zwtData1 = Array(100, 186, 189, 192, 195, 198, 199, 218, 227, 235, 242, 248, 253, 258, 263, _
                95, 173, 176, 179, 181, 185, 187, 201, 210, 218, 225, 231, 236, 240, 244, _
                90, 160, 163, 166, 169, 172, 175, 184, 193, 201, 208, 214, 219, 222, 224, _
                85, 147, 148, 149, 150, 151, 152, 167, 176, 184, 191, 197, 202, 206, 210, _
                80, 134, 133, 132, 131, 130, 129, 150, 159, 167, 174, 180, 185, 190, 195, _
                78, 123, 122, 121, 120, 119, 118, 137, 146, 154, 161, 167, 172, 177, 182, _
                76, 112, 111, 110, 109, 108, 107, 124, 133, 141, 148, 154, 159, 164, 169, _
                74, 101, 100, 99, 98, 97, 96, 111, 120, 128, 135, 141, 146, 151, 156, _
                72, 90, 89, 88, 87, 86, 85, 98, 107, 115, 122, 128, 133, 138, 143, _
                70, 79, 78, 77, 76, 75, 74, 85, 94, 102, 109, 115, 120, 125, 130, _
                68, 68, 67, 66, 65, 64, 63, 72, 81, 89, 96, 102, 107, 112, 117, _
                66, 57, 56, 55, 54, 53, 52, 59, 68, 76, 83, 89, 94, 99, 104, _
                64, 46, 45, 44, 43, 42, 41, 46, 55, 63, 70, 76, 81, 86, 91, _
                62, 35, 34, 33, 32, 31, 30, 33, 42, 50, 57, 63, 68, 73, 78, _
                60, 24, 23, 22, 21, 20, 19, 20, 29, 37, 44, 50, 55, 60, 65, _
                50, 16, 15, 14, 13, 12, 11, 12, 21, 29, 36, 42, 47, 52, 57, _
                40, 8, 7, 6, 5, 4, 3, 4, 13, 21, 28, 34, 39, 44, 49, _
                30, 0, -1, -2, -3, -4, -5, -4, 5, 13, 20, 26, 31, 36, 41, _
                20, -8, -9, -10, -11, -12, -13, -12, -3, 5, 12, 18, 23, 28, 33, _
                10, -16, -17, -18, -19, -20, -21, -20, -11, -3, 4, 10, 15, 20, 25)

        ywqzData0 = Array(100, 48, 49, 50, 51, 13, 14, 15, 16, 17, 18, 19, 20, _
                95, 45, 46, 47, 48, 12, 13, 14, 15, 16, 17, 18, 19, _
                90, 42, 43, 44, 45, 11, 12, 13, 14, 15, 16, 17, 18, _
                85, 39, 40, 41, 42, 10, 11, 12, 13, 14, 15, 16, 17, _
                80, 36, 37, 38, 39, 9, 10, 11, 12, 13, 14, 15, 16, _
                78, 34, 35, 36, 37, 9, 10, 11, 12, 13, 14, 15, 16, _
                76, 32, 33, 34, 35, 8, 9, 10, 11, 12, 13, 14, 15, _
                74, 30, 31, 32, 33, 8, 9, 10, 11, 12, 13, 14, 15, _
                72, 28, 29, 30, 31, 7, 8, 9, 10, 11, 12, 13, 14, _
                70, 26, 27, 28, 29, 7, 8, 9, 10, 11, 12, 13, 14, _
                68, 24, 25, 26, 27, 6, 7, 8, 9, 10, 11, 12, 13, _
                66, 22, 23, 24, 25, 6, 7, 8, 9, 10, 11, 12, 13, _
                64, 20, 21, 22, 23, 5, 6, 7, 8, 9, 10, 11, 12, _
                62, 18, 19, 20, 21, 5, 6, 7, 8, 9, 10, 11, 12, _
                60, 16, 17, 18, 19, 4, 5, 6, 7, 8, 9, 10, 11, _
                50, 14, 15, 16, 17, 3, 4, 5, 6, 7, 8, 9, 10, _
                40, 12, 13, 14, 15, 2, 3, 4, 5, 6, 7, 8, 9, _
                30, 10, 11, 12, 13, 1, 2, 3, 4, 5, 6, 7, 8, _
                20, 8, 9, 10, 11, 1, 1, 2, 3, 4, 5, 6, 7, _
                10, 6, 7, 8, 9, 1, 1, 1, 2, 3, 4, 5, 6)
        ywqzData1 = Array(100, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, _
                95, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, _
                90, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, _
                85, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, _
                80, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, _
                78, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, _
                76, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, _
                74, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, _
                72, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, _
                70, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, _
                68, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, _
                66, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, _
                64, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, _
                62, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, _
                60, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, _
                50, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, _
                40, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, _
                30, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, _
                20, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, _
                10, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17)

        nlpData0 = Array(100, 96, 90, 235, 230, 220, 210, 205, 200, 197, 195, _
                95, 99, 93, 245, 235, 225, 215, 210, 205, 202, 200, _
                90, 102, 96, 255, 240, 230, 220, 215, 210, 207, 205, _
                85, 105, 99, 262, 247, 237, 227, 222, 217, 214, 212, _
                80, 108, 102, 270, 255, 245, 235, 230, 225, 222, 220, _
                78, 111, 105, 275, 260, 250, 240, 235, 230, 227, 225, _
                76, 114, 108, 280, 265, 255, 245, 240, 235, 232, 230, _
                74, 117, 111, 285, 270, 260, 250, 245, 240, 237, 235, _
                72, 120, 114, 290, 275, 265, 255, 250, 245, 242, 240, _
                70, 123, 117, 295, 280, 270, 260, 255, 250, 247, 245, _
                68, 126, 120, 300, 285, 275, 265, 260, 255, 252, 250, _
                66, 129, 123, 305, 290, 280, 270, 265, 260, 257, 255, _
                64, 132, 126, 310, 295, 285, 275, 270, 265, 262, 260, _
                62, 135, 129, 315, 300, 290, 280, 275, 270, 267, 265, _
                60, 138, 132, 320, 305, 295, 285, 280, 275, 272, 270, _
                50, 142, 136, 340, 325, 315, 305, 300, 295, 292, 290, _
                40, 146, 140, 360, 345, 335, 325, 320, 315, 312, 310, _
                30, 150, 144, 380, 365, 355, 345, 340, 335, 332, 330, _
                20, 154, 148, 400, 385, 375, 365, 360, 355, 352, 350, _
                10, 158, 152, 420, 405, 395, 385, 380, 375, 372, 370)
        nlpData1 = Array(100, 101, 97, 215, 210, 205, 204, 202, 200, 198, 196, _
                95, 104, 100, 222, 217, 212, 210, 208, 206, 204, 202, _
                90, 107, 103, 229, 224, 219, 216, 214, 212, 210, 208, _
                85, 110, 106, 237, 232, 227, 223, 221, 219, 217, 215, _
                80, 113, 109, 245, 240, 235, 230, 228, 226, 224, 222, _
                78, 116, 112, 250, 245, 240, 235, 233, 231, 229, 227, _
                76, 119, 115, 255, 250, 245, 240, 238, 236, 234, 232, _
                74, 122, 118, 260, 255, 250, 245, 243, 241, 239, 237, _
                72, 125, 121, 265, 260, 255, 250, 248, 246, 244, 242, _
                70, 128, 124, 270, 265, 260, 255, 253, 251, 249, 247, _
                68, 131, 127, 275, 270, 265, 260, 258, 256, 254, 252, _
                66, 134, 130, 280, 275, 270, 265, 263, 261, 259, 257, _
                64, 137, 133, 285, 280, 275, 270, 268, 266, 264, 262, _
                62, 140, 136, 290, 285, 280, 275, 273, 271, 269, 267, _
                60, 143, 139, 295, 290, 285, 280, 278, 276, 274, 272, _
                50, 147, 143, 305, 300, 295, 290, 288, 286, 284, 282, _
                40, 151, 147, 315, 310, 305, 300, 298, 296, 294, 292, _
                30, 155, 151, 325, 320, 315, 310, 308, 306, 304, 302, _
                20, 159, 155, 335, 330, 325, 320, 318, 316, 314, 312, _
                10, 163, 159, 345, 340, 335, 330, 328, 326, 324, 322)
End Function

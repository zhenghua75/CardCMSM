VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "M1Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim akey(6) As Byte
Dim bkey(6) As Byte
Dim hexkey As String * 12
Dim displaydata(8) As Byte
Dim eepromdata As String * 32
Dim hardver As String * 20
Dim libver As String * 16
Dim ser As String * 16
Dim Snr As Long
Dim Rsnr As Long
Dim data16 As String * 16
Dim data32 As String * 32
Dim databuff32 As String * 32
Dim databuff16 As String * 16
Dim rvalue As Long
Dim wvalue As Long
Dim mlrvalue As Integer
Dim mlwvalue As Integer
Dim cardmode As Integer
Dim loadmode As Integer
Dim sector As Integer
Dim address As Integer
Dim ptrdest As String * 16
Dim ptrsource As String * 16
Dim time As String * 14
Dim timebuff As String * 14
Dim tagtype As Integer
Dim Size As Integer
Dim fIndex As Integer
Dim dot As String

Const port As Integer = 0 '端口
Const baud As Long = 9600 '115200 '通讯波特率9600～115200
Const dataGroup As String = "0000000000000000000000000" '00" '00000"

Const PROGRAMERR As String = "RF999"

Const RFINITERR As String = "RF001"
Const RFREQUESTERR As String = "RF002"
Const RFANTICOLLERR As String = "RF003"
Const RFSELECTERR As String = "RF004"
Const RFLOADKEY_A_ERR As String = "RF005"
Const RFLOADKEY_B_ERR As String = "RF006"
Const RFAUTHENTICATION_A_ERR As String = "RF007"
Const RFAUTHENTICATION_B_ERR As String = "RF008"
Const RFREADERR As String = "RF009"
Const RFWRITEERR As String = "RF010"
Const RFCHANGEB3ERR As String = "RF011"
Const RFWRITEINVERR As String = "RF012"
Const RFSETTIMEERR As String = "RF013"
Const RFSETDISPMODE As String = "RF014"
Const RFCARDIDLENGTH7 As String = "RF015"

Const RFREADCARDIDERR As String = "RF016"
Const RFREADCHARGEERR As String = "RF017"
Const RFREADIGERR As String = "RF018"

Const RFWRITECARDIDERR As String = "RF019"
Const RFWRITECHARGEERR As String = "RF020"
Const RFWRITEIGERR As String = "RF021"

Const RFCARDIDMATCH As String = "RF022"

Const RFCHECKWRITECARDIDERR As String = "RF023"
Const RFCHECKWRITECHARGEERR As String = "RF024"
Const RFCHECKWRITEIGERR As String = "RF025"

Const RFEMP As String = "RF026"

Const RFOK As String = "OPSUCCESS"
Dim ret As String
Dim dInCharge As Double
Dim dInIg As Long

Public Function SetDate(ByVal strDateWeekTime As String) As String
    'YYWWMMDDHHMMSS 年星期月日时分秒 每个2位09060321101010    09年星期6 3月21日 10点10分10秒
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        SetDate = RFINITERR '设备初始化端口失败
        Exit Function
    End If

    st = rf_settimehex(icdev, strDateWeekTime)
    If st <> 0 Then
        SetDate = RFSETTIMEERR '设置时间失败
        Exit Function
    End If

    st = rf_disp_mode(icdev, &H1)
    If st <> 0 Then
        SetDate = RFSETDISPMODE '设置显示模式失败
        Exit Function
    End If

    Call quit

    SetDate = RFOK

End Function

Private Function InitCard() As String
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        InitCard = RFINITERR '设备初始化端口失败
        Exit Function
    End If
    '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位
     If st <> 0 Then
        InitCard = RFREQUESTERR '寻卡请求失败
        Exit Function
    End If
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        InitCard = RFREQUESTERR '寻卡请求失败
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        InitCard = RFANTICOLLERR '卡防冲突，返回卡的序列号
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        InitCard = RFSELECTERR '从多个卡中选取一个给定序列号的卡
        Exit Function
    End If
    InitCard = RFOK
End Function
Private Function CheckA()
    '装载密码A
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        CheckA = RFLOADKEY_A_ERR '装载密码A出错
        Exit Function
     End If
    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        CheckA = RFAUTHENTICATION_A_ERR 'A密码验证错误
        Exit Function
    End If
    CheckA = RFOK
End Function
Private Function CheckB()
'装载B密码
    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          CheckB = RFLOADKEY_B_ERR '装载密码出错_B
          Exit Function
    End If
    '验证密码B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          CheckB = RFAUTHENTICATION_B_ERR '该卡片不属于本系统所使用
          Exit Function
    End If
    CheckB = RFOK
End Function
Private Function WriteCardNo(ByVal strCardNo As String) As String
    '写数据，卡号
    address = 4
    data32 = dataGroup & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        WriteCardNo = RFWRITECARDIDERR '写卡号失败"
        Exit Function
    End If
    'st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    'If st <> 0 Then
    '    WriteCardNo = RFWRITECARDIDERR '写卡号失败"
    '    Exit Function
    'End If
    WriteCardNo = RFOK
End Function
Private Function Write5Charge(ByVal dCharge As Double) As String
    Dim strCharge As String
    strCharge = Format(dCharge, "0.00")
    fIndex = InStr(strCharge, ".")
    Dim strInt As String
    strInt = Left(strCharge, fIndex - 1)
    dot = Right(strCharge, 2)

    address = 5
    Dim strF As String
    strF = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"
    data32 = strInt & Left(strF, Len(strF) - Len(strInt)) & dot
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        Write5Charge = RFWRITECHARGEERR '写余额失败"
        Exit Function
    End If
    'st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    ' If st <> 0 Then
    '    Write5Charge = RFWRITECHARGEERR '写余额失败"
    '    Exit Function
    'End If
    Write5Charge = RFOK
End Function
Private Function Write5Ig(ByVal lIg As Long) As String
    Dim strIg As String
    strIg = CStr(lIg)

    address = 6
    Dim strF As String
    strF = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"
    data32 = strIg & Left(strF, Len(strF) - Len(strIg))
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        Write5Ig = RFWRITEIGERR '写积分失败"
        Exit Function
    End If
    'st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    'If st <> 0 Then
    '    Write5Ig = RFWRITEIGERR '写积分失败"
    '    Exit Function
    'End If
    Write5Ig = RFOK
End Function
Private Function Write7Charge(ByVal dCharge As Double) As String
    '写余额 2位小数
    dCharge = dCharge * 100 '2 ^ 2 '左移2位
    wvalue = CLng(dCharge)
    address = 5
    st = rf_initval(ByVal icdev, address, wvalue)
    If st <> 0 Then
        Write7Charge = RFWRITECHARGEERR '写余额失败
        'Call quit
        Exit Function
    End If
    Write7Charge = RFOK
End Function
Private Function Write7Ig(ByVal lIg As Long) As String
    '写积分
    address = 6
    st = rf_initval(ByVal icdev, address, lIg)
    If st <> 0 Then
        Write7Ig = RFWRITEIGERR '写积分失败
        Exit Function
    End If
    Write7Ig = RFOK
End Function
Private Function Read5Charge(dCharge As Double) As String
    address = 5
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         Read5Charge = RFREADCHARGEERR '读取余额错误
         Exit Function
    End If
    fIndex = InStr(databuff32, "F")
    dot = Right(databuff32, 2)
    data32 = Left(databuff32, fIndex - 1) & "." & dot
    dCharge = CDbl(data32)
    Read5Charge = RFOK
End Function
Private Function Read5Ig(lIg As Long) As String
    address = 6
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         Read5Ig = RFREADIGERR '读取积分错误
         Exit Function
    End If
    fIndex = InStr(databuff32, "F")
    data32 = Left(databuff32, fIndex - 1)
    lIg = CLng(data32)
    Read5Ig = RFOK
End Function
Private Function Read7Charge(dCharge As Double) As String
    address = 5
    st = rf_readval(ByVal icdev, address, rvalue)
    If st <> 0 Then
        Read7Charge = RFREADCHARGEERR '读取余额错误
        Exit Function
    End If
    dCharge = CDbl(rvalue)
    dCharge = dCharge / 100
    Read7Charge = RFOK
End Function
Private Function Read7Ig(lIg As Long) As String
    address = 6
    st = rf_readval(ByVal icdev, address, rvalue)
    If st <> 0 Then
        Read7Ig = RFREADIGERR '读取积分错误
        Exit Function
    End If
    lIg = rvalue
    Read7Ig = RFOK
End Function
Public Function PutCard(ByVal strCardNo As String, ByVal dCharge As Double, ByVal dIg As Long) As String  '发卡
    Dim strCardRead As String

    If Len(strCardNo) <> 7 Then
        PutCard = RFCARDIDLENGTH7 '卡号需为7位
        Exit Function
    End If
    ret = InitCard()
    If ret <> RFOK Then
        PutCard = ret
        Call quit
        Exit Function
    End If
    '装载密码A
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        PutCard = RFLOADKEY_A_ERR '装载密码A出错
        Call quit
        Exit Function
     End If
    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = RFAUTHENTICATION_A_ERR 'A密码验证错误
        Call quit
        Exit Function
    End If
    ret = WriteCardNo(ByVal strCardNo)
    If ret <> RFOK Then
        PutCard = ret
        Call quit
        Exit Function
    End If
    ret = Write7Charge(ByVal dCharge)
    If ret <> RFOK Then
        PutCard = ret
        Call quit
        Exit Function
    End If
    ret = Write7Ig(ByVal dIg)
    If ret <> RFOK Then
        PutCard = ret
        Call quit
        Exit Function
    End If

    '改密码
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5

    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        PutCard = RFCHANGEB3ERR '改密码失败
        Call quit
        Exit Function
    End If

    st = rf_beep(icdev, 5)

    PutCard = RFOK
    '取消设备
    Call quit

End Function
Public Function WriteCharge(ByVal dCharge As Double, ByVal dChargeBak As Double) As String
    ret = InitCard()
    If ret <> RFOK Then
        WriteCharge = ret
        Call quit
        Exit Function
    End If

    ret = CheckB()
    If ret <> RFOK Then
        WriteCharge = ret
        Call quit
        Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         WriteCharge = RFREADCARDIDERR '读取卡号错误
         Call quit
         Exit Function
    End If

    Dim strInCardNo As String
    strInCardNo = Right(databuff32, 7)

    '写余额 2位小数
    If strInCardNo = "0000000" Then
        ret = Write5Charge(ByVal dCharge)
    Else
        ret = Write7Charge(ByVal dCharge)
    End If
    If ret <> RFOK Then
        WriteCharge = ret
        Call quit
        Exit Function
    End If
    '反读余额
    If strInCardNo = "0000000" Then
        ret = Read5Charge(dInCharge)
    Else
        ret = Read7Charge(dInCharge)
    End If
    If ret <> RFOK Then
        If strInCardNo = "0000000" Then
            ret = Write5Charge(ByVal dChargeBak)
        Else
            ret = Write7Charge(ByVal dChargeBak)
        End If
        If ret <> RFOK Then
            st = rf_beep(icdev, 10)
            WriteCharge = "CMT|反读反写出错，trans提交。correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCharge = "ROL|反读反写成功，积分未写，trans回滚。"
            Call quit
            Exit Function
        End If
    End If
    If dCharge <> dChargeBak Or dCharge <> dInCharge Or dChargeBak <> dInCharge Then
        '反写余额
        If strInCardNo = "0000000" Then
            ret = Write5Charge(ByVal dChargeBak)
        Else
            ret = Write7Charge(ByVal dChargeBak)
        End If
        If ret <> RFOK Then
            st = rf_beep(icdev, 10)
            WriteCharge = "CMT|校验错误，反写出错，trans提交。correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCharge = "ROL|校验错误，反写成功，trans回滚。"
            Call quit
            Exit Function
        End If
    End If

    st = rf_beep(icdev, 5)

    WriteCharge = RFOK
    '取消设备
    Call quit
End Function
Public Function WriteIg(ByVal lIg As Long) As String
     ret = InitCard()
    If ret <> RFOK Then
        WriteIg = ret
        Call quit
        Exit Function
    End If

    ret = CheckB()
    If ret <> RFOK Then
        WriteIg = ret
        Call quit
        Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         WriteIg = RFREADCARDIDERR '读取卡号错误
         Call quit
         Exit Function
    End If

    Dim strInCardNo As String
    strInCardNo = Right(databuff32, 7)

    '写积分
    If strInCardNo = "0000000" Then
        ret = Write5Ig(ByVal lIg)
    Else
        ret = Write7Ig(ByVal lIg)
    End If
    If ret <> RFOK Then
        WriteIg = ret
        Call quit
        Exit Function
    End If
    st = rf_beep(icdev, 5)
    WriteIg = RFOK
    '取消设备
    Call quit
End Function
Public Function WriteCard(ByVal dCharge As Double, ByVal dChargeBak As Double, ByVal dIg As Long) As String  '写卡
    Dim strCardRead As String

    ret = InitCard()
    If ret <> RFOK Then
        WriteCard = ret
        Call quit
        Exit Function
    End If

    ret = CheckB()
    If ret <> RFOK Then
        WriteCard = ret
        Call quit
        Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         WriteCard = RFREADCARDIDERR '读取卡号错误
         Call quit
         Exit Function
    End If

    Dim strInCardNo As String
    strInCardNo = Right(databuff32, 7)

    '写余额 2位小数
    If strInCardNo = "0000000" Then
        ret = Write5Charge(ByVal dCharge)
    Else
        ret = Write7Charge(ByVal dCharge)
    End If
    If ret <> RFOK Then
        WriteCard = ret
        Call quit
        Exit Function
    End If
    '写积分
    If strInCardNo = "0000000" Then
        ret = Write5Ig(ByVal dIg)
    Else
        ret = Write7Ig(ByVal dIg)
    End If
    If ret <> RFOK Then
        WriteCard = ret
        Call quit
        Exit Function
    End If

        '反读余额
    If strInCardNo = "0000000" Then
        ret = Read5Charge(dInCharge)
    Else
        ret = Read7Charge(dInCharge)
    End If
    If ret <> RFOK Then
        If strInCardNo = "0000000" Then
            ret = Write5Charge(ByVal dChargeBak)
        Else
            ret = Write7Charge(ByVal dChargeBak)
        End If
        If ret <> RFOK Then
            st = rf_beep(icdev, 10)
            WriteCard = "CMT|反读反写出错，trans提交。correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCard = "ROL|反读反写成功，积分未写，trans回滚。"
            Call quit
            Exit Function
        End If
    End If
    If dCharge <> dChargeBak Or dCharge <> dInCharge Or dChargeBak <> dInCharge Then
        '反写余额
        If strInCardNo = "0000000" Then
            ret = Write5Charge(ByVal dChargeBak)
        Else
            ret = Write7Charge(ByVal dChargeBak)
        End If
        If ret <> RFOK Then
            st = rf_beep(icdev, 10)
            WriteCard = "CMT|校验错误，反写出错，trans提交。correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCard = "ROL|校验错误，反写成功，trans回滚。"
            Call quit
            Exit Function
        End If
    End If


    st = rf_beep(icdev, 5)

    WriteCard = RFOK
    '取消设备
    Call quit


End Function


Public Function ReadCard(strCardNo As String, dCharge As Double, dIg As Long) As String   '读卡
     '初始化端口
    'On Error GoTo Err
    ret = InitCard()
    If ret <> RFOK Then
        ReadCard = ret
        Call quit
        Exit Function
    End If

    ret = CheckB()
    If ret <> RFOK Then
        ReadCard = ret
        Call quit
        Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         ReadCard = RFREADCARDIDERR '读取卡号错误
         Call quit
         Exit Function
    End If

    strCardNo = Right(databuff32, 7)
    Dim Is5 As Boolean
    Is5 = False

    If strCardNo = "0000000" Then
        Is5 = True
        strCardNo = Left(databuff32, 5)

    End If

    fIndex = InStr(strCardNo, "F")
    If fIndex > 0 Then
        ReadCard = RFEMP '为员工卡
        Call quit
        Exit Function
    End If

    If Is5 = True Then
        ret = Read5Charge(dCharge)
        If ret <> RFOK Then
            ReadCard = ret
            Call quit
            Exit Function
        End If
        ret = Read5Ig(dIg)
         If ret <> RFOK Then
            ReadCard = ret
            Call quit
            Exit Function
        End If
    Else
        ret = Read7Charge(dCharge)
        If ret <> RFOK Then
            ReadCard = ret
            Call quit
            Exit Function
        End If
        ret = Read7Ig(dIg)
         If ret <> RFOK Then
            ReadCard = ret
            Call quit
            Exit Function
        End If
    End If

    st = rf_beep(icdev, 5)

    ReadCard = RFOK
    '取消设备
    Call quit
    Exit Function
'Err:
'    Call quit
'    ReadCard = PROGRAMERR
End Function

'员工卡EMP0000
Public Function EmpPutCard(ByVal strCardNo As String) As String   '发卡
    Dim strCardRead As String

    If Len(strCardNo) <> 4 Then
        EmpPutCard = RFCARDIDLENGTH7 '卡号需为4位
        Exit Function
    End If
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        EmpPutCard = RFINITERR '设备初始化端口失败
        Exit Function
    End If

    '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位

    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        EmpPutCard = RFREQUESTERR '寻卡请求失败
    Call quit
    Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        EmpPutCard RFANTICOLLERR '卡防冲突，返回卡的序列号
        Call quit
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        EmpPutCard = RFSELECTERR '从多个卡中选取一个给定序列号的卡
        Call quit
        Exit Function
    End If



    '装载密码A
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    loadmode = 0
    sector = 1

    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        EmpPutCard = RFLOADKEY_A_ERR '装载密码A出错
        Call quit
        Exit Function
     End If

    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        EmpPutCard = RFAUTHENTICATION_A_ERR 'A密码验证错误
        Call quit
        Exit Function
    End If


    '写数据，卡号
    address = 4
    data32 = "000000000000000000000000" & "FFFF" & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        EmpPutCard = RFWRITEERR '写卡号失败"
        Call quit
        Exit Function
    End If


    '改密码

    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5

    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        EmpPutCard = RFCHANGEB3ERR '改密码失败
        Call quit
        Exit Function
    End If

    st = rf_beep(icdev, 5)

    EmpPutCard = RFOK
    '取消设备
    Call quit


End Function
'读取员工卡
Public Function EmpReadCard(strCardNo As String) As String    '读卡
     '初始化端口
    icdev = rf_init(port, baud)

    If icdev < 0 Then
        EmpReadCard = RFINITERR '设备初始化端口失败
        Exit Function
    End If

     '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位

    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        EmpReadCard = RFREQUESTERR '寻卡请求失败
        Call quit
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        EmpReadCard = RFANTICOLLERR '卡防冲突，返回卡的序列号
        Call quit
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        EmpReadCard = RFSELECTERR '从多个卡中选取一个给定序列号的卡
        Call quit
        Exit Function
    End If

    '装载B密码
    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          EmpReadCard = RFLOADKEY_B_ERR '装载密码出错_B
          Call quit
          Exit Function
    End If

    '验证密码B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          EmpReadCard = RFAUTHENTICATION_B_ERR '该卡片不属于本系统所使用
          Call quit
          Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         EmpReadCard = RFREADERR '读取卡号错误
         Call quit
         Exit Function
    End If

    strCardNo = Right(databuff32, 5)
    If strCardNo = "00000" Then
        strCardNo = Left(databuff32, 5) 'Mid(databuff32, 1, 5)
    End If

    st = rf_beep(icdev, 5)

    EmpReadCard = RFOK
    '取消设备
    Call quit

End Function

Public Function RecycleCard() As String    '卡回收
     '初始化端口
    icdev = rf_init(port, baud)

    If icdev < 0 Then
        RecycleCard = RFINITERR '设备初始化端口失败
        Exit Function
    End If

     '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位

    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        RecycleCard = RFREQUESTERR '寻卡请求失败
        Call quit
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        RecycleCard = RFANTICOLLERR '卡防冲突，返回卡的序列号
        Call quit
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        RecycleCard = RFSELECTERR '从多个卡中选取一个给定序列号的卡
        Call quit
        Exit Function
    End If

    '装载B密码
    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          RecycleCard = RFLOADKEY_B_ERR '装载密码出错_B
          Call quit
          Exit Function
    End If

    '验证密码B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          RecycleCard = RFAUTHENTICATION_B_ERR '该卡片不属于本系统所使用
          Call quit
          Exit Function
    End If


    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5
     st = rf_changeb3(icdev, sector, akey(0), 0, 0, 0, 1, 0, bkey(0))
     If st <> 0 Then
        RecycleCard = RFCHANGEB3ERR '改密码失败
        Call quit
        Exit Function
    End If

    st = rf_beep(icdev, 5)

    RecycleCard = RFOK
    '取消设备
    Call quit

End Function


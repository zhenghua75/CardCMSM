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

Const port As Integer = 0 '�˿�
Const baud As Long = 9600 '115200 'ͨѶ������9600��115200
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
    'YYWWMMDDHHMMSS ����������ʱ���� ÿ��2λ09060321101010    09������6 3��21�� 10��10��10��
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        SetDate = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If
    
    st = rf_settimehex(icdev, strDateWeekTime)
    If st <> 0 Then
        SetDate = RFSETTIMEERR '����ʱ��ʧ��
        Exit Function
    End If
    
    st = rf_disp_mode(icdev, &H1)
    If st <> 0 Then
        SetDate = RFSETDISPMODE '������ʾģʽʧ��
        Exit Function
    End If
    
    Call quit
    
    SetDate = RFOK
    
End Function

Private Function InitCard() As String
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        InitCard = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If
    'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
     If st <> 0 Then
        InitCard = RFREQUESTERR 'Ѱ������ʧ��
        Exit Function
    End If
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        InitCard = RFREQUESTERR 'Ѱ������ʧ��
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        InitCard = RFANTICOLLERR '������ͻ�����ؿ������к�
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        InitCard = RFSELECTERR '�Ӷ������ѡȡһ���������кŵĿ�
        Exit Function
    End If
    InitCard = RFOK
End Function
Private Function CheckA()
    'װ������A
    akey(0) = &HB6
    akey(1) = &H48
    akey(2) = &HA7
    akey(3) = &HF3
    akey(4) = &H2
    akey(5) = &H1C
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        CheckA = RFLOADKEY_A_ERR 'װ������A����
        Exit Function
     End If
    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        CheckA = RFAUTHENTICATION_A_ERR 'A������֤����
        Exit Function
    End If
    CheckA = RFOK
End Function
Private Function CheckB()
'װ��B����
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          CheckB = RFLOADKEY_B_ERR 'װ���������_B
          Exit Function
    End If
    '��֤����B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          CheckB = RFAUTHENTICATION_B_ERR '�ÿ�Ƭ�����ڱ�ϵͳ��ʹ��
          Exit Function
    End If
    CheckB = RFOK
End Function
Private Function WriteCardNo(ByVal strCardNo As String) As String
    'д���ݣ�����
    address = 4
    data32 = dataGroup & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        WriteCardNo = RFWRITECARDIDERR 'д����ʧ��"
        Exit Function
    End If
    'st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    'If st <> 0 Then
    '    WriteCardNo = RFWRITECARDIDERR 'д����ʧ��"
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
        Write5Charge = RFWRITECHARGEERR 'д���ʧ��"
        Exit Function
    End If
    'st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    ' If st <> 0 Then
    '    Write5Charge = RFWRITECHARGEERR 'д���ʧ��"
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
        Write5Ig = RFWRITEIGERR 'д����ʧ��"
        Exit Function
    End If
    'st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    'If st <> 0 Then
    '    Write5Ig = RFWRITEIGERR 'д����ʧ��"
    '    Exit Function
    'End If
    Write5Ig = RFOK
End Function
Private Function Write7Charge(ByVal dCharge As Double) As String
    'д��� 2λС��
    dCharge = dCharge * 100 '2 ^ 2 '����2λ
    wvalue = CLng(dCharge)
    address = 5
    st = rf_initval(ByVal icdev, address, wvalue)
    If st <> 0 Then
        Write7Charge = RFWRITECHARGEERR 'д���ʧ��
        'Call quit
        Exit Function
    End If
    Write7Charge = RFOK
End Function
Private Function Write7Ig(ByVal lIg As Long) As String
    'д����
    address = 6
    st = rf_initval(ByVal icdev, address, lIg)
    If st <> 0 Then
        Write7Ig = RFWRITEIGERR 'д����ʧ��
        Exit Function
    End If
    Write7Ig = RFOK
End Function
Private Function Read5Charge(dCharge As Double) As String
    address = 5
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         Read5Charge = RFREADCHARGEERR '��ȡ������
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
         Read5Ig = RFREADIGERR '��ȡ���ִ���
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
        Read7Charge = RFREADCHARGEERR '��ȡ������
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
        Read7Ig = RFREADIGERR '��ȡ���ִ���
        Exit Function
    End If
    lIg = rvalue
    Read7Ig = RFOK
End Function
Public Function PutCard(ByVal strCardNo As String, ByVal dCharge As Double, ByVal dIg As Long) As String  '����
    Dim strCardRead As String
    
    If Len(strCardNo) <> 7 Then
        PutCard = RFCARDIDLENGTH7 '������Ϊ7λ
        Exit Function
    End If
    ret = InitCard()
    If ret <> RFOK Then
        PutCard = ret
        Call quit
        Exit Function
    End If
    'װ������A
    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &HC6
    akey(3) = &H8C
    akey(4) = &HD9
    akey(5) = &HE5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        PutCard = RFLOADKEY_A_ERR 'װ������A����
        Call quit
        Exit Function
     End If
    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = RFAUTHENTICATION_A_ERR 'A������֤����
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

    '������
    akey(0) = &HB6
    akey(1) = &H48
    akey(2) = &HA7
    akey(3) = &HF3
    akey(4) = &H2
    akey(5) = &H1C
    
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    
    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        PutCard = RFCHANGEB3ERR '������ʧ��
        Call quit
        Exit Function
    End If
    
    st = rf_beep(icdev, 5)

    PutCard = RFOK
    'ȡ���豸
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
         WriteCharge = RFREADCARDIDERR '��ȡ���Ŵ���
         Call quit
         Exit Function
    End If
    
    Dim strInCardNo As String
    strInCardNo = Right(databuff32, 7)
    
    'д��� 2λС��
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
    '�������
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
            WriteCharge = "CMT|������д������trans�ύ��correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCharge = "ROL|������д�ɹ�������δд��trans�ع���"
            Call quit
            Exit Function
        End If
    End If
    If dCharge <> dChargeBak Or dCharge <> dInCharge Or dChargeBak <> dInCharge Then
        '��д���
        If strInCardNo = "0000000" Then
            ret = Write5Charge(ByVal dChargeBak)
        Else
            ret = Write7Charge(ByVal dChargeBak)
        End If
        If ret <> RFOK Then
            st = rf_beep(icdev, 10)
            WriteCharge = "CMT|У����󣬷�д������trans�ύ��correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCharge = "ROL|У����󣬷�д�ɹ���trans�ع���"
            Call quit
            Exit Function
        End If
    End If
    
    st = rf_beep(icdev, 5)

    WriteCharge = RFOK
    'ȡ���豸
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
         WriteIg = RFREADCARDIDERR '��ȡ���Ŵ���
         Call quit
         Exit Function
    End If
    
    Dim strInCardNo As String
    strInCardNo = Right(databuff32, 7)
    
    'д����
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
    'ȡ���豸
    Call quit
End Function
Public Function WriteCard(ByVal dCharge As Double, ByVal dChargeBak As Double, ByVal dIg As Long) As String  'д��
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
         WriteCard = RFREADCARDIDERR '��ȡ���Ŵ���
         Call quit
         Exit Function
    End If
    
    Dim strInCardNo As String
    strInCardNo = Right(databuff32, 7)
    
    'д��� 2λС��
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
    'д����
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
    
        '�������
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
            WriteCard = "CMT|������д������trans�ύ��correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCard = "ROL|������д�ɹ�������δд��trans�ع���"
            Call quit
            Exit Function
        End If
    End If
    If dCharge <> dChargeBak Or dCharge <> dInCharge Or dChargeBak <> dInCharge Then
        '��д���
        If strInCardNo = "0000000" Then
            ret = Write5Charge(ByVal dChargeBak)
        Else
            ret = Write7Charge(ByVal dChargeBak)
        End If
        If ret <> RFOK Then
            st = rf_beep(icdev, 10)
            WriteCard = "CMT|У����󣬷�д������trans�ύ��correct data: C:" & CStr(dCharge)
            Call quit
            Exit Function
        Else
            WriteCard = "ROL|У����󣬷�д�ɹ���trans�ع���"
            Call quit
            Exit Function
        End If
    End If


    st = rf_beep(icdev, 5)

    WriteCard = RFOK
    'ȡ���豸
    Call quit
        
       
End Function


Public Function ReadCard(strCardNo As String, dCharge As Double, dIg As Long) As String   '����
     '��ʼ���˿�
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
         ReadCard = RFREADCARDIDERR '��ȡ���Ŵ���
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
        ReadCard = RFEMP 'ΪԱ����
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
    'ȡ���豸
    Call quit
    Exit Function
'Err:
'    Call quit
'    ReadCard = PROGRAMERR
End Function

'Ա����EMP0000
Public Function EmpPutCard(ByVal strCardNo As String) As String   '����
    Dim strCardRead As String
    
    If Len(strCardNo) <> 4 Then
        EmpPutCard = RFCARDIDLENGTH7 '������Ϊ4λ
        Exit Function
    End If
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        EmpPutCard = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If
        
    'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
    
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        EmpPutCard = RFREQUESTERR 'Ѱ������ʧ��
    Call quit
    Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        EmpPutCard RFANTICOLLERR '������ͻ�����ؿ������к�
        Call quit
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        EmpPutCard = RFSELECTERR '�Ӷ������ѡȡһ���������кŵĿ�
        Call quit
        Exit Function
    End If
    
    
    
    'װ������A
    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &HC6
    akey(3) = &H8C
    akey(4) = &HD9
    akey(5) = &HE5
    
    loadmode = 0
    sector = 1
       
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        EmpPutCard = RFLOADKEY_A_ERR 'װ������A����
        Call quit
        Exit Function
     End If

    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        EmpPutCard = RFAUTHENTICATION_A_ERR 'A������֤����
        Call quit
        Exit Function
    End If
    

    'д���ݣ�����
    address = 4
    data32 = "000000000000000000000000" & "FFFF" & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        EmpPutCard = RFWRITEERR 'д����ʧ��"
        Call quit
        Exit Function
    End If
    
    
    '������
    
    akey(0) = &HB6
    akey(1) = &H48
    akey(2) = &HA7
    akey(3) = &HF3
    akey(4) = &H2
    akey(5) = &H1C
    
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    
    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        EmpPutCard = RFCHANGEB3ERR '������ʧ��
        Call quit
        Exit Function
    End If
    
    st = rf_beep(icdev, 5)

    EmpPutCard = RFOK
    'ȡ���豸
    Call quit
        
       
End Function
'��ȡԱ����
Public Function EmpReadCard(strCardNo As String) As String    '����
     '��ʼ���˿�
    icdev = rf_init(port, baud)

    If icdev < 0 Then
        EmpReadCard = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If

     'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
    
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        EmpReadCard = RFREQUESTERR 'Ѱ������ʧ��
        Call quit
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        EmpReadCard = RFANTICOLLERR '������ͻ�����ؿ������к�
        Call quit
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        EmpReadCard = RFSELECTERR '�Ӷ������ѡȡһ���������кŵĿ�
        Call quit
        Exit Function
    End If
    
    'װ��B����
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          EmpReadCard = RFLOADKEY_B_ERR 'װ���������_B
          Call quit
          Exit Function
    End If
    
    '��֤����B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          EmpReadCard = RFAUTHENTICATION_B_ERR '�ÿ�Ƭ�����ڱ�ϵͳ��ʹ��
          Call quit
          Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         EmpReadCard = RFREADERR '��ȡ���Ŵ���
         Call quit
         Exit Function
    End If
    
    strCardNo = Right(databuff32, 5)
    If strCardNo = "00000" Then
        strCardNo = Left(databuff32, 5) 'Mid(databuff32, 1, 5)
    End If
    
    st = rf_beep(icdev, 5)
    
    EmpReadCard = RFOK
    'ȡ���豸
    Call quit
    
End Function

Public Function RecycleCard() As String    '������
     '��ʼ���˿�
    icdev = rf_init(port, baud)

    If icdev < 0 Then
        RecycleCard = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If

     'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
    
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        RecycleCard = RFREQUESTERR 'Ѱ������ʧ��
        Call quit
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        RecycleCard = RFANTICOLLERR '������ͻ�����ؿ������к�
        Call quit
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        RecycleCard = RFSELECTERR '�Ӷ������ѡȡһ���������кŵĿ�
        Call quit
        Exit Function
    End If
    
    'װ��B����
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          RecycleCard = RFLOADKEY_B_ERR 'װ���������_B
          Call quit
          Exit Function
    End If
    
    '��֤����B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          RecycleCard = RFAUTHENTICATION_B_ERR '�ÿ�Ƭ�����ڱ�ϵͳ��ʹ��
          Call quit
          Exit Function
    End If


    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &HC6
    akey(3) = &H8C
    akey(4) = &HD9
    akey(5) = &HE5

    bkey(0) = &HB0
    bkey(1) = &H1B
    bkey(2) = &H4C
    bkey(3) = &H49
    bkey(4) = &HA3
    bkey(5) = &HD3
     st = rf_changeb3(icdev, sector, akey(0), 0, 0, 0, 1, 0, bkey(0))
     If st <> 0 Then
        RecycleCard = RFCHANGEB3ERR '������ʧ��
        Call quit
        Exit Function
    End If
   
    st = rf_beep(icdev, 5)
    
    RecycleCard = RFOK
    'ȡ���豸
    Call quit
    
End Function

' ========================================================================
'  “d˜b‰ž‘Î“ú•ñ WŒvƒVƒXƒeƒ€
'  ƒf[ƒ^ƒx[ƒX (Access) ‚ðŽ©“®‚Åì‚è‚Ü‚·
'
'  ‚±‚Ìƒtƒ@ƒCƒ‹‚ðƒ_ƒuƒ‹ƒNƒŠƒbƒN‚·‚é‚ÆA“¯‚¶ƒtƒHƒ‹ƒ_‚É
'  “ú•ñWŒv_be.accdb ‚ðì‚è‚Ü‚·B10 •b‚Ù‚Ç‚ÅI‚í‚è‚Ü‚·B
'
'  Access ‚ªƒCƒ“ƒXƒg[ƒ‹‚³‚ê‚Ä‚¢‚éƒpƒ\ƒRƒ“‚ÅŽÀs‚µ‚Ä‚­‚¾‚³‚¢B
'  ì‚ç‚ê‚é accdb ‚É‚Í VBA ‚ª“ü‚ç‚È‚¢‚Ì‚ÅAƒ}ƒNƒ‚ÌŒx‚ào‚Ü‚¹‚ñB
'
'  ¦ ‚±‚Ìƒtƒ@ƒCƒ‹‚Í tools/gen_installer_vbs.py ‚ª src/*.bas ‚©‚ç
'     Ž©“®¶¬‚µ‚Ä‚¢‚Ü‚·BŽè‚Å•ÒW‚¹‚¸A¶¬‚µ’¼‚µ‚Ä‚­‚¾‚³‚¢B
' ========================================================================
Option Explicit

Dim fso, shell, here, dbPath, acc, db, i, total, done
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = fso.BuildPath(here, "“ú•ñWŒv_be.accdb")

' ‚·‚Å‚É‚ ‚éê‡‚Íì‚è’¼‚µ‚Ä‚æ‚¢‚©Šm”F‚·‚é
If fso.FileExists(dbPath) Then
  If MsgBox("‚·‚Å‚É “ú•ñWŒv_be.accdb ‚ª‚ ‚è‚Ü‚·B" & vbCrLf & vbCrLf & _
            "ì‚è’¼‚·‚ÆA“ü—ÍÏ‚Ý‚Ìƒf[ƒ^‚Í‚·‚×‚ÄÁ‚¦‚Ü‚·B" & vbCrLf & _
            "–{“–‚Éì‚è’¼‚µ‚Ü‚·‚©H", _
            vbExclamation + vbYesNo + vbDefaultButton2, "Šm”F") <> vbYes Then
    MsgBox "’†Ž~‚µ‚Ü‚µ‚½B", vbInformation, "“d˜b‰ž‘Î“ú•ñ"
    WScript.Quit
  End If
  On Error Resume Next
  fso.DeleteFile dbPath
  If Err.Number <> 0 Then
    MsgBox "ŒÃ‚¢ƒtƒ@ƒCƒ‹‚ðÁ‚¹‚Ü‚¹‚ñ‚Å‚µ‚½B" & vbCrLf & vbCrLf & _
           "Access ‚ÅŠJ‚¢‚½‚Ü‚Ü‚É‚È‚Á‚Ä‚¢‚È‚¢‚©Šm”F‚µ‚Ä‚­‚¾‚³‚¢B", _
           vbCritical, "“d˜b‰ž‘Î“ú•ñ"
    WScript.Quit
  End If
  On Error GoTo 0
End If

On Error Resume Next
Set acc = CreateObject("Access.Application")
If Err.Number <> 0 Then
  MsgBox "Access ‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ‚Å‚µ‚½B" & vbCrLf & vbCrLf & _
         "‚±‚Ìƒpƒ\ƒRƒ“‚É Microsoft Access ‚ª“ü‚Á‚Ä‚¢‚é‚©Šm”F‚µ‚Ä‚­‚¾‚³‚¢B", _
         vbCritical, "“d˜b‰ž‘Î“ú•ñ"
  WScript.Quit
End If
On Error GoTo 0

acc.Visible = False
acc.NewCurrentDatabase dbPath
Set db = acc.CurrentDb

done = 0
total = 192

' --- •\‚ðì‚é ---
Run "CREATE TABLE [M_WŒv—ñ] ([WŒv—ñID] LONG NOT NULL CONSTRAINT [PK_WŒv—ñ] PRIMARY KEY,[WŒv—ñ–¼] TEXT(50) NOT NULL,[•\Ž¦‡] LONG NOT NULL)"
Run "CREATE TABLE [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID] LONG NOT NULL CONSTRAINT [PK_ƒuƒƒbƒN] PRIMARY KEY,[ƒuƒƒbƒN–¼] TEXT(50) NOT NULL,[»•i•Ê] BIT NOT NULL,[•\Ž¦‡] LONG NOT NULL)"
Run "CREATE TABLE [M_»•i] ([»•iID] LONG NOT NULL CONSTRAINT [PK_»•i] PRIMARY KEY,[ƒuƒƒbƒNID] LONG NOT NULL,[»•i–¼] TEXT(100) NOT NULL,[“K—pŠJŽn“ú] DATETIME,[“K—pI—¹“ú] DATETIME,[•\Ž¦‡] LONG NOT NULL,[—LŒø] BIT NOT NULL)"
Run "CREATE TABLE [M_‹æ•ª] ([‹æ•ªID] LONG NOT NULL CONSTRAINT [PK_‹æ•ª] PRIMARY KEY,[ƒuƒƒbƒNID] LONG NOT NULL,[‹æ•ª–¼] TEXT(100) NOT NULL,[WŒv—ñID] LONG NOT NULL,[“à–ó‹æ•ª] TEXT(20),[•\Ž¦‡] LONG NOT NULL,[—LŒø] BIT NOT NULL,[‹Œ“]‹L–¼] TEXT(100),CONSTRAINT [FK_‹æ•ª_ƒuƒƒbƒN] FOREIGN KEY ([ƒuƒƒbƒNID]) REFERENCES [M_ƒuƒƒbƒN]([ƒuƒƒbƒNID]),CONSTRAINT [FK_‹æ•ª_WŒv—ñ] FOREIGN KEY ([WŒv—ñID]) REFERENCES [M_WŒv—ñ]([WŒv—ñID]))"
Run "CREATE TABLE [M_’S“–ŽÒ] ([’S“–ŽÒID] LONG NOT NULL CONSTRAINT [PK_’S“–ŽÒ] PRIMARY KEY,[’S“–ŽÒƒR[ƒh] TEXT(10) NOT NULL,[©] TEXT(50) NOT NULL,[–¼] TEXT(50),[Ž–¼] TEXT(100) NOT NULL,[ƒJƒi] TEXT(100),[Eˆõ‹æ•ª] TEXT(10) NOT NULL,[ÝÐŠJŽn“ú] DATETIME,[ÝÐI—¹“ú] DATETIME,[•\Ž¦‡] LONG NOT NULL,[—LŒø] BIT NOT NULL,[”õl] MEMO)"
Run "CREATE TABLE [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID] LONG NOT NULL CONSTRAINT [PK_‹Æ–±€–Ú] PRIMARY KEY,[”Ô†] TEXT(4) NOT NULL,[€–Ú–¼] TEXT(100) NOT NULL,[’ •[•\Ž¦–¼] TEXT(100) NOT NULL,[•\Ž¦‡] LONG NOT NULL,[—LŒø] BIT NOT NULL)"
Run "CREATE TABLE [T_“ú•ñ] ([‘ÎÛ“ú] DATETIME NOT NULL CONSTRAINT [PK_“ú•ñ] PRIMARY KEY,[‰ñü”] LONG,[“Á‹LŽ–€] MEMO,[Eˆõ‘ã‘ÖˆÄŒ] MEMO,[—v–]] MEMO,[ó‘Ô] TEXT(10) NOT NULL,[Šm’è“úŽž] DATETIME,[XV“úŽž] DATETIME)"
Run "CREATE TABLE [T_o‹Î] ([o‹ÎID] COUNTER NOT NULL CONSTRAINT [PK_o‹Î] PRIMARY KEY,[‘ÎÛ“ú] DATETIME NOT NULL,[’S“–ŽÒID] LONG NOT NULL,[‹Î–±ŽžŠÔ] TEXT(50),[”õl] TEXT(255),CONSTRAINT [UQ_o‹Î] UNIQUE ([‘ÎÛ“ú],[’S“–ŽÒID]))"
Run "CREATE TABLE [T_Žó“d] ([Žó“dID] COUNTER NOT NULL CONSTRAINT [PK_Žó“d] PRIMARY KEY,[‘ÎÛ“ú] DATETIME NOT NULL,[’S“–ŽÒID] LONG NOT NULL,[‹æ•ªID] LONG NOT NULL,[»•iID] LONG NOT NULL,[Œ”] LONG NOT NULL,[”õl2] TEXT(255),[”õl3] TEXT(255),[“o˜^“úŽž] DATETIME,[XV“úŽž] DATETIME,[“o˜^ŽÒ] TEXT(100),CONSTRAINT [UQ_Žó“d] UNIQUE ([‘ÎÛ“ú],[’S“–ŽÒID],[‹æ•ªID],[»•iID]))"
Run "CREATE TABLE [T_‹Æ–±ŽÀÑ] ([ŽÀÑID] COUNTER NOT NULL CONSTRAINT [PK_‹Æ–±ŽÀÑ] PRIMARY KEY,[‘ÎÛ“ú] DATETIME NOT NULL,[’S“–ŽÒID] LONG NOT NULL,[‹Æ–±€–ÚID] LONG NOT NULL,[Œ”] LONG NOT NULL,CONSTRAINT [UQ_‹Æ–±ŽÀÑ] UNIQUE ([‘ÎÛ“ú],[’S“–ŽÒID],[‹Æ–±€–ÚID]))"
Run "CREATE TABLE [T_ŽæžƒƒO] ([ƒƒOID] COUNTER NOT NULL CONSTRAINT [PK_ŽæžƒƒO] PRIMARY KEY,[Žæž“úŽž] DATETIME,[ŽæžŒ³] TEXT(255),[s”Ô†] LONG,[“ú•t] TEXT(50),[’S“–ŽÒ] TEXT(100),[»•i–¼] TEXT(100),[‹æ•ª–¼] TEXT(100),[Œ”] LONG,[——R] MEMO)"
Run "CREATE INDEX [IX_’S“–ŽÒ_ƒR[ƒh] ON [M_’S“–ŽÒ] ([’S“–ŽÒƒR[ƒh])"
Run "CREATE INDEX [IX_Žó“d_“ú•t] ON [T_Žó“d] ([‘ÎÛ“ú])"

' --- ‘I‘ðŽˆ‚Ì‚à‚Æ‚É‚È‚éƒf[ƒ^‚ð“ü‚ê‚é ---
Run "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (3,'\ž•û–@',1)"
Run "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (4,'’Š‘IŒ‹‰Ê',2)"
Run "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (5,'”[•t‘”­‘—',3)"
Run "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (6,'¤•i”­‘—',4)"
Run "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (7,'‚»‚Ì‘¼',5)"
Run "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (8,'¤•iŒðŠ·',6)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (0,'i‚È‚µj',False,0)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (1,'‹æ•ª',True,1)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (2,'‚»‚Ì‘¼',True,2)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (3,'ŒÚ‹qî•ñ',False,3)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (4,'ƒCƒxƒ“ƒgŠÖŒW',False,4)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (5,'‚»‚Ì‘¼‚Ì‚»‚Ì‘¼‡@',False,5)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (6,'‚»‚Ì‘¼‚Ì‚»‚Ì‘¼‡A',False,6)"
Run "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (7,'“ÁŽê‚È–â‡‚¹',False,7)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (0,0,'i»•iŽw’è‚È‚µj',0,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (1,1,'ƒ~ƒ“ƒg',1,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (2,1,'’Êíƒvƒ‹[ƒt',2,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (3,1,'’Ê”Ni‹L”O“úEƒWƒƒƒpƒ“j',3,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (4,1,'‘—§Œö‰€‹L”O‰Ý',4,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (5,1,'ƒhƒ‰ƒSƒ“ƒvƒ‹[ƒtE‰Ý•¼ƒZƒbƒg',5,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (6,1,'’¹bl•¨‹Y‰æ‰Ý•¼ƒZƒbƒg',6,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (7,1,'‚h‚b‚c‚bƒƒ_ƒ‹',7,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (8,1,'÷‚Ì’Ê‚è”²‚¯ƒvƒ‹[ƒt',8,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (9,1,'÷‚Ì’Ê‚è”²‚¯‰Ý•¼ƒZƒbƒg',9,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (10,1,'‰Ô‚Ì‚Ü‚í‚è‚Ý‚¿‰Ý•¼ƒZƒbƒg',10,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (11,1,'÷‚Ì’Ê‚è”²‚¯‹L”Oƒƒ_ƒ‹',11,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (12,1,'ƒ‹àƒƒ_ƒ‹-¯À-',12,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (13,1,'‘•óÍ”vu’¹bl•¨‹Y‰æv',13,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (14,1,'Š±Žxƒƒ_ƒ‹',14,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (15,1,'ƒAƒWƒA‘å‰ï‹L”O‰Ý',15,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (16,1,'ƒRƒiƒ“ƒvƒ‹[ƒtE‰Ý•¼ƒZƒbƒg',16,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (17,1,'º˜a100”N‹L”O‰Ý•¼',17,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (18,1,'Žµ•óÍ”vu’·•l‰gŽRÕv',18,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (19,1,'i—\”õ19j',19,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (20,1,'i—\”õ20j',20,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (21,1,'i—\”õ21j',21,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (22,1,'i—\”õ22j',22,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (23,1,'i—\”õ23j',23,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (24,2,'ƒIƒŠƒ“ƒsƒbƒN‹L”O‰Ý(‰ß‹Žj',1,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (25,2,'cŽºŠÖŒW‹L”O‰Ý(‰ß‹Ž)',2,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (26,2,'ƒ~ƒ“ƒg',3,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (27,2,'’Êíƒvƒ‹[ƒt',4,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (28,2,'’Ê”Ni‹L”O“úEƒWƒƒƒpƒ“j',5,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (29,2,'¢ŠE•¶‰»ˆâŽYƒZƒbƒg',6,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (30,2,'‚h‚b‚c‚bƒƒ_ƒ‹',7,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (31,2,'Š±Žxƒƒ_ƒ‹',8,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (32,2,'ƒ‹àƒƒ_ƒ‹-¯À-',9,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (33,2,'–œ”Ž‹L”O‰Ý',10,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (34,2,'‘—§Œö‰€‹L”O‰Ý',11,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (35,2,'–œ”Žƒƒ_ƒ‹(‰ß‹Žj',12,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (36,2,'¢ŠE—¤ãƒvƒ‹[ƒt‰Ý•¼ƒZƒbƒg',13,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (37,2,'’¹bl•¨‹Y‰æƒP[ƒX',14,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (38,2,'’n•ûŽ©Ž¡ç‰~‹â‰Ý•¼',15,True)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (39,2,'i—\”õ16j',16,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (40,2,'i—\”õ17j',17,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (41,2,'i—\”õ18j',18,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (42,2,'i—\”õ19j',19,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (43,2,'i—\”õ20j',20,False)"
Run "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (44,2,'i—\”õ21j',21,False)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (1,1,'\žŠÖŒW',3,Null,1,True,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (2,1,'Žó’',3,Null,2,True,'Žó’')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (3,1,'’Š‘IŒ‹‰Ê',4,Null,3,True,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (4,1,'•¥ž—pŽ†EÄ”­sE‰Â',5,Null,4,True,'•¥ž—pŽ†Ä”­si‰Âj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (5,1,'•¥ž—pŽ†EÄ”­sE•s‰Â',5,Null,5,True,'•¥ž—pŽ†Ä”­si•s‰Âj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (6,1,'•¥ž—pŽ†E“à—e“™EÆ‰ï',5,Null,6,True,'•¥ž—pŽ†“à—eÆ‰ï')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (7,1,'•¥ž—pŽ†E”­‘—ó‹µEÆ‰ï',5,Null,7,True,'•¥ž—pŽ†”­‘—ó‹µ')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (8,1,'“ü‹àŠÖŒW',7,Null,8,True,'“ü‹àŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (9,1,'¤•i”­‘—Æ‰ïE–â‡‚¹',6,Null,9,True,'¤•i”­‘—–â‡‚¹')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (10,1,'¤•i”­‘—Æ‰ïE–¢’…',6,Null,10,True,'¤•i”­‘—Žó—ÌÆ‰ï')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (11,1,'»•iŒðŠ·',8,Null,11,True,'»•iŒðŠ·')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (12,2,'”Ì”„—\’è',7,Null,1,True,'”Ì”„—\’è')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (13,2,'»•i“à—e',7,Null,2,True,'»•i“à—e')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (14,2,'ÝŒÉÆ‰ï',7,Null,3,True,'ÝŒÉÆ‰ï')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (15,2,'‰¿ŠiÆ‰ï',7,Null,4,True,'‰¿ŠiÆ‰ï')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (16,2,'“à—e–â‡‚¹',7,Null,5,False,'“à—e–â‡‚¹')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (17,2,'‘gž‚Ýƒ~ƒXŒðŠ·',7,Null,6,False,'‘gž‚Ýƒ~ƒXŒðŠ·')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (18,2,'‚»‚Ì‘¼ŒðŠ·',7,Null,7,False,'‚»‚Ì‘¼ŒðŠ·')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (19,2,'i–¢Žg—p8j',7,Null,8,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (20,2,'i–¢Žg—p9j',7,Null,9,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (21,2,'i–¢Žg—p10j',7,Null,10,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (22,2,'i–¢Žg—p11j',7,Null,11,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (23,3,'‚c‚l’âŽ~“™EŽ€–S',7,Null,1,True,'‚c‚l’âŽ~(‡@Ž€–S)')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (24,3,'‚c‚l’âŽ~“™E•a‹C“™',7,Null,2,True,'‚c‚l’âŽ~(‡A•a‹C“™)')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (25,3,'‚c‚l’âŽ~“™EŽí—Þ‘½‚¢',7,Null,3,True,'‚c‚l’âŽ~(‡BŽí—Þ‘½‚¢)')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (26,3,'‚c‚l’âŽ~“™E“Á’è¤•i‚Ì‚Ý',7,Null,4,True,'‚c‚l’âŽ~(‡C“Á’è¤•i‚Ì‚Ý)')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (27,3,'‚c‚l’âŽ~“™E——R‚È‚µ',7,Null,5,True,'‚c‚l’âŽ~(‡D——R‚È‚µ)')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (28,3,'‚c‚lÄŠJEˆ¶–¼l•ÏX',7,Null,6,True,'‚c‚lÄŠJEˆ¶–¼l•ÏX')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (29,3,'‚c‚lŠm”F',7,Null,7,True,'‚c‚lŠm”F')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (30,3,'ZŠ•ÏX',7,Null,8,True,'ZŠ•ÏX')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (31,3,'V‹K“o˜^',7,Null,9,True,'V‹K“o˜^')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (32,3,'Ž‘—¿‘—•t',7,Null,10,True,'Ž‘—¿‘—•t')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (33,3,'w“ü—š—ð',7,Null,11,True,'w“ü—š—ð')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (34,4,'‚¨‹à‚ÆØŽè',7,Null,1,True,'ƒCƒxƒ“ƒgŠÖŒWi‚¨‹à‚ÆØŽèj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (35,4,'‰F²_‹{',7,Null,2,True,'ƒCƒxƒ“ƒgŠÖŒWi‘¢•¼‹Ç ‚h‚mj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (36,4,'÷‚Ì’Ê‚è”²‚¯',7,Null,3,True,'ƒCƒxƒ“ƒgŠÖŒWi÷‚Ì’Ê‚è”²‚¯j')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (37,4,'‰Ô‚Ì‚Ü‚í‚è“¹',7,Null,4,True,'ƒCƒxƒ“ƒgŠÖŒWi‚Ü‚í‚è“¹j')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (38,4,'‚s‚h‚b‚b',7,Null,5,True,'ƒCƒxƒ“ƒgŠÖŒWiTICCj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (39,4,'‘‘Ì',7,Null,6,True,'ƒCƒxƒ“ƒgŠÖŒWi‘‘Ìj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (40,4,'‘åãƒRƒCƒ“ƒVƒ‡[',7,Null,7,True,'ƒCƒxƒ“ƒgŠÖŒWi‘åãƒRƒCƒ“j')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (41,4,'‚³‚¢‚½‚ÜƒtƒFƒA',7,Null,8,True,'ƒCƒxƒ“ƒgŠÖŒWi‚³‚¢‚½‚ÜƒtƒFƒAj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (42,4,'²”Œ‹æ–¯‚Ü‚Â‚è',7,Null,9,True,'ƒCƒxƒ“ƒgŠÖŒWi–¼ŒÃ‰®‰Ý•¼‚Ü‚Â‚èj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (43,4,'ŠÏ÷‰ï',7,Null,10,True,'ƒCƒxƒ“ƒgŠÖŒWiŠÏ÷‰ï‚Ì“d˜b“]‘—j')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (44,4,'ƒAƒ“ƒP[ƒg',7,Null,11,True,'ƒCƒxƒ“ƒgŠÖŒWi—\”õj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (45,5,'ƒIƒ“ƒ‰ƒCƒ““o˜^ŠÖŒW',7,Null,1,True,'ƒIƒ“ƒ‰ƒCƒ““o˜^ŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (46,5,'ƒNƒŒƒWƒbƒgŠÖŒW',7,Null,2,True,'ƒNƒŒƒWƒbƒgŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (47,5,'‹àH•i‚fŠÖŒW',7,Null,3,True,'‹àH•i‚f‚Ö“d˜b“]‘—')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (48,5,'L•ñiHêŒ©ŠwŠÜ‚Þj',7,Null,4,True,'L•ñŽºiHêŒ©ŠwŠÜ‚Þj')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (49,5,'ƒ~ƒ“ƒgƒVƒ‡ƒbƒvˆÄ“à',7,Null,5,True,'ƒ~ƒ“ƒgƒVƒ‡ƒbƒvˆÄ“à')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (50,5,'‹àH•i§’k‰ï',7,Null,6,True,'÷ƒTƒ|[ƒg')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (51,5,'—¬’Ê‰Ý•¼ŠÖŒW',7,Null,7,True,'—¬’Ê‰Ý•¼ŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (52,5,'Ž†•¼ŠÖŒW',7,Null,8,True,'Ž†•¼ŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (53,5,'ƒ~ƒ“ƒgƒNƒ‰ƒuŠÖŒW',7,Null,9,True,'ƒ~ƒ“ƒgƒNƒ‰ƒuŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (54,5,'1‰~‚Ê‚¢‚®‚é‚Ý',7,Null,10,True,'1‰~‚Ê‚¢‚®‚é‚Ý')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (55,5,'‹L”O‰ÝˆøŠ·ŠÖŒW',7,Null,11,True,'‹L”O‰ÝˆøŠ·ŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (56,6,'•Ô •i',7,'•Ô•i',1,True,'•Ô•i')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (57,6,'•Ô ‹à',7,'•Ô‹à',2,True,'•Ô‹à')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (58,6,'i–¢Žg—p3j',7,Null,3,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (59,6,'i–¢Žg—p4j',7,Null,4,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (60,6,'i–¢Žg—p5j',7,Null,5,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (61,6,'‰Ý•¼¾¯Äˆ•ªŠÖŒW',7,Null,6,True,'‰Ý•¼¾¯Äˆ•ªŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (62,6,'‰Ý•¼ôò•û–@',7,Null,7,True,'‰Ý•¼ôò•û–@')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (63,6,'–¼Ì–¢Šm’è»•i—\’è',7,Null,8,True,'–¼Ì–¢Šm’è»•i')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (64,6,'‹ÆŽÒ”Ì”„ŠÖŒW',7,Null,9,True,'‹ÆŽÒ”Ì”„ŠÖŒW')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (65,6,'’Š‘I”{—¦“™(‰ß‹Žj',7,Null,10,True,'’Š‘I”{—¦“™')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (66,6,'i–¢Žg—p11j',7,Null,11,False,Null)"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (67,7,'•Ê“Y‚Ì‚Æ‚¨‚è',7,Null,1,True,'“ÁŽê‚È–â‡‚¹')"
Run "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (68,7,'‰º‹L‚Ì‚Æ‚¨‚èi•Ê“Y•s—vj',7,Null,2,True,'“ÁŽê‚È–â‡‚¹')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (1,'001','îà“ˆ','–¼•Û”ü','îà“ˆ –¼•Û”ü',Null,'ƒp[ƒg',Null,Null,1,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (2,'002','–x',Null,'–x ^Ÿ',Null,'ƒp[ƒg',Null,Null,2,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (3,'003','ŽOŒ´','‘½ŒbŽq','ŽOŒ´ ‘½ŒbŽq',Null,'ƒp[ƒg',Null,Null,3,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (4,'004','’JŒû',Null,'’JŒû —F‰ÀŽq',Null,'ƒp[ƒg',Null,Null,4,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (5,'006','‘X',Null,'‘X ‹IŽq',Null,'ƒp[ƒg',Null,Null,5,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (6,'007','ŽRŒû','ŒbŽq','ŽRŒû ŒbŽq',Null,'ƒp[ƒg',Null,Null,6,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (7,'008','ˆ§â',']—¢‰Á','ˆ§â ]—¢‰Á',Null,'ƒp[ƒg',Null,Null,7,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (8,'009','¼‰ª',Null,'¼‰ª —R”üŽq',Null,'ƒp[ƒg',Null,Null,8,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (9,'010','ÎŒ´','—TŽq','ÎŒ´ —TŽq',Null,'ƒp[ƒg',Null,Null,9,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (10,'011','Îˆä','~Žq','Îˆä ~Žq',Null,'ƒp[ƒg',Null,Null,10,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (11,'012','îà‹´','—ŒbŽq','îà‹´ —ŒbŽq',Null,'ƒp[ƒg',Null,Null,11,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (12,'013','—\”õ',Null,'—\”õ',Null,'ƒp[ƒg',Null,Null,12,False,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (13,'014','•Ÿ“‡',Null,'•Ÿ“‡',Null,'ƒp[ƒg',Null,Null,13,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwîà‹´x‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (14,'015','—é–Ø',Null,'—é–Ø',Null,'ƒp[ƒg',Null,Null,14,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (15,'016','¬àV','•üŽq','¬àV •üŽq',Null,'ƒp[ƒg',Null,Null,15,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (16,'017','Ž›–{','–@Žq','Ž›–{ –@Žq',Null,'ƒp[ƒg',Null,Null,16,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (17,'018','—é–Ø’q','’qŽq','—é–Ø ’qŽq',Null,'ƒp[ƒg',Null,Null,17,True,Null)"
Run "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (18,'100','ŒÚ‹q‚f',Null,'ŒÚ‹qG',Null,'Eˆõ',Null,Null,18,True,Null)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (1,'‡@','‡@Žó’“ü—Í','‡@@Žó’“ü—Í',1,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (2,'‡A','‡AŽó’ƒ`ƒFƒbƒN','‡A@Žó’ƒ`ƒFƒbƒN',2,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (3,'‡B','‡B•¥ž‘ƒ`ƒFƒbƒN','‡B@–ß‚è—X•Öˆ—^•¥ž‘ƒ`ƒFƒbƒN',3,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (4,'‡C','‡C–ß‚è—X•Ö','‡C@‚c‚lˆ—^–ß‚è—X•Ö',4,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (5,'‡D','‡D‰Ë“d','‡D@‰Ë“d',5,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (6,'‡E','‡E‚»‚Ì‘¼','‡E@‚»‚Ì‘¼i@@@@@@@@j',6,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (7,'‡F','‡FƒGƒNƒZƒ‹“ü—Í','‡F@ƒGƒNƒZƒ‹“ü—Í',7,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (8,'‡G','‡GƒGƒNƒZƒ‹ƒ`ƒFƒbƒN','‡G@ƒGƒNƒZƒ‹ƒ`ƒFƒbƒN',8,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (9,'‡H','‡H±Ý¹°Ä“ü—Í','‡H@ƒAƒ“ƒP[ƒg“ü—Í',9,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (10,'‡I','‡IÊ¶Þ·ÃÞ°À“ü—Í','‡I@Ê¶Þ·ÃÞ°À“ü—Í',10,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (11,'‡J','‡JÊ¶Þ·ÃÞ°ÀÁª¯¸','‡J@ƒnƒKƒLƒf[ƒ^ƒ`ƒFƒbƒN',11,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (12,'‡K','‡KV‹Kº°ÄÞŽæ‚è','‡K@V‹Kº°ÄÞŽæ‚è',12,True)"
Run "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (13,'‡L','‡LŒÚ‹q®—','‡L@ŒÚ‹q®—',13,True)"

' --- WŒv‚Ì‚µ‚©‚½‚ð“o˜^‚·‚é ---
MakeQuery "Q_‘I‘ð_’S“–ŽÒ", "SELECT OP.[’S“–ŽÒID], OP.[’S“–ŽÒƒR[ƒh], OP.[Ž–¼], OP.[Eˆõ‹æ•ª], OP.[•\Ž¦‡] FROM [M_’S“–ŽÒ] AS OP WHERE OP.[—LŒø]=True   AND (OP.[ÝÐŠJŽn“ú] Is Null Or OP.[ÝÐŠJŽn“ú] <= Date())   AND (OP.[ÝÐI—¹“ú] Is Null Or OP.[ÝÐI—¹“ú] >= Date()) ORDER BY OP.[•\Ž¦‡];"
MakeQuery "Q_‘I‘ð_’S“–ŽÒ_Šî€“ú", "PARAMETERS [Šî€“ú] DateTime; SELECT OP.[’S“–ŽÒID], OP.[’S“–ŽÒƒR[ƒh], OP.[Ž–¼], OP.[Eˆõ‹æ•ª], OP.[•\Ž¦‡] FROM [M_’S“–ŽÒ] AS OP WHERE OP.[—LŒø]=True   AND (OP.[ÝÐŠJŽn“ú] Is Null Or OP.[ÝÐŠJŽn“ú] <= [Šî€“ú])   AND (OP.[ÝÐI—¹“ú] Is Null Or OP.[ÝÐI—¹“ú] >= [Šî€“ú]) ORDER BY OP.[•\Ž¦‡];"
MakeQuery "Q_‘I‘ð_‹æ•ª", "SELECT KB.[‹æ•ªID], KB.[ƒuƒƒbƒNID], BK.[ƒuƒƒbƒN–¼], KB.[‹æ•ª–¼],        BK.[ƒuƒƒbƒN–¼] & ' / ' & KB.[‹æ•ª–¼] AS [•\Ž¦–¼],        KB.[WŒv—ñID], SC.[WŒv—ñ–¼], BK.[»•i•Ê],        BK.[•\Ž¦‡]*1000 + KB.[•\Ž¦‡] AS [•À‚Ñ‡] FROM ([M_‹æ•ª] AS KB INNER JOIN [M_ƒuƒƒbƒN] AS BK ON KB.[ƒuƒƒbƒNID]=BK.[ƒuƒƒbƒNID])      INNER JOIN [M_WŒv—ñ] AS SC ON KB.[WŒv—ñID]=SC.[WŒv—ñID] WHERE KB.[—LŒø]=True ORDER BY BK.[•\Ž¦‡]*1000 + KB.[•\Ž¦‡];"
MakeQuery "Q_‘I‘ð_»•i", "SELECT PR.[»•iID], PR.[ƒuƒƒbƒNID], BK.[ƒuƒƒbƒN–¼], PR.[»•i–¼],        IIf(PR.[»•iID]=0, PR.[»•i–¼], BK.[ƒuƒƒbƒN–¼] & ' / ' & PR.[»•i–¼]) AS [•\Ž¦–¼],        PR.[•\Ž¦‡] FROM [M_»•i] AS PR INNER JOIN [M_ƒuƒƒbƒN] AS BK ON PR.[ƒuƒƒbƒNID]=BK.[ƒuƒƒbƒNID] WHERE PR.[—LŒø]=True   AND (PR.[“K—pŠJŽn“ú] Is Null Or PR.[“K—pŠJŽn“ú] <= Date())   AND (PR.[“K—pI—¹“ú] Is Null Or PR.[“K—pI—¹“ú] >= Date()) ORDER BY PR.[ƒuƒƒbƒNID], PR.[•\Ž¦‡];"
MakeQuery "Q_Žó“d–¾×", "SELECT J.[Žó“dID], J.[‘ÎÛ“ú], OP.[’S“–ŽÒID], OP.[’S“–ŽÒƒR[ƒh], OP.[Ž–¼],        OP.[©] AS [’S“–ŽÒ], OP.[Eˆõ‹æ•ª], BK.[ƒuƒƒbƒN–¼],        PR.[»•i–¼], KB.[‹æ•ªID], KB.[‹æ•ª–¼], KB.[WŒv—ñID], SC.[WŒv—ñ–¼],        IIf(KB.[WŒv—ñID]=3,J.[Œ”],0) AS [\ž•û–@],        IIf(KB.[WŒv—ñID]=4,J.[Œ”],0) AS [’Š‘IŒ‹‰Ê],        IIf(KB.[WŒv—ñID]=5,J.[Œ”],0) AS [”[•t‘”­‘—],        IIf(KB.[WŒv—ñID]=6,J.[Œ”],0) AS [¤•i”­‘—],        IIf(KB.[WŒv—ñID]=7,J.[Œ”],0) AS [‚»‚Ì‘¼],        " & _
        "IIf(KB.[WŒv—ñID]=8,J.[Œ”],0) AS [¤•iŒðŠ·],        J.[Œ”] AS [Œv], J.[”õl2], J.[”õl3] FROM ((((([T_Žó“d] AS J   INNER JOIN [M_’S“–ŽÒ] AS OP ON J.[’S“–ŽÒID]=OP.[’S“–ŽÒID])   INNER JOIN [M_‹æ•ª]   AS KB ON J.[‹æ•ªID]=KB.[‹æ•ªID])   INNER JOIN [M_WŒv—ñ] AS SC ON KB.[WŒv—ñID]=SC.[WŒv—ñID])   INNER JOIN [M_ƒuƒƒbƒN] AS BK ON KB.[ƒuƒƒbƒNID]=BK.[ƒuƒƒbƒNID])   INNER JOIN [M_»•i]   AS PR ON J.[»•iID]=PR.[»•iID]) WHERE J.[Œ”]<>0;"
MakeQuery "Q_“ú•ÊWŒv", "TRANSFORM Sum(J.[Œ”]) AS [Œ”Œv] SELECT J.[‘ÎÛ“ú], Sum(J.[Œ”]) AS [Œv] FROM ([T_Žó“d] AS J INNER JOIN [M_‹æ•ª] AS KB ON J.[‹æ•ªID]=KB.[‹æ•ªID])      INNER JOIN [M_WŒv—ñ] AS SC ON KB.[WŒv—ñID]=SC.[WŒv—ñID] GROUP BY J.[‘ÎÛ“ú] PIVOT SC.[WŒv—ñ–¼] IN ('\ž•û–@','’Š‘IŒ‹‰Ê','”[•t‘”­‘—','¤•i”­‘—','‚»‚Ì‘¼','¤•iŒðŠ·');"
MakeQuery "Q_“ú•ñ_Žó“d", "SELECT J.[‘ÎÛ“ú],  Sum(IIf(KB.[WŒv—ñID]=3,J.[Œ”],0)) AS [\ž],  Sum(IIf(KB.[WŒv—ñID]=4,J.[Œ”],0)) AS [’Š‘I],  Sum(IIf(KB.[WŒv—ñID]=5,J.[Œ”],0)) AS [•¥ž—pŽ†],  Sum(IIf(KB.[WŒv—ñID]=6,J.[Œ”],0)) AS [¤•i”­‘—],  Sum(IIf(KB.[WŒv—ñID] In (7,8),J.[Œ”],0)) AS [‚»‚Ì‘¼],  Sum(J.[Œ”]) AS [‡Œv],  Sum(IIf(KB.[WŒv—ñID]=8,J.[Œ”],0)) AS [“àŒðŠ·],  Sum(IIf(KB.[“à–ó‹æ•ª]='•Ô‹à',J.[Œ”],0)) AS [“à•Ô‹à],  Sum(IIf(OP.[Eˆõ‹æ•ª]='Eˆõ' And KB.[WŒv—ñID]=3,J.[Œ”],0)) " & _
        "AS [\ž_Eˆõ],  Sum(IIf(OP.[Eˆõ‹æ•ª]='Eˆõ' And KB.[WŒv—ñID]=4,J.[Œ”],0)) AS [’Š‘I_Eˆõ],  Sum(IIf(OP.[Eˆõ‹æ•ª]='Eˆõ' And KB.[WŒv—ñID]=5,J.[Œ”],0)) AS [•¥ž—pŽ†_Eˆõ],  Sum(IIf(OP.[Eˆõ‹æ•ª]='Eˆõ' And KB.[WŒv—ñID]=6,J.[Œ”],0)) AS [¤•i”­‘—_Eˆõ],  Sum(IIf(OP.[Eˆõ‹æ•ª]='Eˆõ' And KB.[WŒv—ñID] In (7,8),J.[Œ”],0)) AS [‚»‚Ì‘¼_Eˆõ],  Sum(IIf(OP.[Eˆõ‹æ•ª]='Eˆõ',J.[Œ”],0)) AS [‡Œv_Eˆõ] FROM ([T_Žó“d] AS J INNER JOIN [M_‹æ•ª] AS KB ON J.[‹æ•ªID]=KB.[‹æ•ªID])      INNE" & _
        "R JOIN [M_’S“–ŽÒ] AS OP ON J.[’S“–ŽÒID]=OP.[’S“–ŽÒID] GROUP BY J.[‘ÎÛ“ú];"
MakeQuery "Q_“ú•ñ_‹Æ–±", "SELECT W.[‘ÎÛ“ú], TM.[‹Æ–±€–ÚID], TM.[”Ô†], TM.[€–Ú–¼], TM.[’ •[•\Ž¦–¼],        TM.[•\Ž¦‡], Sum(W.[Œ”]) AS [Œ”] FROM [T_‹Æ–±ŽÀÑ] AS W INNER JOIN [M_‹Æ–±€–Ú] AS TM      ON W.[‹Æ–±€–ÚID]=TM.[‹Æ–±€–ÚID] GROUP BY W.[‘ÎÛ“ú], TM.[‹Æ–±€–ÚID], TM.[”Ô†], TM.[€–Ú–¼], TM.[’ •[•\Ž¦–¼], TM.[•\Ž¦‡];"
MakeQuery "Q_“ú•ñ_o‹Î", "SELECT A.[‘ÎÛ“ú], OP.[’S“–ŽÒID], OP.[Ž–¼], OP.[Eˆõ‹æ•ª],        A.[‹Î–±ŽžŠÔ], A.[”õl], OP.[•\Ž¦‡] FROM [T_o‹Î] AS A INNER JOIN [M_’S“–ŽÒ] AS OP ON A.[’S“–ŽÒID]=OP.[’S“–ŽÒID] ORDER BY A.[‘ÎÛ“ú], OP.[•\Ž¦‡];"
MakeQuery "Q_“ú•ñ_ƒwƒbƒ_", "SELECT H.[‘ÎÛ“ú], H.[‰ñü”], H.[“Á‹LŽ–€], H.[Eˆõ‘ã‘ÖˆÄŒ], H.[—v–]],        H.[ó‘Ô], H.[Šm’è“úŽž],        (SELECT Count(*) FROM [T_o‹Î] AS A WHERE A.[‘ÎÛ“ú]=H.[‘ÎÛ“ú]) AS [o‹ÎŽÒ”],        Nz(R.[\ž],0) AS [\ž], Nz(R.[’Š‘I],0) AS [’Š‘I],        Nz(R.[•¥ž—pŽ†],0) AS [•¥ž—pŽ†], Nz(R.[¤•i”­‘—],0) AS [¤•i”­‘—],        Nz(R.[‚»‚Ì‘¼],0) AS [‚»‚Ì‘¼], Nz(R.[‡Œv],0) AS [‡Œv],        Nz(R.[“àŒðŠ·],0) AS [“àŒðŠ·], Nz(R.[“à•Ô‹à],0) AS [“à•Ô‹à],        Nz(R.[\ž_Eˆõ],0) AS [" & _
        "\ž_Eˆõ], Nz(R.[’Š‘I_Eˆõ],0) AS [’Š‘I_Eˆõ],        Nz(R.[•¥ž—pŽ†_Eˆõ],0) AS [•¥ž—pŽ†_Eˆõ],        Nz(R.[¤•i”­‘—_Eˆõ],0) AS [¤•i”­‘—_Eˆõ],        Nz(R.[‚»‚Ì‘¼_Eˆõ],0) AS [‚»‚Ì‘¼_Eˆõ], Nz(R.[‡Œv_Eˆõ],0) AS [‡Œv_Eˆõ] FROM [T_“ú•ñ] AS H LEFT JOIN [Q_“ú•ñ_Žó“d] AS R ON H.[‘ÎÛ“ú]=R.[‘ÎÛ“ú];"
MakeQuery "Q_’S“–ŽÒ•Ê“úŽŸ", "SELECT J.[‘ÎÛ“ú], OP.[’S“–ŽÒID], OP.[’S“–ŽÒƒR[ƒh], OP.[Ž–¼],        Sum(J.[Œ”]) AS [Žó“dŒ”] FROM [T_Žó“d] AS J INNER JOIN [M_’S“–ŽÒ] AS OP ON J.[’S“–ŽÒID]=OP.[’S“–ŽÒID] GROUP BY J.[‘ÎÛ“ú], OP.[’S“–ŽÒID], OP.[’S“–ŽÒƒR[ƒh], OP.[Ž–¼];"
MakeQuery "Q_–¢“ü—Íƒ`ƒFƒbƒN", "SELECT A.[‘ÎÛ“ú], OP.[’S“–ŽÒID], OP.[’S“–ŽÒƒR[ƒh], OP.[Ž–¼] FROM [T_o‹Î] AS A INNER JOIN [M_’S“–ŽÒ] AS OP ON A.[’S“–ŽÒID]=OP.[’S“–ŽÒID] WHERE NOT EXISTS (SELECT 1 FROM [T_Žó“d] AS J                   WHERE J.[‘ÎÛ“ú]=A.[‘ÎÛ“ú] AND J.[’S“–ŽÒID]=A.[’S“–ŽÒID] AND J.[Œ”]<>0)   AND NOT EXISTS (SELECT 1 FROM [T_‹Æ–±ŽÀÑ] AS W                   WHERE W.[‘ÎÛ“ú]=A.[‘ÎÛ“ú] AND W.[’S“–ŽÒID]=A.[’S“–ŽÒID] AND W.[Œ”]<>0) ORDER BY A.[‘ÎÛ“ú], OP.[•\Ž¦‡];"
MakeQuery "Q_TŽŸ–¾×", "PARAMETERS [ŠJŽn“ú] DateTime, [I—¹“ú] DateTime; SELECT * FROM [Q_Žó“d–¾×] WHERE [‘ÎÛ“ú] Between [ŠJŽn“ú] And [I—¹“ú] ORDER BY [‘ÎÛ“ú], [Ž–¼], [‹æ•ªID];"
MakeQuery "Q_TŽŸWŒv", "PARAMETERS [ŠJŽn“ú] DateTime, [I—¹“ú] DateTime; SELECT J.[‘ÎÛ“ú],  Sum(IIf(KB.[WŒv—ñID]=3,J.[Œ”],0)) AS [\ž•û–@],  Sum(IIf(KB.[WŒv—ñID]=4,J.[Œ”],0)) AS [’Š‘IŒ‹‰Ê],  Sum(IIf(KB.[WŒv—ñID]=5,J.[Œ”],0)) AS [”[•t‘”­‘—],  Sum(IIf(KB.[WŒv—ñID]=6,J.[Œ”],0)) AS [¤•i”­‘—],  Sum(IIf(KB.[WŒv—ñID]=7,J.[Œ”],0)) AS [‚»‚Ì‘¼],  Sum(IIf(KB.[WŒv—ñID]=8,J.[Œ”],0)) AS [¤•iŒðŠ·],  Sum(J.[Œ”]) AS [Œv] FROM [T_Žó“d] AS J INNER JOIN [M_‹æ•ª] AS KB ON J.[‹æ•ªID]=KB.[" & _
        "‹æ•ªID] WHERE J.[‘ÎÛ“ú] Between [ŠJŽn“ú] And [I—¹“ú] GROUP BY J.[‘ÎÛ“ú] ORDER BY J.[‘ÎÛ“ú];"
MakeQuery "Q_TŽŸWŒvŽw’è", "SELECT J.[‘ÎÛ“ú],  Sum(IIf(KB.[WŒv—ñID]=3,J.[Œ”],0)) AS [\ž•û–@],  Sum(IIf(KB.[WŒv—ñID]=4,J.[Œ”],0)) AS [’Š‘IŒ‹‰Ê],  Sum(IIf(KB.[WŒv—ñID]=5,J.[Œ”],0)) AS [”[•t‘”­‘—],  Sum(IIf(KB.[WŒv—ñID]=6,J.[Œ”],0)) AS [¤•i”­‘—],  Sum(IIf(KB.[WŒv—ñID]=7,J.[Œ”],0)) AS [‚»‚Ì‘¼],  Sum(IIf(KB.[WŒv—ñID]=8,J.[Œ”],0)) AS [¤•iŒðŠ·],  Sum(J.[Œ”]) AS [Œv] FROM [T_Žó“d] AS J INNER JOIN [M_‹æ•ª] AS KB ON J.[‹æ•ªID]=KB.[‹æ•ªID] GROUP BY J.[‘ÎÛ“ú];"

' --- •\‚Ç‚¤‚µ‚Ì‚Â‚È‚ª‚è‚ð“o˜^‚·‚é ---
'     ‚±‚±‚ªŒø‚¢‚Ä‚¢‚é‚ÆAŽÀÑ‚©‚çŽg‚í‚ê‚Ä‚¢‚éƒ}ƒXƒ^‚ÍÁ‚¹‚È‚­‚È‚é
AddRel "R_Žó“d_’S“–ŽÒ", "M_’S“–ŽÒ", "T_Žó“d", "’S“–ŽÒID"
AddRel "R_Žó“d_‹æ•ª", "M_‹æ•ª", "T_Žó“d", "‹æ•ªID"
AddRel "R_Žó“d_»•i", "M_»•i", "T_Žó“d", "»•iID"
AddRel "R_o‹Î_’S“–ŽÒ", "M_’S“–ŽÒ", "T_o‹Î", "’S“–ŽÒID"
AddRel "R_‹Æ–±ŽÀÑ_’S“–ŽÒ", "M_’S“–ŽÒ", "T_‹Æ–±ŽÀÑ", "’S“–ŽÒID"
AddRel "R_‹Æ–±ŽÀÑ_€–Ú", "M_‹Æ–±€–Ú", "T_‹Æ–±ŽÀÑ", "‹Æ–±€–ÚID"

acc.CloseCurrentDatabase
acc.Quit
Set db = Nothing
Set acc = Nothing

MsgBox "ƒf[ƒ^ƒx[ƒX‚ðì‚è‚Ü‚µ‚½B" & vbCrLf & vbCrLf & _
       dbPath & vbCrLf & vbCrLf & _
       "‚±‚Ìƒtƒ@ƒCƒ‹‚ð‹¤—LƒtƒHƒ‹ƒ_‚É’u‚¢‚Ä‚­‚¾‚³‚¢B" & vbCrLf & _
       "’u‚¢‚½êŠ‚ÍA‚ ‚Æ‚Å Web ƒT[ƒo[‚ÌÝ’è‚É‘‚«‚Ü‚·B", _
       vbInformation, "“d˜b‰ž‘Î“ú•ñ"

' ------------------------------------------------------------------------
Sub Run(sql)
  On Error Resume Next
  db.Execute sql, 128        ' 128 = dbFailOnError
  If Err.Number <> 0 Then Fail sql, Err.Description
  On Error GoTo 0
  done = done + 1
End Sub

Sub MakeQuery(nm, sql)
  On Error Resume Next
  db.CreateQueryDef nm, sql
  If Err.Number <> 0 Then Fail nm, Err.Description
  On Error GoTo 0
  done = done + 1
End Sub

Sub AddRel(nm, parentTable, childTable, fld)
  Dim rel
  On Error Resume Next
  Set rel = db.CreateRelation(nm, parentTable, childTable, 0)   ' 0 = ®‡«‚ðŽç‚é
  rel.Fields.Append rel.CreateField(fld)
  rel.Fields(fld).ForeignName = fld
  db.Relations.Append rel
  If Err.Number <> 0 Then Err.Clear    ' ‚Â‚È‚ª‚è‚Íì‚ê‚È‚­‚Ä‚à’v–½‚Å‚Í‚È‚¢
  On Error GoTo 0
  done = done + 1
End Sub

Sub Fail(what, why)
  On Error Resume Next
  acc.CloseCurrentDatabase
  acc.Quit
  On Error GoTo 0
  MsgBox "ì¬‚ÉŽ¸”s‚µ‚Ü‚µ‚½B" & vbCrLf & vbCrLf & _
         "ˆ—: " & Left(what, 120) & vbCrLf & _
         "——R: " & why & vbCrLf & vbCrLf & _
         "‚±‚Ì‰æ–Ê‚ðŽÊ^‚ÉŽB‚Á‚ÄA’S“–ŽÒ‚É‚¨’m‚ç‚¹‚­‚¾‚³‚¢B", _
         vbCritical, "“d˜b‰ž‘Î“ú•ñ"
  WScript.Quit
End Sub

Attribute VB_Name = "modSetupMaster"
Option Compare Database
Option Explicit

'==============================================================================
' ƒ}ƒXƒ^‰Šúƒf[ƒ^
'
' ‚±‚Ìƒ‚ƒWƒ…[ƒ‹‚Í tools/gen_master_vba.py ‚ª data/*.csv ‚©‚çŽ©“®¶¬‚·‚éB
' Žè‚Å•ÒW‚¹‚¸ACSV ‚ð’¼‚µ‚Ä‚©‚çÄ¶¬‚·‚é‚±‚ÆB
'
' o“T: Œ»s‚Ì‹L“ü—pƒtƒH[ƒ€ (Sheet1 ‚ÌŒ©o‚µ‚Í“ú•ñWŒvŠÇ——pƒtƒH[ƒ€‚Ö‚ÌŠO•”ƒŠƒ“ƒN)
'       ‚¨‚æ‚Ñ Sheet2 ‚Ì 2 s–Ú (WŒv•\‚Ì—ñ”Ô†) / 3 s–Ú (‹Œ“]‹L–¼)B
'==============================================================================

Public Sub Setup_Master()
    Echo_ "ƒ}ƒXƒ^‚ð“o˜^‚µ‚Ä‚¢‚Ü‚·..."

    ' ƒ}ƒXƒ^‚Í–ˆ‰ñì‚è’¼‚·‚Ì‚ÅA‚Ü‚¸‘SÁ‚µ‚·‚é
    ExecSQL "DELETE FROM [M_‹æ•ª]"
    ExecSQL "DELETE FROM [M_»•i]"
    ExecSQL "DELETE FROM [M_ƒuƒƒbƒN]"
    ExecSQL "DELETE FROM [M_WŒv—ñ]"
    ExecSQL "DELETE FROM [M_‹Æ–±€–Ú]"
    ExecSQL "DELETE FROM [M_’S“–ŽÒ]"

    Master_WŒv—ñ
    Master_ƒuƒƒbƒN
    Master_»•i
    Master_‹æ•ª_1
    Master_‹æ•ª_2
    Master_’S“–ŽÒ
    Master_‹Æ–±€–Ú

    Echo_ "  ƒ}ƒXƒ^“o˜^ Š®—¹"
End Sub

Private Sub Master_WŒv—ñ()
    ExecSQL "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (3,'\ž•û–@',1)"
    ExecSQL "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (4,'’Š‘IŒ‹‰Ê',2)"
    ExecSQL "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (5,'”[•t‘”­‘—',3)"
    ExecSQL "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (6,'¤•i”­‘—',4)"
    ExecSQL "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (7,'‚»‚Ì‘¼',5)"
    ExecSQL "INSERT INTO [M_WŒv—ñ] ([WŒv—ñID],[WŒv—ñ–¼],[•\Ž¦‡]) VALUES (8,'¤•iŒðŠ·',6)"
End Sub

Private Sub Master_ƒuƒƒbƒN()
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (0,'i‚È‚µj',False,0)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (1,'‹æ•ª',True,1)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (2,'‚»‚Ì‘¼',True,2)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (3,'ŒÚ‹qî•ñ',False,3)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (4,'ƒCƒxƒ“ƒgŠÖŒW',False,4)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (5,'‚»‚Ì‘¼‚Ì‚»‚Ì‘¼‡@',False,5)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (6,'‚»‚Ì‘¼‚Ì‚»‚Ì‘¼‡A',False,6)"
    ExecSQL "INSERT INTO [M_ƒuƒƒbƒN] ([ƒuƒƒbƒNID],[ƒuƒƒbƒN–¼],[»•i•Ê],[•\Ž¦‡]) VALUES (7,'“ÁŽê‚È–â‡‚¹',False,7)"
End Sub

Private Sub Master_»•i()
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (0,0,'i»•iŽw’è‚È‚µj',0,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (1,1,'ƒ~ƒ“ƒg',1,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (2,1,'’Êíƒvƒ‹[ƒt',2,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (3,1,'’Ê”Ni‹L”O“úEƒWƒƒƒpƒ“j',3,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (4,1,'‘—§Œö‰€‹L”O‰Ý',4,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (5,1,'ƒhƒ‰ƒSƒ“ƒvƒ‹[ƒtE‰Ý•¼ƒZƒbƒg',5,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (6,1,'’¹bl•¨‹Y‰æ‰Ý•¼ƒZƒbƒg',6,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (7,1,'‚h‚b‚c‚bƒƒ_ƒ‹',7,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (8,1,'÷‚Ì’Ê‚è”²‚¯ƒvƒ‹[ƒt',8,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (9,1,'÷‚Ì’Ê‚è”²‚¯‰Ý•¼ƒZƒbƒg',9,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (10,1,'‰Ô‚Ì‚Ü‚í‚è‚Ý‚¿‰Ý•¼ƒZƒbƒg',10,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (11,1,'÷‚Ì’Ê‚è”²‚¯‹L”Oƒƒ_ƒ‹',11,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (12,1,'ƒ‹àƒƒ_ƒ‹-¯À-',12,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (13,1,'‘•óÍ”vu’¹bl•¨‹Y‰æv',13,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (14,1,'Š±Žxƒƒ_ƒ‹',14,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (15,1,'ƒAƒWƒA‘å‰ï‹L”O‰Ý',15,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (16,1,'ƒRƒiƒ“ƒvƒ‹[ƒtE‰Ý•¼ƒZƒbƒg',16,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (17,1,'º˜a100”N‹L”O‰Ý•¼',17,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (18,1,'Žµ•óÍ”vu’·•l‰gŽRÕv',18,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (19,1,'i—\”õ19j',19,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (20,1,'i—\”õ20j',20,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (21,1,'i—\”õ21j',21,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (22,1,'i—\”õ22j',22,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (23,1,'i—\”õ23j',23,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (24,2,'ƒIƒŠƒ“ƒsƒbƒN‹L”O‰Ý(‰ß‹Žj',1,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (25,2,'cŽºŠÖŒW‹L”O‰Ý(‰ß‹Ž)',2,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (26,2,'ƒ~ƒ“ƒg',3,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (27,2,'’Êíƒvƒ‹[ƒt',4,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (28,2,'’Ê”Ni‹L”O“úEƒWƒƒƒpƒ“j',5,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (29,2,'¢ŠE•¶‰»ˆâŽYƒZƒbƒg',6,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (30,2,'‚h‚b‚c‚bƒƒ_ƒ‹',7,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (31,2,'Š±Žxƒƒ_ƒ‹',8,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (32,2,'ƒ‹àƒƒ_ƒ‹-¯À-',9,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (33,2,'–œ”Ž‹L”O‰Ý',10,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (34,2,'‘—§Œö‰€‹L”O‰Ý',11,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (35,2,'–œ”Žƒƒ_ƒ‹(‰ß‹Žj',12,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (36,2,'¢ŠE—¤ãƒvƒ‹[ƒt‰Ý•¼ƒZƒbƒg',13,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (37,2,'’¹bl•¨‹Y‰æƒP[ƒX',14,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (38,2,'’n•ûŽ©Ž¡ç‰~‹â‰Ý•¼',15,True)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (39,2,'i—\”õ16j',16,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (40,2,'i—\”õ17j',17,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (41,2,'i—\”õ18j',18,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (42,2,'i—\”õ19j',19,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (43,2,'i—\”õ20j',20,False)"
    ExecSQL "INSERT INTO [M_»•i] ([»•iID],[ƒuƒƒbƒNID],[»•i–¼],[•\Ž¦‡],[—LŒø]) VALUES (44,2,'i—\”õ21j',21,False)"
End Sub

Private Sub Master_‹æ•ª_1()
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (1,1,'\žŠÖŒW',3,Null,1,True,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (2,1,'Žó’',3,Null,2,True,'Žó’')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (3,1,'’Š‘IŒ‹‰Ê',4,Null,3,True,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (4,1,'•¥ž—pŽ†EÄ”­sE‰Â',5,Null,4,True,'•¥ž—pŽ†Ä”­si‰Âj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (5,1,'•¥ž—pŽ†EÄ”­sE•s‰Â',5,Null,5,True,'•¥ž—pŽ†Ä”­si•s‰Âj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (6,1,'•¥ž—pŽ†E“à—e“™EÆ‰ï',5,Null,6,True,'•¥ž—pŽ†“à—eÆ‰ï')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (7,1,'•¥ž—pŽ†E”­‘—ó‹µEÆ‰ï',5,Null,7,True,'•¥ž—pŽ†”­‘—ó‹µ')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (8,1,'“ü‹àŠÖŒW',7,Null,8,True,'“ü‹àŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (9,1,'¤•i”­‘—Æ‰ïE–â‡‚¹',6,Null,9,True,'¤•i”­‘—–â‡‚¹')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (10,1,'¤•i”­‘—Æ‰ïE–¢’…',6,Null,10,True,'¤•i”­‘—Žó—ÌÆ‰ï')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (11,1,'»•iŒðŠ·',8,Null,11,True,'»•iŒðŠ·')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (12,2,'”Ì”„—\’è',7,Null,1,True,'”Ì”„—\’è')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (13,2,'»•i“à—e',7,Null,2,True,'»•i“à—e')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (14,2,'ÝŒÉÆ‰ï',7,Null,3,True,'ÝŒÉÆ‰ï')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (15,2,'‰¿ŠiÆ‰ï',7,Null,4,True,'‰¿ŠiÆ‰ï')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (16,2,'“à—e–â‡‚¹',7,Null,5,False,'“à—e–â‡‚¹')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (17,2,'‘gž‚Ýƒ~ƒXŒðŠ·',7,Null,6,False,'‘gž‚Ýƒ~ƒXŒðŠ·')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (18,2,'‚»‚Ì‘¼ŒðŠ·',7,Null,7,False,'‚»‚Ì‘¼ŒðŠ·')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (19,2,'i–¢Žg—p8j',7,Null,8,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (20,2,'i–¢Žg—p9j',7,Null,9,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (21,2,'i–¢Žg—p10j',7,Null,10,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (22,2,'i–¢Žg—p11j',7,Null,11,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (23,3,'‚c‚l’âŽ~“™EŽ€–S',7,Null,1,True,'‚c‚l’âŽ~(‡@Ž€–S)')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (24,3,'‚c‚l’âŽ~“™E•a‹C“™',7,Null,2,True,'‚c‚l’âŽ~(‡A•a‹C“™)')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (25,3,'‚c‚l’âŽ~“™EŽí—Þ‘½‚¢',7,Null,3,True,'‚c‚l’âŽ~(‡BŽí—Þ‘½‚¢)')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (26,3,'‚c‚l’âŽ~“™E“Á’è¤•i‚Ì‚Ý',7,Null,4,True,'‚c‚l’âŽ~(‡C“Á’è¤•i‚Ì‚Ý)')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (27,3,'‚c‚l’âŽ~“™E——R‚È‚µ',7,Null,5,True,'‚c‚l’âŽ~(‡D——R‚È‚µ)')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (28,3,'‚c‚lÄŠJEˆ¶–¼l•ÏX',7,Null,6,True,'‚c‚lÄŠJEˆ¶–¼l•ÏX')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (29,3,'‚c‚lŠm”F',7,Null,7,True,'‚c‚lŠm”F')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (30,3,'ZŠ•ÏX',7,Null,8,True,'ZŠ•ÏX')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (31,3,'V‹K“o˜^',7,Null,9,True,'V‹K“o˜^')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (32,3,'Ž‘—¿‘—•t',7,Null,10,True,'Ž‘—¿‘—•t')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (33,3,'w“ü—š—ð',7,Null,11,True,'w“ü—š—ð')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (34,4,'‚¨‹à‚ÆØŽè',7,Null,1,True,'ƒCƒxƒ“ƒgŠÖŒWi‚¨‹à‚ÆØŽèj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (35,4,'‰F²_‹{',7,Null,2,True,'ƒCƒxƒ“ƒgŠÖŒWi‘¢•¼‹Ç ‚h‚mj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (36,4,'÷‚Ì’Ê‚è”²‚¯',7,Null,3,True,'ƒCƒxƒ“ƒgŠÖŒWi÷‚Ì’Ê‚è”²‚¯j')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (37,4,'‰Ô‚Ì‚Ü‚í‚è“¹',7,Null,4,True,'ƒCƒxƒ“ƒgŠÖŒWi‚Ü‚í‚è“¹j')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (38,4,'‚s‚h‚b‚b',7,Null,5,True,'ƒCƒxƒ“ƒgŠÖŒWiTICCj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (39,4,'‘‘Ì',7,Null,6,True,'ƒCƒxƒ“ƒgŠÖŒWi‘‘Ìj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (40,4,'‘åãƒRƒCƒ“ƒVƒ‡[',7,Null,7,True,'ƒCƒxƒ“ƒgŠÖŒWi‘åãƒRƒCƒ“j')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (41,4,'‚³‚¢‚½‚ÜƒtƒFƒA',7,Null,8,True,'ƒCƒxƒ“ƒgŠÖŒWi‚³‚¢‚½‚ÜƒtƒFƒAj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (42,4,'²”Œ‹æ–¯‚Ü‚Â‚è',7,Null,9,True,'ƒCƒxƒ“ƒgŠÖŒWi–¼ŒÃ‰®‰Ý•¼‚Ü‚Â‚èj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (43,4,'ŠÏ÷‰ï',7,Null,10,True,'ƒCƒxƒ“ƒgŠÖŒWiŠÏ÷‰ï‚Ì“d˜b“]‘—j')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (44,4,'ƒAƒ“ƒP[ƒg',7,Null,11,True,'ƒCƒxƒ“ƒgŠÖŒWi—\”õj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (45,5,'ƒIƒ“ƒ‰ƒCƒ““o˜^ŠÖŒW',7,Null,1,True,'ƒIƒ“ƒ‰ƒCƒ““o˜^ŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (46,5,'ƒNƒŒƒWƒbƒgŠÖŒW',7,Null,2,True,'ƒNƒŒƒWƒbƒgŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (47,5,'‹àH•i‚fŠÖŒW',7,Null,3,True,'‹àH•i‚f‚Ö“d˜b“]‘—')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (48,5,'L•ñiHêŒ©ŠwŠÜ‚Þj',7,Null,4,True,'L•ñŽºiHêŒ©ŠwŠÜ‚Þj')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (49,5,'ƒ~ƒ“ƒgƒVƒ‡ƒbƒvˆÄ“à',7,Null,5,True,'ƒ~ƒ“ƒgƒVƒ‡ƒbƒvˆÄ“à')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (50,5,'‹àH•i§’k‰ï',7,Null,6,True,'÷ƒTƒ|[ƒg')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (51,5,'—¬’Ê‰Ý•¼ŠÖŒW',7,Null,7,True,'—¬’Ê‰Ý•¼ŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (52,5,'Ž†•¼ŠÖŒW',7,Null,8,True,'Ž†•¼ŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (53,5,'ƒ~ƒ“ƒgƒNƒ‰ƒuŠÖŒW',7,Null,9,True,'ƒ~ƒ“ƒgƒNƒ‰ƒuŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (54,5,'1‰~‚Ê‚¢‚®‚é‚Ý',7,Null,10,True,'1‰~‚Ê‚¢‚®‚é‚Ý')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (55,5,'‹L”O‰ÝˆøŠ·ŠÖŒW',7,Null,11,True,'‹L”O‰ÝˆøŠ·ŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (56,6,'•Ô •i',7,'•Ô•i',1,True,'•Ô•i')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (57,6,'•Ô ‹à',7,'•Ô‹à',2,True,'•Ô‹à')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (58,6,'i–¢Žg—p3j',7,Null,3,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (59,6,'i–¢Žg—p4j',7,Null,4,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (60,6,'i–¢Žg—p5j',7,Null,5,False,Null)"
End Sub

Private Sub Master_‹æ•ª_2()
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (61,6,'‰Ý•¼¾¯Äˆ•ªŠÖŒW',7,Null,6,True,'‰Ý•¼¾¯Äˆ•ªŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (62,6,'‰Ý•¼ôò•û–@',7,Null,7,True,'‰Ý•¼ôò•û–@')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (63,6,'–¼Ì–¢Šm’è»•i—\’è',7,Null,8,True,'–¼Ì–¢Šm’è»•i')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (64,6,'‹ÆŽÒ”Ì”„ŠÖŒW',7,Null,9,True,'‹ÆŽÒ”Ì”„ŠÖŒW')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (65,6,'’Š‘I”{—¦“™(‰ß‹Žj',7,Null,10,True,'’Š‘I”{—¦“™')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (66,6,'i–¢Žg—p11j',7,Null,11,False,Null)"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (67,7,'•Ê“Y‚Ì‚Æ‚¨‚è',7,Null,1,True,'“ÁŽê‚È–â‡‚¹')"
    ExecSQL "INSERT INTO [M_‹æ•ª] ([‹æ•ªID],[ƒuƒƒbƒNID],[‹æ•ª–¼],[WŒv—ñID],[“à–ó‹æ•ª],[•\Ž¦‡],[—LŒø],[‹Œ“]‹L–¼]) VALUES (68,7,'‰º‹L‚Ì‚Æ‚¨‚èi•Ê“Y•s—vj',7,Null,2,True,'“ÁŽê‚È–â‡‚¹')"
End Sub

Private Sub Master_’S“–ŽÒ()
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (1,'001','îà“ˆ','–¼•Û”ü','îà“ˆ –¼•Û”ü',Null,'ƒp[ƒg',Null,Null,1,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (2,'002','–x',Null,'–x ^Ÿ',Null,'ƒp[ƒg',Null,Null,2,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (3,'003','ŽOŒ´','‘½ŒbŽq','ŽOŒ´ ‘½ŒbŽq',Null,'ƒp[ƒg',Null,Null,3,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (4,'004','’JŒû',Null,'’JŒû —F‰ÀŽq',Null,'ƒp[ƒg',Null,Null,4,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (5,'006','‘X',Null,'‘X ‹IŽq',Null,'ƒp[ƒg',Null,Null,5,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (6,'007','ŽRŒû','ŒbŽq','ŽRŒû ŒbŽq',Null,'ƒp[ƒg',Null,Null,6,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (7,'008','ˆ§â',']—¢‰Á','ˆ§â ]—¢‰Á',Null,'ƒp[ƒg',Null,Null,7,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (8,'009','¼‰ª',Null,'¼‰ª —R”üŽq',Null,'ƒp[ƒg',Null,Null,8,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (9,'010','ÎŒ´','—TŽq','ÎŒ´ —TŽq',Null,'ƒp[ƒg',Null,Null,9,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (10,'011','Îˆä','~Žq','Îˆä ~Žq',Null,'ƒp[ƒg',Null,Null,10,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (11,'012','îà‹´','—ŒbŽq','îà‹´ —ŒbŽq',Null,'ƒp[ƒg',Null,Null,11,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (12,'013','—\”õ',Null,'—\”õ',Null,'ƒp[ƒg',Null,Null,12,False,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (13,'014','•Ÿ“‡',Null,'•Ÿ“‡',Null,'ƒp[ƒg',Null,Null,13,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwîà‹´x‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (14,'015','—é–Ø',Null,'—é–Ø',Null,'ƒp[ƒg',Null,Null,14,True,'Œ»s Excel ‚Ì Sheet1!B1 ‚ªwŒÚ‹qGx‚Ì‚Ü‚Ü‚ÅAWŒv•\‚É•Êl‚Æ‚µ‚ÄŒvã‚³‚ê‚Ä‚¢‚½BAccess ‚Å‚Í’S“–ŽÒƒ}ƒXƒ^‚ÅŠÇ—‚·‚éB')"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (15,'016','¬àV','•üŽq','¬àV •üŽq',Null,'ƒp[ƒg',Null,Null,15,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (16,'017','Ž›–{','–@Žq','Ž›–{ –@Žq',Null,'ƒp[ƒg',Null,Null,16,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (17,'018','—é–Ø’q','’qŽq','—é–Ø ’qŽq',Null,'ƒp[ƒg',Null,Null,17,True,Null)"
    ExecSQL "INSERT INTO [M_’S“–ŽÒ] ([’S“–ŽÒID],[’S“–ŽÒƒR[ƒh],[©],[–¼],[Ž–¼],[ƒJƒi],[Eˆõ‹æ•ª],[ÝÐŠJŽn“ú],[ÝÐI—¹“ú],[•\Ž¦‡],[—LŒø],[”õl]) VALUES (18,'100','ŒÚ‹q‚f',Null,'ŒÚ‹qG',Null,'Eˆõ',Null,Null,18,True,Null)"
End Sub

Private Sub Master_‹Æ–±€–Ú()
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (1,'‡@','‡@Žó’“ü—Í','‡@@Žó’“ü—Í',1,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (2,'‡A','‡AŽó’ƒ`ƒFƒbƒN','‡A@Žó’ƒ`ƒFƒbƒN',2,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (3,'‡B','‡B•¥ž‘ƒ`ƒFƒbƒN','‡B@–ß‚è—X•Öˆ—^•¥ž‘ƒ`ƒFƒbƒN',3,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (4,'‡C','‡C–ß‚è—X•Ö','‡C@‚c‚lˆ—^–ß‚è—X•Ö',4,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (5,'‡D','‡D‰Ë“d','‡D@‰Ë“d',5,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (6,'‡E','‡E‚»‚Ì‘¼','‡E@‚»‚Ì‘¼i@@@@@@@@j',6,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (7,'‡F','‡FƒGƒNƒZƒ‹“ü—Í','‡F@ƒGƒNƒZƒ‹“ü—Í',7,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (8,'‡G','‡GƒGƒNƒZƒ‹ƒ`ƒFƒbƒN','‡G@ƒGƒNƒZƒ‹ƒ`ƒFƒbƒN',8,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (9,'‡H','‡H±Ý¹°Ä“ü—Í','‡H@ƒAƒ“ƒP[ƒg“ü—Í',9,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (10,'‡I','‡IÊ¶Þ·ÃÞ°À“ü—Í','‡I@Ê¶Þ·ÃÞ°À“ü—Í',10,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (11,'‡J','‡JÊ¶Þ·ÃÞ°ÀÁª¯¸','‡J@ƒnƒKƒLƒf[ƒ^ƒ`ƒFƒbƒN',11,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (12,'‡K','‡KV‹Kº°ÄÞŽæ‚è','‡K@V‹Kº°ÄÞŽæ‚è',12,True)"
    ExecSQL "INSERT INTO [M_‹Æ–±€–Ú] ([‹Æ–±€–ÚID],[”Ô†],[€–Ú–¼],[’ •[•\Ž¦–¼],[•\Ž¦‡],[—LŒø]) VALUES (13,'‡L','‡LŒÚ‹q®—','‡L@ŒÚ‹q®—',13,True)"
End Sub

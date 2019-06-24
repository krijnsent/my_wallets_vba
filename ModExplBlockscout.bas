Attribute VB_Name = "ModExplBlockscout"
'https://blockscout.com/eth/mainnet/api_docs

Sub TestBlockScout()

'Source: https://github.com/krijnsent/crypto_vba

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExplBlockscout"

'Bunch of test addresses
Dim EthAddressA As String
Dim EthAddressB As String
Dim EthAddressC As String
Dim EthAddressD As String
EthAddressA = "0x0D5b36603eeDE0792d6fdA1Aca78AD7412fE79aa"  '-> simple older address, all types of transactions
EthAddressB = "0x2C0f0D5545ceccC6dA8612E47A75D336031d499E"  '-> simple address, no internal tx
EthAddressC = "0xddbd2b932c763ba5b1b7ae3b362eac3e8d40121a"  '-> 1000d+ address, not too many transactions
EthAddressD = "0x4e83362442b8d1bec281594cea3050c8eb01311c"  '-> 1000+ token transactions

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBlockscoutBasics")

'Error, unknown command
'Put the credentials in a dictionary
Dim Params1 As New Dictionary
TestResult = PublicBlockscout("GET", Params1)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"status":"0","result":null,"message":"Params 'module' and 'action' are required parameters"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("status"), "0"

'Error, parameter missing
Params1.Add "module", "transaction"
Params1.Add "action", "gettxinfo"
TestResult = PublicBlockscout("GET", Params1)
'{"status":"0","result":null,"message":"Query parameter txhash is required"}
Test.IsOk InStr(TestResult, "Query parameter") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "0"
Test.IsEqual JsonResult("message"), "Query parameter txhash is required"

'Latest block number
Dim Params2 As New Dictionary
Params2.Add "module", "block"
Params2.Add "action", "eth_block_number"
TestResult = PublicBlockscout("GET", Params2)
'e.g. {"jsonrpc":"2.0","result":"0x78722b","id":1}
Test.IsOk InStr(TestResult, "jsonrpc") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("jsonrpc"), "2.0"
Test.IsEqual JsonResult("id"), 1

'Test functions
Set Test = Suite.Test("TestBlockscoutBalances")
'Error address
TestResult = GetBlockscoutBalances("BLA")
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(3, 1), "symbol"
Test.IsEqual TestResult(3, 2), "ERROR"
Test.IsEqual TestResult(4, 2), "ETH:Invalid address hash, tokens:Invalid address format"

'OK, with headers
TestResult = GetBlockscoutBalances(EthAddressA)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(2, 2), EthAddressA
Test.IsEqual TestResult(3, 2), "ETH"
Test.IsOk TestResult(4, 2) >= 0

'OK, without headers
TestResult = GetBlockscoutBalances(EthAddressD, False)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 1
Test.IsApproximate TestResult(1, 1), Now(), 10
Test.IsEqual TestResult(2, 1), EthAddressD
Test.IsEqual TestResult(3, 1), "ETH"
Test.IsOk TestResult(4, 1) >= 0

TestResult = GetBlockscoutBalances(EthAddressD, , False, True)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsOk UBound(TestResult, 2) >= 2
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual LCase(TestResult(3, 2)), LCase(EthAddressA)
For rw = 2 To UBound(TestResult, 2)
    Test.IsEqual TestResult(2, rw), EthAddressD
Next rw

TestResult = GetBlockscoutBalances(EthAddressD, , True, False)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsOk UBound(TestResult, 2) = 2
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual LCase(TestResult(3, 2)), LCase(EthAddressA)
Test.IsOk TestResult(4, 2) >= 0

TestResult = GetBlockscoutBalances(EthAddressD, , False, False)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(3, 2), "ERROR"
Test.IsEqual TestResult(4, 2), "NO BALANCES FOUND"

Set Test = Suite.Test("TestBlockscoutTransactions")
'NOT DEVELOPED, blockscout API fails on Internal Transactions
'TestResult = GetBlockscoutTransactions(Address)
'TestResult = GetBlockscoutTransactions(Address, False)

End Sub

Function PublicBlockscout(ReqType As String, ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://blockscout.com/eth/mainnet/api"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
Url = PublicApiSite & MethodParams

PublicBlockscout = WebRequestURL(Url, ReqType)

End Function

Function GetBlockscoutBalances(Address As String, Optional ReturnHeaders As Boolean = True, Optional NormalBal As Boolean = True, Optional TokenBal As Boolean = True) As Variant

Dim JsonResEth As Dictionary
Dim JsonResTok As Dictionary

'headers: timestamp, address, token/eth, amount
ts = Now
cZ = 1
If ReturnHeaders Then
    ReDim ResArr(1 To 4, 1 To cZ)
    ResArr(1, 1) = "timestamp"
    ResArr(2, 1) = "address"
    ResArr(3, 1) = "symbol"
    ResArr(4, 1) = "amount"
    cZ = cZ + 1
End If
        
Dim Params As New Dictionary
Params.Add "module", "account"
Params.Add "action", "balance"
Params.Add "address", Address
If NormalBal Then
    EthResult = PublicBlockscout("GET", Params)
Else
    'Don't load the normal transactions, use a placeholder instead
    EthResult = "{""message"":""OK"",""result"":[]}"
End If

Params("action") = "tokenlist"  'token balances
If TokenBal Then
    TokResult = PublicBlockscout("GET", Params)
Else
    'Don't load the normal transactions, use a placeholder instead
    TokResult = "{""message"":""OK"",""result"":[]}"
End If
Set JsonResEth = JsonConverter.ParseJson(EthResult)
Set JsonResTok = JsonConverter.ParseJson(TokResult)

If JsonResEth("message") = "OK" And JsonResTok("message") = "OK" Then
    If NormalBal = False And TokenBal = False Then
        ReDim Preserve ResArr(1 To 4, 1 To cZ)
        ResArr(1, cZ) = ts
        ResArr(2, cZ) = Address
        ResArr(3, cZ) = "ERROR"
        ResArr(4, cZ) = "NO BALANCES FOUND"
    End If
    
    If NormalBal Then
        ReDim Preserve ResArr(1 To 4, 1 To cZ)
        ResArr(1, cZ) = ts
        ResArr(2, cZ) = Address
        ResArr(3, cZ) = "ETH"
        ResArr(4, cZ) = JsonResEth("result") / 10 ^ 18
        cZ = cZ + 1
    End If
    
    If TokenBal Then
        ResArrTok = JsonToArray(JsonResTok)
        TblTok = ArrayTable(ResArrTok, False)
        ReDim Preserve ResArr(1 To 4, 1 To UBound(TblTok, 2) + cZ - 1)
        For rw = 1 To UBound(TblTok, 2)
            ResArr(1, cZ + rw - 1) = ts
            ResArr(2, cZ + rw - 1) = Address
            If Len(TblTok(3, rw)) > 0 Then
                ResArr(3, cZ + rw - 1) = TblTok(3, rw)
            Else
                ResArr(3, cZ + rw - 1) = Left(TblTok(6, rw), 8) & "?"
            End If
            If TblTok(5, rw) <> "" Then
                ResArr(4, cZ + rw - 1) = TblTok(7, rw) / 10 ^ TblTok(5, rw)
            Else
                ResArr(4, cZ + rw - 1) = TblTok(7, rw) * 1
            End If
        Next rw
    End If
    
    
Else
    'Invalid return
    ReDim Preserve ResArr(1 To 4, 1 To cZ)
    ResArr(1, cZ) = ts
    ResArr(2, cZ) = Address
    ResArr(3, cZ) = "ERROR"
    ResArr(4, cZ) = "ETH:" & JsonResEth("message") & ", tokens:" & JsonResTok("message")
End If

GetBlockscoutBalances = ResArr

End Function

Function GetBlockscoutTransactions(Address As String, Optional ReturnHeaders As Boolean = True, Optional NormalTx As Boolean = True, Optional InternalTx As Boolean = True, Optional TokenTx As Boolean = True) As Variant

Dim JsonResN As Dictionary
Dim JsonResI As Dictionary
Dim JsonResT As Dictionary

'NOT DEVELOPED DUE TO ERROR IN the txlistinternal API call

'module=account&action=txlist&address=0x
'module=account&action=txlistinternal&address=0x -> 20190606 gives ERROR
'module=account&action=tokentx&address=0x

End Function

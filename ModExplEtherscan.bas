Attribute VB_Name = "ModExplEtherscan"
Sub TestEtherscan()

'Source: https://github.com/krijnsent/crypto_vba
'Powered by Etherscan.io APIs -> https://etherscan.io/apis
'Please get a (free) API key from Etherscan.io to use this code

Dim Apikey As String
Apikey = "YourApiKeyToken"

'Remove this line, unless you define 1 constant somewhere ( Public Const apikey_etherscan = "the key to use everywhere" etc )
Apikey = apikey_etherscan

'Bunch of test addresses
Dim EthAddressA As String
Dim EthAddressB As String
Dim EthAddressC As String
Dim EthAddressD As String
Dim EthAddressE As String
EthAddressA = "0x0D5b36603eeDE0792d6fdA1Aca78AD7412fE79aa"  '-> simple older address, all types of transactions
EthAddressB = "0x2C0f0D5545ceccC6dA8612E47A75D336031d499E"  '-> simple address, no internal tx
EthAddressC = "0xddbd2b932c763ba5b1b7ae3b362eac3e8d40121a"  '-> 1000d+ address, not too many transactions
EthAddressD = "0x4e83362442b8d1bec281594cea3050c8eb01311c"  '-> 1000+ token transactions
EthAddressE = "0xB22234F7cFeb779a56B56f075B98A27acb117A31" '-> one tx in, one tx out

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExplEtherscan"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestEtherscanBasics")

'Error, unknown command
'Put the credentials in a dictionary
Dim Params1 As New Dictionary
TestResult = PublicEtherscan("GET", Params1)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"status":"0","result":null,"message":"Params 'module' and 'action' are required parameters"}}
Test.IsOk InStr(LCase(TestResult), "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "0"
Test.IsEqual JsonResult("message"), "NOTOK"
Test.IsEqual JsonResult("result"), "Error! Missing Or invalid Module name"

'Error, parameter missing
Params1.Add "module", "transaction"
Params1.Add "action", "gettxinfo"
TestResult = PublicEtherscan("GET", Params1)
'{"status":"0","result":null,"message":"Query parameter txhash is required"}
Test.IsOk InStr(LCase(TestResult), "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "0"
Test.IsEqual JsonResult("message"), "NOTOK"

'Latest block number
Dim Params2 As New Dictionary
Params2.Add "module", "proxy"
Params2.Add "action", "eth_blockNumber"
Params2.Add "apikey", Apikey
TestResult = PublicEtherscan("GET", Params2)
'e.g. {"jsonrpc":"2.0","result":"0x78722b","id":1}
Test.IsOk InStr(TestResult, "jsonrpc") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("jsonrpc"), "2.0"
Test.IsOk JsonResult("id") > 0

'Test functions
Set Test = Suite.Test("TestEtherscanBalances")
'For token balances, use Blockscout, as Etherscan doesn't have an API call for token balances
'Error address
TestResult = GetEtherscanBalances("BLA", Apikey)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(3, 1), "symbol"
Test.IsEqual TestResult(3, 2), "ERROR"
Test.IsEqual TestResult(4, 2), "Error! Invalid address format"

'OK, with headers
TestResult = GetEtherscanBalances(EthAddressA, Apikey)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(2, 2), EthAddressA
Test.IsEqual TestResult(3, 2), "ETH"
Test.IsOk TestResult(4, 2) >= 0

'OK, without headers
TestResult = GetEtherscanBalances(EthAddressB, Apikey, False)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 1
Test.IsApproximate TestResult(1, 1), Now(), 10
Test.IsEqual TestResult(2, 1), EthAddressB
Test.IsEqual TestResult(3, 1), "ETH"
Test.IsOk TestResult(4, 1) >= 0

Set Test = Suite.Test("TestEtherscanTransactions")
'Error address
TestResult = GetEtherscanTransactions("BLA", Apikey)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(2, 1), "timestamp"
Test.IsApproximate TestResult(2, 2), Now(), 4
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual TestResult(3, 2), "BLA"
Test.IsEqual TestResult(10, 2), "tx:Error! Invalid address format tx_internal:Error! Invalid address format tx_token:Error! Invalid address format"

'OK address, default settings
TestResult = GetEtherscanTransactions(EthAddressA, Apikey)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 2
Test.IsEqual TestResult(2, 1), "timestamp"
Test.IsEqual TestResult(3, 1), "address"
For rw = 2 To UBound(TestResult, 2)
    Test.IsEqual LCase(TestResult(3, rw)), LCase(EthAddressA)
    Test.Includes Array("in", "out"), TestResult(5, rw)
    Test.Includes Array("normal", "int", "token"), TestResult(6, rw)
    Test.Includes Array("tx", "tx_fee"), TestResult(7, rw)
Next rw

'OK, without headers, EthAddressD has many transactions
TestResult = GetEtherscanTransactions(EthAddressD, Apikey, False)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 1000
Test.NotEqual TestResult(2, 1), "timestamp"
Test.IsEqual LCase(TestResult(3, 1)), LCase(EthAddressD)

'No token transactions
TestResult = GetEtherscanTransactions(EthAddressA, Apikey, , True, True, False)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 2
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual LCase(TestResult(3, 2)), LCase(EthAddressA)
For rw = 2 To UBound(TestResult, 2)
    Test.Includes Array("normal", "int"), TestResult(6, rw)
Next rw

'No internal transactions
TestResult = GetEtherscanTransactions(EthAddressA, Apikey, , True, False, True)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 2
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual LCase(TestResult(3, 2)), LCase(EthAddressA)
For rw = 2 To UBound(TestResult, 2)
    Test.Includes Array("normal", "token"), TestResult(6, rw)
Next rw

'No normal transactions
TestResult = GetEtherscanTransactions(EthAddressA, Apikey, , False, True, True)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 2
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual LCase(TestResult(3, 2)), LCase(EthAddressA)
For rw = 2 To UBound(TestResult, 2)
    Test.Includes Array("token", "int"), TestResult(6, rw)
Next rw

'No transactions, empty result
TestResult = GetEtherscanTransactions(EthAddressA, Apikey, , False, False, False)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 2
Test.IsEqual TestResult(3, 1), "address"
Test.IsEqual LCase(TestResult(3, 2)), LCase(EthAddressA)

TestResult = GetEtherscanTransactions(EthAddressE, Apikey)
Test.IsEqual UBound(TestResult, 1), 10
Test.IsOk UBound(TestResult, 2) >= 2

End Sub

Function PublicEtherscan(ReqType As String, ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "http://api.etherscan.io/api"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
Url = PublicApiSite & MethodParams

'Debug.Print Url
PublicEtherscan = WebRequestURL(Url, ReqType)

End Function

Function GetEtherscanBalances(Address As String, Apikey As String, Optional ReturnHeaders As Boolean = True) As Variant

Dim JsonResEth As Dictionary

'headers: timestamp, address, token/eth, amount
ts = Now
If ReturnHeaders Then
    ReDim ResArr(1 To 4, 1 To 1)
    ResArr(1, 1) = "timestamp"
    ResArr(2, 1) = "address"
    ResArr(3, 1) = "symbol"
    ResArr(4, 1) = "amount"
End If

Dim Params1 As New Dictionary
Params1.Add "module", "account"
Params1.Add "action", "balance"
Params1.Add "address", Address
Params1.Add "apikey", Apikey
EthResult = PublicEtherscan("GET", Params1)
Set JsonResEth = JsonConverter.ParseJson(EthResult)
If JsonResEth("message") = "OK" Then
    If ReturnHeaders Then NrRw = 2 Else NrRw = 1
    ReDim Preserve ResArr(1 To 4, 1 To NrRw)
    ResArr(1, NrRw) = ts
    ResArr(2, NrRw) = Address
    ResArr(3, NrRw) = "ETH"
    ResArr(4, NrRw) = JsonResEth("result") / 10 ^ 18
Else
    'Invalid return
    If ReturnHeaders Then NrRw = 2 Else NrRw = 1
    ReDim Preserve ResArr(1 To 4, 1 To NrRw)
    ResArr(1, NrRw) = ts
    ResArr(2, NrRw) = Address
    ResArr(3, NrRw) = "ERROR"
    ResArr(4, NrRw) = JsonResEth("result")
End If

'Etherscan doesn't have an API call for token balances, so none will be returned
GetEtherscanBalances = ResArr

End Function

Function GetEtherscanTransactions(Address As String, Apikey As String, Optional ReturnHeaders As Boolean = True, Optional NormalTx As Boolean = True, Optional InternalTx As Boolean = True, Optional TokenTx As Boolean = True) As Variant

Dim JsonResN As Dictionary
Dim JsonResI As Dictionary
Dim JsonResT As Dictionary

ts = Now
ReDim ResArr(1 To 10, 1 To 1)
If ReturnHeaders Then
    ResArr(1, 1) = "tx_id"
    ResArr(2, 1) = "timestamp"
    ResArr(3, 1) = "address"
    ResArr(4, 1) = "counterparty"
    ResArr(5, 1) = "direction"
    ResArr(6, 1) = "group"
    ResArr(7, 1) = "type"
    ResArr(8, 1) = "amount"
    ResArr(9, 1) = "curr"
    ResArr(10, 1) = "description"
End If

'Get all transactions
Dim Params As New Dictionary
Params.Add "module", "account"
Params.Add "action", "txlist" 'normal transactions
Params.Add "address", Address
Params.Add "apikey", Apikey
Params.Add "sort", "desc"

If NormalTx Then
    NResult = PublicEtherscan("GET", Params)
Else
    'Don't load the normal transactions, use a placeholder instead
    NResult = "{""message"":""OK"",""result"":[]}"
End If

Params("action") = "txlistinternal"  'internal transactions
If InternalTx Then
    IResult = PublicEtherscan("GET", Params)
Else
    'Don't load the normal transactions, use a placeholder instead
    IResult = "{""message"":""OK"",""result"":[]}"
End If

Params("action") = "tokentx"  'token transactions
If TokenTx Then
    TResult = PublicEtherscan("GET", Params)
Else
    'Don't load the normal transactions, use a placeholder instead
    TResult = "{""message"":""OK"",""result"":[]}"
End If

Set JsonResN = JsonConverter.ParseJson(NResult)
Set JsonResI = JsonConverter.ParseJson(IResult)
Set JsonResT = JsonConverter.ParseJson(TResult)

If (JsonResN("message") = "OK" Or JsonResN("message") = "No transactions found") And (JsonResI("message") = "OK" Or JsonResI("message") = "No transactions found") And (JsonResT("message") = "OK" Or JsonResT("message") = "No transactions found") Then
    nN = JsonResN("result").Count
    nI = JsonResI("result").Count
    nT = JsonResT("result").Count
    cN = 1
    cI = 1
    cT = 1
    If ReturnHeaders Then cZ = 2 Else cZ = 1
    If nN + nI + nT = 0 Then
        'No results, add placeholder
        ReDim Preserve ResArr(1 To 10, 1 To cZ)
        ResArr(2, cZ) = ts
        ResArr(3, cZ) = Address
        ResArr(10, cZ) = "NO TRANSACTIONS FOUND"
    End If
    For N = 1 To nN + nI + nT
        If cN <= nN Then tN = JsonResN("result")(cN)("timeStamp") Else tN = 0
        If cI <= nI Then tI = JsonResI("result")(cI)("timeStamp") Else tI = 0
        If cT <= nT Then tT = JsonResT("result")(cT)("timeStamp") Else tT = 0
        
        If tN >= tI And tN >= tT Then
            'JsonResN is most recent
            Set rw = JsonResN("result")(cN)
            ReDim Preserve ResArr(1 To 10, 1 To cZ)
            
            ResArr(1, cZ) = rw("hash")
            ResArr(2, cZ) = UnixTimeToDate(Val(rw("timeStamp")))
            
            If LCase(rw("from")) = LCase(Address) Then
                ReDim Preserve ResArr(1 To 10, 1 To cZ + 1)
                ResArr(3, cZ) = rw("from")
                ResArr(4, cZ) = rw("to")
                ResArr(5, cZ) = "out"
                ResArr(6, cZ) = "normal"
                ResArr(7, cZ) = "tx"
                ResArr(8, cZ) = -1 * rw("value") / 10 ^ 18
                ResArr(9, cZ) = "ETH"
                If rw("isError") = "1" Then
                    ResArr(10, cZ) = "ERROR TRANSACTION"
                End If
                
                'Extra Tx for the fee
                ResArr(1, cZ + 1) = rw("hash")
                ResArr(2, cZ + 1) = UnixTimeToDate(Val(rw("timeStamp")))
                ResArr(3, cZ + 1) = rw("from")
                ResArr(4, cZ + 1) = rw("to")
                ResArr(5, cZ + 1) = "out"
                ResArr(6, cZ + 1) = "normal"
                ResArr(7, cZ + 1) = "tx_fee"
                ResArr(8, cZ + 1) = -1 * rw("gasUsed") * rw("gasPrice") / 10 ^ 18
                ResArr(9, cZ + 1) = "ETH"
                
                'Correct for token transaction
                If cT <= nT And rw("value") = 0 Then
                    If rw("blockNumber") = JsonResT("result")(cT)("blockNumber") Then
                        If rw("nonce") = JsonResT("result")(cT)("nonce") And LCase(JsonResT("result")(cT)("from")) = LCase(Address) Then
                            ResArr(8, cZ + 1) = 0
                            ResArr(10, cZ + 1) = "NOFEE"
                        End If
                    End If
                End If
                cZ = cZ + 2
            Else
                ResArr(3, cZ) = rw("to")
                ResArr(4, cZ) = rw("from")
                ResArr(5, cZ) = "in"
                ResArr(6, cZ) = "normal"
                ResArr(7, cZ) = "tx"
                ResArr(8, cZ) = 1 * rw("value") / 10 ^ 18
                ResArr(9, cZ) = "ETH"
                If rw("isError") = "1" Then
                    ResArr(10, cZ) = "ERROR TRANSACTION"
                End If
                cZ = cZ + 1
            End If
            cN = cN + 1
        ElseIf tI >= tN And tI >= tT Then
            'JsonResI is most recent, internal transaction
            Set rw = JsonResI("result")(cI)
            ReDim Preserve ResArr(1 To 10, 1 To cZ)
            ResArr(1, cZ) = rw("hash")
            ResArr(2, cZ) = UnixTimeToDate(Val(rw("timeStamp")))
                        
            If LCase(rw("from")) = LCase(Address) Then
                ReDim Preserve ResArr(1 To 8, 1 To cZ + 1)
                ResArr(3, cZ) = rw("from")
                ResArr(4, cZ) = rw("to")
                ResArr(5, cZ) = "out"
                ResArr(6, cZ) = "int"
                ResArr(7, cZ) = "tx"
                ResArr(8, cZ) = -1 * rw("value") / 10 ^ 18
                ResArr(9, cZ) = "ETH"
                If rw("isError") = "1" Then
                    ResArr(10, cZ) = "ERROR"
                End If
                
                'Transaction fee
                ResArr(1, cZ + 1) = rw("hash")
                ResArr(2, cZ + 1) = UnixTimeToDate(Val(rw("timeStamp")))
                ResArr(3, cZ + 1) = rw("from")
                ResArr(4, cZ + 1) = rw("to")
                ResArr(5, cZ + 1) = "out"
                ResArr(6, cZ + 1) = "int"
                ResArr(7, cZ + 1) = "tx_fee"
                ResArr(8, cZ + 1) = -1 * rw("gasUsed") * rw("gasPrice") / 10 ^ 18
                ResArr(9, cZ + 1) = "ETH"
                cZ = cZ + 2
            Else
                ResArr(3, cZ) = rw("to")
                ResArr(4, cZ) = rw("from")
                ResArr(5, cZ) = "in"
                ResArr(6, cZ) = "int"
                ResArr(7, cZ) = "tx"
                ResArr(8, cZ) = 1 * rw("value") / 10 ^ 18
                ResArr(9, cZ) = "ETH"
                cZ = cZ + 1
            End If
            cI = cI + 1
        Else
            'JsonResT is most recent, Token transaction
            Set rw = JsonResT("result")(cT)
            tDec = 0
            If Len(rw("tokenDecimal")) > 0 Then tDec = rw("tokenDecimal") * 1
            
            ReDim Preserve ResArr(1 To 10, 1 To cZ)
            ResArr(1, cZ) = rw("hash")
            ResArr(2, cZ) = UnixTimeToDate(Val(rw("timeStamp")))
            
            If LCase(rw("from")) = LCase(Address) Then
                ReDim Preserve ResArr(1 To 10, 1 To cZ + 1)
                ResArr(3, cZ) = rw("from")
                ResArr(4, cZ) = rw("to")
                ResArr(5, cZ) = "out"
                ResArr(6, cZ) = "token"
                ResArr(7, cZ) = "tx"
                ResArr(8, cZ) = -1 * rw("value") / 10 ^ tDec
                ResArr(9, cZ) = rw("tokenSymbol")
                If rw("isError") = "1" Then
                    ResArr(10, cZ) = "ERROR"
                End If
                ResArr(10, cZ) = ResArr(10, cZ) & rw("contractAddress")
                
                'Transaction fee
                ResArr(1, cZ + 1) = rw("hash")
                ResArr(2, cZ + 1) = UnixTimeToDate(Val(rw("timeStamp")))
                ResArr(3, cZ + 1) = rw("from")
                ResArr(4, cZ + 1) = rw("to")
                ResArr(5, cZ + 1) = "out"
                ResArr(6, cZ + 1) = "token"
                ResArr(7, cZ + 1) = "tx_fee"
                ResArr(8, cZ + 1) = -1 * rw("gasUsed") * rw("gasPrice") / 10 ^ 18
                ResArr(9, cZ + 1) = "ETH"
                
                'Weird tokens that have no impact on fees
                If rw("tokenSymbol") = "blockwell.ai KYC Casper Token" Then
                    ResArr(8, cZ + 1) = 0
                    ResArr(10, cZ + 1) = "NO_FEE_CORRECTION"
                End If
                
                cZ = cZ + 2
            Else
                ResArr(3, cZ) = rw("to")
                ResArr(4, cZ) = rw("from")
                ResArr(5, cZ) = "in"
                ResArr(6, cZ) = "token"
                ResArr(7, cZ) = "tx"
                ResArr(8, cZ) = 1 * rw("value") / 10 ^ tDec
                ResArr(9, cZ) = rw("tokenSymbol")
                cZ = cZ + 1
            End If
            'Debug.Print N & " T " & tT
            cT = cT + 1
        End If
    Next N
Else
    If ReturnHeaders Then cZ = 2 Else cZ = 1
    ReDim Preserve ResArr(1 To 10, 1 To cZ)
    ResArr(2, cZ) = ts
    ResArr(3, cZ) = Address
    
    ResArr(10, cZ) = "tx:"
    If JsonResN("message") = "OK" Or JsonResN("message") = "No transactions found" Then
        ResArr(10, cZ) = ResArr(10, cZ) & JsonResN("result").Count
    Else
        ResArr(10, cZ) = ResArr(10, cZ) & JsonResN("result")
    End If
    
    ResArr(10, cZ) = ResArr(10, cZ) & " tx_internal:"
    If JsonResI("message") = "OK" Or JsonResI("message") = "No transactions found" Then
        ResArr(10, cZ) = ResArr(10, cZ) & JsonResI("result").Count
    Else
        ResArr(10, cZ) = ResArr(10, cZ) & JsonResI("result")
    End If
    
    ResArr(10, cZ) = ResArr(10, cZ) & " tx_token:"
    If JsonResT("message") = "OK" Or JsonResT("message") = "No transactions found" Then
        ResArr(10, cZ) = ResArr(10, cZ) & JsonResT("result").Count
    Else
        ResArr(10, cZ) = ResArr(10, cZ) & JsonResT("result")
    End If
End If

GetEtherscanTransactions = ResArr

End Function


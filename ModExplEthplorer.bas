Attribute VB_Name = "ModExplEthplorer"
'https://github.com/EverexIO/Ethplorer/wiki/Ethplorer-API#get-last-block

Sub TestEthplorer()

'Source: https://github.com/krijnsent/crypto_vba
'Powered by Ethplorer
'Please get a (free) API key from Ethplorer

Dim Apikey As String
Apikey = "freekey" 'good enough for simple request, please get a free key from ethplorer if you want to use it more often

'Remove this line, unless you define 1 constant somewhere ( Public Const apikey_etherscan = "the key to use everywhere" etc )
Apikey = apikey_ethplorer

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
Suite.Description = "ModExplEthplorer"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestEthplorerBasics")

'Error, unknown command
'Put the credentials in a dictionary, start with an empty one
Dim Params1 As New Dictionary
TestResult = PublicEthplorer("", "GET", Params1)
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"error":{"code":17,"message":"Invalid request, check API manual there: https://github.com/EverexIO/Ethplorer/wiki/Ethplorer-API"}}}
Test.IsOk InStr(LCase(TestResult), "error") > 0, "test error 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test error 2 failed, result: ${1}"
Test.IsEqual JsonResult("response_txt")("error")("code"), 17, "test error 3 failed, result: ${1}"

'Simple request, fails because it has no apiKey
TestResult = PublicEthplorer("getLastBlock", "GET", Params1)
'{"error_nr":401,"error_txt":"HTTP-Unauthorized","response_txt":{"error":{"code":1,"message":"Invalid API key"}}}
Test.IsOk InStr(LCase(TestResult), "error") > 0, "test error 4 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("response_txt")("error")("message"), "Invalid API key", "test error 5 failed, result: ${1}"

Params1.Add "apiKey", Apikey
TestResult = PublicEthplorer("getLastBlock", "GET", Params1)
'{"lastBlock":11824478}
Test.IsOk InStr(LCase(TestResult), "lastblock") > 0, "test block 1, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("lastBlock") > 1000, "test block 2, result: ${1}"

Set Test = Suite.Test("TestEthplorerBalances")
'For token balances, use Ethplorer, as Etherscan doesn't have an API call for token balances

TestResult = GetEthplorerBalances(EthAddressB, Apikey)
'Debug.Print TestResult
Test.IsEqual UBound(TestResult, 1), 4, "GetBalance, error 1, result: ${1}"
Test.IsEqual UBound(TestResult, 2), 6, "GetBalance, error 2, result: ${1}"
Test.IsEqual TestResult(3, 1), "symbol", "GetBalance, error 3, result: ${1}"
Test.IsEqual TestResult(3, 2), "ETH", "GetBalance, error 4, result: ${1}"
Test.IsEqual TestResult(4, 2), 0.000934, "GetBalance, error 5, result: ${1}"
Test.IsEqual TestResult(3, 4), "TFL", "GetBalance, error 6, result: ${1}"

End Sub

Function PublicEthplorer(Method As String, ReqType As String, ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.ethplorer.io/"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
url = PublicApiSite & Method & MethodParams

Debug.Print url

PublicEthplorer = WebRequestURL(url, ReqType)

End Function

Function GetEthplorerBalances(Address As String, Apikey As String, Optional ReturnHeaders As Boolean = True, Optional NormalBal As Boolean = True, Optional TokenBal As Boolean = True) As Variant

Dim JsonResEth As Dictionary
Dim NrRw As Integer

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
Params1.Add "apiKey", Apikey
EthResult = PublicEthplorer("getAddressInfo/" & Address, "GET", Params1)
Set JsonResEth = JsonConverter.ParseJson(EthResult)
If JsonResEth("address") = LCase(Address) Then
    If NormalBal Then
        If ReturnHeaders Then NrRw = 2 Else NrRw = 1
        ReDim Preserve ResArr(1 To 4, 1 To NrRw)
        ResArr(1, NrRw) = ts
        ResArr(2, NrRw) = Address
        ResArr(3, NrRw) = "ETH"
        ResArr(4, NrRw) = JsonResEth("ETH")("balance")
    End If
    
    If TokenBal Then
        For Each tok In JsonResEth("tokens")
            NrRw = NrRw + 1
            ReDim Preserve ResArr(1 To 4, 1 To NrRw)
            ResArr(1, NrRw) = ts
            ResArr(2, NrRw) = Address
            ResArr(3, NrRw) = tok("tokenInfo")("symbol")
            ResArr(4, NrRw) = Val(tok("balance")) / (10 ^ Val(tok("tokenInfo")("decimals")))
        Next tok
    End If
    
Else
    'Invalid return
    If ReturnHeaders Then NrRw = 2 Else NrRw = 1
    ReDim Preserve ResArr(1 To 4, 1 To NrRw)
    ResArr(1, NrRw) = ts
    ResArr(2, NrRw) = Address
    ResArr(3, NrRw) = "ERROR"
    ResArr(4, NrRw) = JsonResEth("result")
End If

GetEthplorerBalances = ResArr

End Function

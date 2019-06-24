# my_wallets_vba
An Excel/VBA project to connect to various cryptocurrency wallets through APIs. This project is related to crypto_vba (connection to crypto exchanges), using part of that code.

# APIs:
Get information from:
- [Ethersscan](https://etherscan.io)
- [Blockscout](https://blockscout.com/)

The API response is pure JSON, for which I included https://github.com/VBA-tools/VBA-JSON to process and a function to build on that.
Please consider the code I provide as simple building blocks: if you want to build a project based on this code, you will have to know (some) VBA. There are plenty of courses available online, two simple ones I send starters to are: https://www.excel-pratique.com/en/ and https://homeandlearn.org/.

# How to use?
Import the .bas files you need or simply take the sample Excel file. In the modules you'll find some examples how to use the code. Feel free to create an issue if things don't work for you. The project uses quite some Dictionaries in VBA, check out e.g. https://excelmacromastery.com/vba-dictionary/ if you want to know a bit more about them.

# Functions:
- Get Balances -> using Blockscout, as they do provide both ETH and token balances
- Get transactions -> using Etherscan, as they are the only ones providing normal, internal & token transactions. 
  - The code is processing the raw returned data, so it can easily be used in a pivot table.
  - If you switch on all 3 types of transactions, the ETH total should add up to your balance

# ToDo
- Other coins, suggestions are welcome as an issue in this project. On the wish list are currently BTC, EOS and XLM.
- Better error handling

# Donate
If this project/the Excel saves you a lot of programming time, consider sending me a coffee or a beer:<br/>
ETH (or ERC-20 tokens): 0x9070C5D93ADb58B8cc0b281051710CB67a40C72B<br/>
<b>Cheers!</b>
# AutoDescription
## Program name: AutoDescription
```利用客製程式讀取Excel，根據Excel組成及規則，啟動按鈕後，自動填入Items Title Block頁籤中的Description。```

### Function
1. readExcel
>   讀Excel，每一列(row)的subclass對應到系統，再對應每一欄(column)的組成，按照符號規則呼叫不同function填入字串。
2. getString_DollarSigns
>   處理Excel cell包含\$的字串。
3. getString_NoSigns
>   處理Excel cell沒有特殊符號的字串。
4. getString_Asterisk
>   處理Excel cell包含\*的字串。

### 符號
| Sign  | Meaning  | 備註 |
| :------------: |:---------------| :-----|
| 無符號      | text/list | 其他類型 |
| \$      | cell        |  最前面 |
| \* | 檢查        |   最前面最後面 |
| \! | 空格        |   最後面 |

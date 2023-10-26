# Iniクラス

INIファイルを読み取るクラス   
VBAではよく`GetPrivateProfileString`を使うという解説がされているがこれは
- 処理が遅い
- 定義を宣言しないといけない
- Declareを使う(オフィスの64bitと32bitの使い分けが必要)
- 特定の文字コードでしか使えない  

等、不便でしかないので自分でテキストを読み込んで使う  
シンプルに使うため読み込みのみ対応  

---

Class to read INI file.  
In VBA, it is often explained that `GetPrivateProfileString` is used, but this  
- slow processing  
- Must declare definition  
- Use Declare  
- Can only be used with specific character codes  

it's just an inconvenience, so I read the text myself and use it.  
Read-only support for simple use!  

---

Example  
```vb
Dim value As String
With New Ini
    .Init "ini file path", cceShift_jis
    value = .GetValue("Section", "Parameter")
End With
```

## 目次

- [Iniクラス](#iniクラス)
  - [目次](#目次)
  - [メンバ](#メンバ)
    - [Init](#init)
    - [GetValue](#getvalue)

## メンバ  

### Init
コンストラクタ(newした後、呼び出す)  

#### Syntax  
```vb  
Function Init( _
        ByVal FilePath As String, _
        Optional ByVal CharacterCode As CharacterCodeEnum, _
        Optional ByVal NewlineCharacter As NewlineCharacterEnum _
) As Ini
```  

#### Parameters  
**FilePath**  
iniファイルのフルパス  

**CharacterCode**  
文字コード。既定値は"UTF-8"  

CharacterCodeEnum列挙
| メンバ | 内容 |
|:---|:---|
|cceUtf8|UTF-8|
|cceShift_jis|Shift_JIS|
|cceAscii|ASCII|

**NewlineCharacter**  
改行コード。既定値は"CRLF"

NewlineCharacterEnum列挙
| メンバ | 内容 |
|:---|:---|
|nceCRLF|CRLF|
|nceLF|LF|

#### Return Value  
コンストラクタした自分自身  

#### Remarks  
もしファイルが見つからない場合"Nothing"を返します。  

---

### GetValue  
値を取得する。    

#### Syntax  
```vb  
Function GetValue( _
        ByVal Section As String, _
        ByVal Parameter As String _
) As String
```  

#### Parameters  
**Section**  
セクション指定  

**Parameter**  
パラメータ名指定  

#### Return Value  
値  

#### Remarks  
存在しない場合は空文字""を返します。  


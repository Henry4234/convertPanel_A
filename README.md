# convert Panel_A
把Panel A的PDF檔匯入既有的xlsm中

範例檔的xlsm如下：
[不規則抗體.zip](https://github.com/Henry4234/convert-Panel_A/files/11009534/default.zip)

而Panel_A的PDF範例檔可以在官網中下載，這裡也附上其中一個批號的範例檔：
[RA209.pdf](https://github.com/Henry4234/convert-Panel_A/files/11009537/RA209.pdf)

利用python中的pdfplumber套件，把Panel_A的批號(Lot No.)及整個表格(antigrams)讀取出來

`pdf = pdfplumber.open(path1)`
`page0 = pdf.pages[0]`
`table = page0.extract_table()`
`pagetext = str(page0.extract_tex`
 
利用pandas整理過後，將不必要的資訊刪除

`df = pd.DataFrame(rawdata)`
`df.columns = df.iloc[0]`
`df = df[1:]`
`df["Jsa"] = df["Jsa"].replace("/","?")`
`df["P1"] = df["P1"].replace("+s","+")`
`df = df.replace("0","-")`
 
最後利用openpyxl將整理過後的資料貼上另存新檔

`wb = load_workbook(filename=path2, read_only=False, keep_vba=True)`
`temp = wb["temp"]`

前端是用tkinter搭配ttk。

實際成品如下：


![螢幕擷取畫面 2023-03-19 045240](https://user-images.githubusercontent.com/102476562/226139137-5a3829aa-0eb8-448c-8c79-25a31e77848f.png)


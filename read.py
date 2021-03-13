from openpyxl import load_workbook

wb = load_workbook(filename = 'demo.xlsx')
# シート名、大文字小目区別される
ws = wb['Sheet1']

for row in ws.rows:
    print(row)
    for col in row:
        print(col, col.value, type(col.value))

# 1.25は float型
# 1.9999999999は、Excel上の表示は2だが、値が取得できている(float型)
# 文字列として取得する、取得処理をカスタマイズする方法は提供されていないようだ
# 100は int型
# セルの書式で 1.000も、値としては1 (int型)が取得できている
# 2021/03/13はdatetime.datetime型
# 100a 、かきく500 はどちらもstr型
# 日本語も特に文字化けなし

# Excelの１行目をヘッダー行とすれば、その行の各列の値から意味を決定することもできそう


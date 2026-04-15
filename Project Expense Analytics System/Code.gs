=QUERY('支出流水帳共筆（當月在最上面）'!A:H, "select E, sum(H) where E is not null and E != '類別標題1' and E != '專案' group by E label E '類別', sum(H) '總金額'", 1)
//公式一
=QUERY(INDIRECT("'支出流水帳共筆（當月在最上面）'!A2:R"), "select Col5, sum(Col8) where Col5 is not null and Col18 is not null group by Col5 pivot Col18 label Col5 '類別/會計月'", 1)
//公式二

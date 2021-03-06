SELECT 郵便番号,
       [都道府県名-カナ] & [市区町村名-カナ] & [町域名-カナ] AS [住所-カナ], 
       [都道府県名] & [市区町村名] & [町域名] AS 住所,
       更新の表示,
       変更理由,
       Date() AS 登録年月日
FROM tblUpdateData
WHERE 更新の表示=1;

DELETE *
FROM tblAddressMaster
WHERE 郵便番号 IN
    (SELECT 郵便番号
     FROM selectDelData)
  AND 住所 IN
    (SELECT 住所
     FROM selectDelData);


SELECT tblUpdateData.郵便番号,
       tblUpdateData.[住所-カナ],
       tblUpdateData.住所,
       tblUpdateData.更新の表示,
       tblUpdateData.変更理由,
       tblUpdateData.登録年月日
FROM tblUpdateData
WHERE 更新の表示=2;


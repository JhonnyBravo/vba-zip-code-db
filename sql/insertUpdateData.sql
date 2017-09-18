PARAMETERS ID Long, zipCode Text (7),
                    phonetic Text (255),
                    address Text (255),
                    updateCode Short,
                    reasonCode Short,
                    resistrationDate DateTime;


INSERT INTO tblUpdateData
VALUES (ID,
        zipCode,
        phonetic,
        address,
        updateCode,
        reasonCode,
        resistrationDate);


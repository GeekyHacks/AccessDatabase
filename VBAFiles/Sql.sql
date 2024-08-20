SELECT T1.Airport,
    T1.MinVendorPrice,
    T1.FlightType,
    T2.TableName,
    T2.EstimatedTotalCol,
    T2.ID,
    T2.LOWCURRENCY
FROM (
        SELECT Airport,
            FlightType,
            MIN(VendorPrice) AS MinVendorPrice
        FROM (
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'AEGFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM AEGFuelT
                UNION ALL
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'ASMFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM ASMFuelT
                UNION ALL
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'UCIGFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM UCIGFuelT
                UNION ALL
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'LINKAEROFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM LINKAEROFuelT
                UNION ALL
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'HadidFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM HadidFuelT
                UNION ALL
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'MASFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM MASFuelT
                UNION ALL
                SELECT Airport,
                    VendorPrice,
                    FlightType,
                    'YPCFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM YPCFuelT
            ) AS AllVendorPrices
        GROUP BY Airport,
            FlightType
        HAVING COUNT(*) >= 1
    ) AS T1
    INNER JOIN (
        SELECT ID,
            Airport,
            FlightType,
            VendorPrice,
            TableName,
            LOWCURRENCY,
         EstimatedTotalCol
        FROM (
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'AEGFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM AEGFuelT
                UNION ALL
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'ASMFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM ASMFuelT
                UNION ALL
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'YPCFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM YPCFuelT
                UNION ALL
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'UCIGFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM UCIGFuelT
                UNION ALL
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'LINKAEROFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM LINKAEROFuelT
                UNION ALL
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'MASFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM MASFuelT
                UNION ALL
                SELECT ID,
                    Airport,
                    FlightType,
                    VendorPrice,
                    'HadidFuelT' AS TableName,
                    Currency AS LOWCURRENCY,
                    EstimatedTotal AS EstimatedTotalCol
                FROM HadidFuelT
            ) AS AllVendorPrices
    ) AS T2 ON (T1.MinVendorPrice = T2.VendorPrice)
    AND (T1.Airport = T2.Airport)
    AND (T1.FlightType = T2.FlightType);
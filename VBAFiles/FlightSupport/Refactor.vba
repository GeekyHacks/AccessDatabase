SELECT DISTINCT T1.Airport, T1.MinVendorPrice, T1.FlightType, T2.TableName, T2.EstimatedTotalCol, T2.ID, T2.LOWCURRENCY, IIf(
        [T2].[VendorPrice] > [T2].[EstimatedTotalCol],
        [T2].[VendorPrice],
        [T2].[EstimatedTotalCol]
    ) AS MinPrice
FROM (SELECT Airport, FlightType, MIN(VendorPrice) AS MinVendorPrice FROM (SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'AEGFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM AEGFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'ASMFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM ASMFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'UCIGFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM UCIGFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'LINKAEROFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM LINKAEROFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'HadidFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM HadidFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'MASFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM MASFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'YPCFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM YPCFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'MixJetFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM MixJetFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'AviaryFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM AviaryFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'AurroraFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM AurroraFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'JetexFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM JetexFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'HonestyAirFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM HonestyAirFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'WoezonFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM WoezonFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'ParadiseFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM ParadiseFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'GACFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM GACFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'JBSFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM JBSFuelT
        UNION ALL
        SELECT 
            Airport,
            VendorPrice,
            FlightType,
            'ProJetFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM ProJetFuelT
    )  AS AllVendorPrices GROUP BY Airport, FlightType HAVING COUNT(*) >= 1)  AS T1 INNER JOIN (SELECT ID, Airport, FlightType, VendorPrice, TableName, LOWCURRENCY, EstimatedTotalCol FROM (SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'AEGFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM AEGFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'ASMFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM ASMFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'YPCFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM YPCFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'UCIGFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM UCIGFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'LINKAEROFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM LINKAEROFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'HadidFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM HadidFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'MASFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM MASFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'MixJetFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM MixJetFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'AviaryFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM AviaryFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'AurroraFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM AurroraFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'JetexFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM JetexFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'HonestyAirFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM HonestyAirFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'WoezonFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM WoezonFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'ParadiseFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM ParadiseFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'GACFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM GACFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'JBSFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM JBSFuelT
        UNION ALL
        SELECT 
            ID,
            Airport,
            FlightType,
            VendorPrice,
            'ProJetFuelT' AS TableName,
            Currency AS LOWCURRENCY,
            EstimatedTotal AS EstimatedTotalCol
        FROM ProJetFuelT
    )  AS AllVendorDetails)  AS T2 ON (T1.MinVendorPrice = T2.VendorPrice) AND (T1.FlightType = T2.FlightType) AND (T1.Airport = T2.Airport);

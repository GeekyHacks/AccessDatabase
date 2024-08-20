SELECT IIf(
        T1.Airport = 'Dest',
        'FlightSupport_DestinationT',
        'FlightSupport_OriginT'
    ) AS SourceTable,
    T1.Airport,
    CombinedLocations.Loc_ID,
    CombinedLocations.Location,
    CombinedLocations.Sect_ID,
    CombinedLocations.AirportName,
    FlightSupport_SectorsT_New_V12.FlightRefNo
FROM FlightSupport_SectorsT_New_V12
    INNER JOIN (
        (
            SELECT 'Dest' AS Airport,
                Dest_Location_ID AS Loc_ID,
                DestAirportCode AS Location,
                Dest_Sect_ID AS Sect_ID,
                Dest_AirportName AS AirportName
            FROM FlightSupport_DestinationT
            UNION ALL
            SELECT 'Origin' AS Airport,
                Origin_Location_ID AS Loc_ID,
                OriginAirportCode AS Location,
                Origin_Sect_ID AS Sect_ID,
                Origin_AirportName AS AirportName
            FROM FlightSupport_OriginT
        ) AS T1
        INNER JOIN (
            SELECT Origin_Location_ID AS Loc_ID,
                OriginAirportCode AS Location,
                Origin_Sect_ID AS Sect_ID,
                Origin_AirportName AS AirportName
            FROM FlightSupport_OriginT
            UNION ALL
            SELECT Dest_Location_ID AS Loc_ID,
                DestAirportCode AS Location,
                Dest_Sect_ID AS Sect_ID,
                Dest_AirportName AS AirportName
            FROM FlightSupport_DestinationT
        ) AS CombinedLocations ON T1.Loc_ID = CombinedLocations.Loc_ID
    ) ON FlightSupport_SectorsT_New_V12.Sect_ID = T1.Sect_ID
ORDER BY CombinedLocations.Sect_ID DESC;



SELECT IIF(
        T1.Airport = 'Dest',
        'FlightSupport_DestinationT',
        'FlightSupport_OriginT'
    ) AS SourceTable,
    T1.Airport,
    T1.Loc_ID,
    T1.Sector,
    T1.Location,
    T1.Sect_ID,
    T1.AirportName,
    FlightSupport_SectorsT_New_V12.FlightRefNo
FROM (
        SELECT 'Dest' AS Airport,
            Dest_Location_ID AS Loc_ID,
            DestAirportCode AS Location,
            Dest_Sect_ID AS Sect_ID,
            SectorNo AS Sector,
            Dest_AirportName AS AirportName
        FROM FlightSupport_DestinationT
        UNION ALL
        SELECT 'Origin' AS Airport,
            Origin_Location_ID AS Loc_ID,
            OriginAirportCode AS Location,
            Origin_Sect_ID AS Sect_ID,
            SectorNo AS Sector,
            Origin_AirportName AS AirportName
        FROM FlightSupport_OriginT
    ) AS T1
    INNER JOIN FlightSupport_SectorsT_New_V12 ON T1.Sect_ID = FlightSupport_SectorsT_New_V12.Sect_ID
ORDER BY T1.Sect_ID DESC;
let ExtractDoublon = (pTable as any, pKey as text) as table =>
    let
        Source = Excel.CurrentWorkbook(){[Name=pTable]}[Content],

        // Key for regroup values
        KeyGroup =pKey,

        // Temp columns for computing
        IndexName="도Index도",
        ColExpand1="도Duplicates도",
        ColExpand2="도Detail1도",
        ColExpand3="도Detail2도",
        
        // Add Index 
        ADD_INDEX = Table.AddIndexColumn(Source, IndexName, 0, 1),    

        // Groups by key parameter
        GROUPING_ROWS = Table.Group(ADD_INDEX, KeyGroup, {{ColExpand1, each Table.RowCount(_), type number}, {ColExpand2, each _, type table}}),    

        // Tips to get doublons
        SUBTRACT = Table.TransformColumns(GROUPING_ROWS, {{ColExpand1, each _ - 1, type number}}),    

        // Expand results table
        EXTEND = Table.ExpandTableColumn(SUBTRACT, ColExpand2, {ColExpand3, IndexName}, {ColExpand3, IndexName}),

        // Sort and fixe doublons
        SORT = Table.Sort(EXTEND,{{IndexName, Order.Ascending}}),
        DOUBLONS_FIXED = Table.RemoveColumns(SORT,{IndexName}),    
        EXTRACT_1 = Table.SelectRows(DOUBLONS_FIXED, each ([도Duplicates도] >= 1)),    

        // Get pKey column only
        DELETE_COLUMNS = Table.RemoveColumns(EXTRACT_1,{ColExpand1, ColExpand3}),
        DOUBLONS_ENDED = Table.Distinct(DELETE_COLUMNS, {KeyGroup})

    in
        DOUBLONS_ENDED
in
    ExtractDoublon
	
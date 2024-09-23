Attribute VB_Name = "test"



Sub ManageFirstITSheet()

    Dim ws As Worksheet
    
    
  
    Dim searchStrings As Variant
    Dim searchString As Variant
   
    

    
    
    
    ' handel subject that have pass or fail only
    
    ' searchStrings = Array("—«”»") ' Add more substrings as needed
    
    
    ' Flag = False

    
    Set ws = ThisWorkbook.Sheets("Sheet1")

    '///////////////////////////////////////////////////////////////////////////
    '/////////////////////////// Remove the Circle of all cells ///////////////
    '/////////////////////////////////////////////////////////////////////////
    RemoveShapesFromCells "H11:H427,L11:L427,P11:P427,T11:T427,X11:X427,AB11:AB427,AF11:AF427,AJ11:AJ427,AN11:AN427,AR11:AR427,AV11:AV427,AZ11:AZ427"
    ' searchStrings = Array("„", "Ã‹ Ã‹", "Ã‹ //","·","‰«ÃÕ") ' Add more substrings as needed
    ' UnblockAllCellsNotUsed ws,"E11:AZ427",searchStrings ,"123456"

    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "H11:H427","IT Essentials", 0
    CalculateStatusStudentsOFSubjectBefore "L11:L427","Technical English I", 15
    CalculateStatusStudentsOFSubjectBefore "P11:P427","Intro To Cyber Security", 30
    CalculateStatusStudentsOFSubjectBefore "T11:T427","Mathematics I", 45
    CalculateStatusStudentsOFSubjectBefore "X11:X427","Physics", 60
    CalculateStatusStudentsOFSubjectBefore "AB11:AB427","Programming Essentials in python", 75
    CalculateStatusStudentsOFSubjectBefore "AF11:AF427","Programming essentials in c", 90
    CalculateStatusStudentsOFSubjectBefore "AJ11:AJ427","Cyber security essentials", 105
    CalculateStatusStudentsOFSubjectBefore "AN11:AN427","Into to iot &in connecting things", 120
    CalculateStatusStudentsOFSubjectBefore "AR11:AR427","Ms office", 135
    CalculateStatusStudentsOFSubjectBefore "AV11:AV427","Technical english II", 150
    CalculateStatusStudentsOFSubjectBefore "AZ11:AZ427","mathematics II", 165
    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")
    HndelAbsentSubjectHaveFailedOrBassTotalScore_150 ws, "E11:H427,I11:L427", searchStrings
    HndelAbsentSubjectHaveTotalScore_50_Normal ws, "M11:P427", searchStrings
    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "Q11:AF427,AK11:AN427,AS11:AZ427", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "AG11:AJ427,AO11:AR427", searchStrings


    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("—«”»") ' Add more substrings as needed
    HndelSubjectHaveFailedOrBassTotalScore_150 ws, "E11:H427,I11:L427", searchStrings
    
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    HandelSubjectHaveTotalScore_50_Normal ws, "M11:P427", searchStrings
    HandelSubjectHaveTotalScore_150_Normal ws, "Q11:AF427,AK11:AN427,AS11:AZ427", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "AG11:AJ427,AO11:AR427", searchStrings



    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change ///////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter "H11:H427","IT Essentials", 0
    CalculateStatusStudentsOFSubjectAfter "L11:L427","Technical English I", 15
    CalculateStatusStudentsOFSubjectAfter "P11:P427","Intro To Cyber Security", 30
    CalculateStatusStudentsOFSubjectAfter "T11:T427","Mathematics I", 45
    CalculateStatusStudentsOFSubjectAfter "X11:X427","Physics", 60
    CalculateStatusStudentsOFSubjectAfter "AB11:AB427","Programming Essentials in python", 75
    CalculateStatusStudentsOFSubjectAfter "AF11:AF427","Programming essentials in c", 90
    CalculateStatusStudentsOFSubjectAfter "AJ11:AJ427","Cyber security essentials", 105
    CalculateStatusStudentsOFSubjectAfter "AN11:AN427","Into to iot &in connecting things", 120
    CalculateStatusStudentsOFSubjectAfter "AR11:AR427","Ms office", 135
    CalculateStatusStudentsOFSubjectAfter "AV11:AV427","Technical english II", 150
    CalculateStatusStudentsOFSubjectAfter "AZ11:AZ427","mathematics II", 165


    '///////////////////////////////////////////////////////////////////////////
    '////////////////////////// Block All Unused Cell /////////////////////////
    '/////////////////////////////////////////////////////////////////////////
    'searchStrings = Array("„", "Ã‹ Ã‹", "Ã‹ //","·","‰«ÃÕ") ' Add more substrings as needed
    'BlockAllCellsNotUsed ws,"H11:H427,L11:L427,P11:P427,T11:T427,X11:X427,AB11:AB427,AF11:AF427,AJ11:AJ427,AN11:AN427,AR11:AR427,AV11:AV427,AZ11:AZ427",searchStrings,"123456"
    
    


    
   
End Sub

Sub ManageSecondITSheet()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("Sheet1")    
    ' searchStrings = Array("„", "Ã‹ Ã‹", "Ã‹ //","·","‰«ÃÕ") ' Add more substrings as needed
    ' UnblockAllCellsNotUsed ws,"E11:AZ427",searchStrings ,"123456"

    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "G14:G236","Linux Essentials", 0
    CalculateStatusStudentsOFSubjectBefore "K14:K236","programming Essentials in c ++", 15
    CalculateStatusStudentsOFSubjectBefore "O14:O236","Web Programming l", 30
    CalculateStatusStudentsOFSubjectBefore "S14:S236","Introduction to DB ", 45
    CalculateStatusStudentsOFSubjectBefore "W14:W236","Digital Engineering", 60
    CalculateStatusStudentsOFSubjectBefore "AA11:AA236","Operating System", 75
    CalculateStatusStudentsOFSubjectBefore "AE11:AE236","Web Programming II", 90
    CalculateStatusStudentsOFSubjectBefore "AI11:AI236","Database Programming", 105
    CalculateStatusStudentsOFSubjectBefore "AM11:AM236","Data Structure", 120
    CalculateStatusStudentsOFSubjectBefore "AQ11:AQ236","CCNA R&S I", 135
    CalculateStatusStudentsOFSubjectBefore "AU11:AU236","Java Programming I", 150
    CalculateStatusStudentsOFSubjectBefore "AY11:AY236","Capstone Design", 165
    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")

    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "D14:O236,T14:W236,AB14:AE236,AJ14:AU236", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "P14:S236,X14:AA236,AF14:AI236,AV14:AY236", searchStrings

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    
    HandelSubjectHaveTotalScore_150_Normal ws, "D14:O236,T14:W236,AB14:AE236,AJ14:AU236", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "P14:S236,X14:AA236,AF14:AI236,AV14:AY236", searchStrings
    '///////////////////////////////////////////////////////////////////////////
    '/////////////////////////// Remove the Circle of all cells ///////////////
    '/////////////////////////////////////////////////////////////////////////
    RemoveShapesFromCells "D14:AY236"
    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter "G14:G236","Linux Essentials", 0
    CalculateStatusStudentsOFSubjectAfter "K14:K236","programming Essentials in c ++", 15
    CalculateStatusStudentsOFSubjectAfter "O14:O236","Web Programming l", 30
    CalculateStatusStudentsOFSubjectAfter "S14:S236","Introduction to DB ", 45
    CalculateStatusStudentsOFSubjectAfter "W14:W236","Digital Engineering", 60
    CalculateStatusStudentsOFSubjectAfter "AA11:AA236","Operating System", 75
    CalculateStatusStudentsOFSubjectAfter "AE11:AE236","Web Programming II", 90
    CalculateStatusStudentsOFSubjectAfter "AI11:AI236","Database Programming", 105
    CalculateStatusStudentsOFSubjectAfter "AM11:AM236","Data Structure", 120
    CalculateStatusStudentsOFSubjectAfter "AQ11:AQ236","CCNA R&S I", 135
    CalculateStatusStudentsOFSubjectAfter "AU11:AU236","Java Programming I", 150
    CalculateStatusStudentsOFSubjectAfter "AY11:AY236","Capstone Design", 165

    '///////////////////////////////////////////////////////////////////////////
    '////////////////////////// Block All Unused Cell /////////////////////////
    '/////////////////////////////////////////////////////////////////////////
    ' searchStrings = Array("„", "Ã‹ Ã‹", "Ã‹ //","·","‰«ÃÕ") ' Add more substrings as needed
    ' BlockAllCellsNotUsed ws,"H11:H427,L11:L427,P11:P427,T11:T427,X11:X427,AB11:AB427,AF11:AF427,AJ11:AJ427,AN11:AN427,AR11:AR427,AV11:AV427,AZ11:AZ427",searchStrings,"123456"
    
    


End Sub
Sub ManageFirstFoodSheet()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    ' CalculateStatusStudentsOFSubjectBefore "H13:H115","Practical Exercises in food Analysis", 0
    ' CalculateStatusStudentsOFSubjectBefore "L13:L115","Information Tecnology", 15
    ' CalculateStatusStudentsOFSubjectBefore "P13:P115","Research and analysis skills ", 30
    ' CalculateStatusStudentsOFSubjectBefore "T13:T115","Food Saftey", 45
    ' CalculateStatusStudentsOFSubjectBefore "X13:X115","General chemistry", 60
    ' CalculateStatusStudentsOFSubjectBefore "AB13:AB115","Basics of Food preservation", 75
    ' CalculateStatusStudentsOFSubjectBefore "AF13:AF115","Principles of milk and its product", 90
    ' CalculateStatusStudentsOFSubjectBefore "AJ13:AJ115","Human Nutrition", 105
    ' CalculateStatusStudentsOFSubjectBefore "AN13:AN115","English 1", 120
    ' CalculateStatusStudentsOFSubjectBefore "AR13:AR115","Mathmatics ", 135
    ' CalculateStatusStudentsOFSubjectBefore "AV13:AV115"," Analytical Chamistry", 150
    ' CalculateStatusStudentsOFSubjectBefore "AZ13:AZ115","Basics of food technology", 165
    ' CalculateStatusStudentsOFSubjectBefore "BD13:BD115","Meat and its products technology", 165
    ' CalculateStatusStudentsOFSubjectBefore "BH13:BH115","Grain products technology", 165
    ' CalculateStatusStudentsOFSubjectBefore "BL13:BL115","Principles of negotiations ", 165




    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' searchStrings = Array("€‹", "€")
    ' HndelAbsentSubjectHaveFailedOrBassTotalScore_50 ws, "E13:H115,M13:P115,AK13:AN115", searchStrings
    ' HndelAbsentSubjectHaveFailedOrBassTotalScore_100 ws, "I13:L115", searchStrings
    ' HndelAbsentSubjectHaveTotalScore_50_Normal ws, "BI13:BL115", searchStrings
    ' HandelAbsentSubjectHaveTotalScore_150_Normal ws, "Q13:AF115,AO13:AR115,AW13:BH115", searchStrings
    ' HndelAbsentSubjectHaveTotalScore_100_Normal ws, "AS13:AV115", searchStrings


    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' searchStrings = Array("—«”»") ' Add more substrings as needed
    ' HndelSubjectHaveFailedOrBassTotalScore_50 ws, "H13:H115,P13:P115,AN13:AN115", searchStrings
    ' HndelSubjectHaveFailedOrBassTotalScore_100 ws, "L13:L115", searchStrings
    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    ' HandelSubjectHaveTotalScore_50_Normal ws, "BL13:BL115", searchStrings
    ' HandelSubjectHaveTotalScore_150_Normal ws, "T13:T115,X13:X115,AB13:AB115,AF13:AF115,AR13:AR115,AZ13:AZ115,BD13:BD115,BH13:BH115", searchStrings
    ' HandelSubjectHaveTotalScore_100_Normal ws, "AV13:AV115", searchStrings









    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    ' CalculateStatusStudentsOFSubjectAfter "H13:H115","Practical Exercises in food Analysis", 0
    ' CalculateStatusStudentsOFSubjectAfter "L13:L115","Information Tecnology", 15
    ' CalculateStatusStudentsOFSubjectAfter "P13:P115","Research and analysis skills ", 30
    ' CalculateStatusStudentsOFSubjectAfter "T13:T115","Food Saftey", 45
    ' CalculateStatusStudentsOFSubjectAfter "X13:X115","General chemistry", 60
    ' CalculateStatusStudentsOFSubjectAfter "AB13:AB115","Basics of Food preservation", 75
    ' CalculateStatusStudentsOFSubjectAfter "AF13:AF115","Principles of milk and its product", 90
    ' CalculateStatusStudentsOFSubjectAfter "AJ13:AJ115","Human Nutrition", 105
    ' CalculateStatusStudentsOFSubjectAfter "AN13:AN115","English 1", 120
    ' CalculateStatusStudentsOFSubjectAfter "AR13:AR115","Mathmatics ", 135
    ' CalculateStatusStudentsOFSubjectAfter "AV13:AV115"," Analytical Chamistry", 150
    ' CalculateStatusStudentsOFSubjectAfter "AZ13:AZ115","Basics of food technology", 165
    ' CalculateStatusStudentsOFSubjectAfter "BD13:BD115","Meat and its products technology", 165
    ' CalculateStatusStudentsOFSubjectAfter "BH13:BH115","Grain products technology", 165
    ' CalculateStatusStudentsOFSubjectAfter "BL13:BL115","Principles of negotiations ", 165


    '///////////////////////////////////////////////////////////////////////////
    '/////////////////////////// Remove the Circle of all cells ///////////////
    '/////////////////////////////////////////////////////////////////////////
    RemoveShapesFromCells "H13:H118,L13:L118,P13:P118,T13:T118,X13:X118,AB13:AB118,AF13:AF118,AJ13:AJ118,AN13:AN118,AR13:AR118,AV13:AV118,AZ13:AZ118,BD13:BD118,BH13:BH118,BL13:BL118"






End Sub

Sub ManageSecondFoodSheet()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("Sheet1")

    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "G16:G153","English language 2", 0
    CalculateStatusStudentsOFSubjectBefore "K16:K153","Human rights and Eygptian Labor law ", 15
    CalculateStatusStudentsOFSubjectBefore "O16:O153","Profession Ethhics ", 30
    CalculateStatusStudentsOFSubjectBefore "S16:S153","Enterpreneurship ", 45
    CalculateStatusStudentsOFSubjectBefore "W16:W153","Food Addditives", 60
    CalculateStatusStudentsOFSubjectBefore "AA16:AA153","Food Microbiology", 75
    CalculateStatusStudentsOFSubjectBefore "AE16:AE153"," organic chemistry", 90
    CalculateStatusStudentsOFSubjectBefore "AI16:AI153"," Field Training", 105
    CalculateStatusStudentsOFSubjectBefore "AM16:AM153","communication and presentaion", 120
    CalculateStatusStudentsOFSubjectBefore "AQ16:AQ153","English Technical Language", 135
    CalculateStatusStudentsOFSubjectBefore "AU16:AU153","Baiscs of oils and their products", 150
    CalculateStatusStudentsOFSubjectBefore "AY16:AY153","Human feeding", 165

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")
    HndelAbsentSubjectHaveFailedOrBassTotalScore_50 ws, "D16:O153", searchStrings
    HndelAbsentSubjectHaveTotalScore_50_Normal ws, "AJ16:AQ153", searchStrings
    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "T16:AE153,AR16:AY153", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "P16:S153,AF16:AI153", searchStrings


 ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("—«”»") ' Add more substrings as needed
    HndelSubjectHaveFailedOrBassTotalScore_50 ws, "D16:O153", searchStrings
    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    HandelSubjectHaveTotalScore_50_Normal ws, "AJ16:AQ153", searchStrings
    HandelSubjectHaveTotalScore_150_Normal ws, "T16:AE153,AR16:AY153", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "P16:S153,AF16:AI153", searchStrings

    '///////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change //////////////////////////////
    '/////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter "G16:G153","English language 2", 0
    CalculateStatusStudentsOFSubjectAfter "K16:K153","Human rights and Eygptian Labor law ", 15
    CalculateStatusStudentsOFSubjectAfter "O16:O153","Profession Ethhics ", 30
    CalculateStatusStudentsOFSubjectAfter "S16:S153","Enterpreneurship ", 45
    CalculateStatusStudentsOFSubjectAfter "W16:W153","Food Addditives", 60
    CalculateStatusStudentsOFSubjectAfter "AA16:AA153","Food Microbiology", 75
    CalculateStatusStudentsOFSubjectAfter "AE16:AE153"," organic chemistry", 90
    CalculateStatusStudentsOFSubjectAfter "AI16:AI153"," Field Training", 105
    CalculateStatusStudentsOFSubjectAfter "AM16:AM153","communication and presentaion", 120
    CalculateStatusStudentsOFSubjectAfter "AQ16:AQ153","English Technical Language", 135
    CalculateStatusStudentsOFSubjectAfter "AU16:AU153","Baiscs of oils and their products", 150
    CalculateStatusStudentsOFSubjectAfter "AY16:AY153","Human feeding", 165


    
    '///////////////////////////////////////////////////////////////////////////
    '/////////////////////////// Remove the Circle of all cells ///////////////
    '/////////////////////////////////////////////////////////////////////////
    RemoveShapesFromCells "D16:AY153"

End Sub 



Sub ManageFirstRailwaySheet()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("sheet1")



    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "H13:H401","English Language -1", 0
    CalculateStatusStudentsOFSubjectBefore "L13:L401","Information Tecnology", 18
    CalculateStatusStudentsOFSubjectBefore "P13:P401","Research and analysis skills ", 36
    CalculateStatusStudentsOFSubjectBefore "T13:T401","Mathematics", 54
    CalculateStatusStudentsOFSubjectBefore "X13:X401","Engineering Drawing & Projection", 72
    CalculateStatusStudentsOFSubjectBefore "AB13:AB401","Foundation Workshops ", 90
    CalculateStatusStudentsOFSubjectBefore "AF13:AF401","Occupational saftey and health", 108
    CalculateStatusStudentsOFSubjectBefore "AJ13:AJ401","Basics of electical and electronic tecnology", 126
    CalculateStatusStudentsOFSubjectBefore "AN13:AN401","Electronic & control system", 144
    CalculateStatusStudentsOFSubjectBefore "AS13:AS401","social issues ", 162
    CalculateStatusStudentsOFSubjectBefore "AW13:AW401"," Physics", 180
    CalculateStatusStudentsOFSubjectBefore "BA13:BA401","Technical reports ", 198
    CalculateStatusStudentsOFSubjectBefore "BE13:BE401","Main element of railway tracks", 216
    CalculateStatusStudentsOFSubjectBefore "BI13:BI401","Thermenal Machine ", 234
    CalculateStatusStudentsOFSubjectBefore "BM13:BM401","Brake systems ", 252
    CalculateStatusStudentsOFSubjectBefore "BQ13:BQ401","Principle of Negotiation ", 270
    CalculateStatusStudentsOFSubjectBefore "BU13:BU401","Field Training", 288


    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")
    HndelAbsentSubjectHaveFailedOrBassTotalScore_50 ws, "M13:P401,AP13:AS401", searchStrings
    HndelAbsentSubjectHaveFailedOrBassTotalScore_100 ws, "E13:H401,I13:L401", searchStrings
    HndelAbsentSubjectHaveTotalScore_50_Normal ws, "Y13:AB401,BN13:BQ401", searchStrings
    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "Q13:T401,AT13:AW401,BF13:BM401", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "U13:X401,AC13:AN401,AX13:BE401,BR13:BU401", searchStrings


    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("—«”»") ' Add more substrings as needed
    HndelSubjectHaveFailedOrBassTotalScore_50 ws, "M13:P401,AP13:AS401", searchStrings
    HndelSubjectHaveFailedOrBassTotalScore_100 ws, "E13:H401,I13:L401", searchStrings
    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    HandelSubjectHaveTotalScore_50_Normal ws, "Y13:AB401,BN13:BQ401", searchStrings
    HandelSubjectHaveTotalScore_150_Normal ws, "Q13:T401,AT13:AW401,BF13:BM401", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "U13:X401,AC13:AN401,AX13:BE401,BR13:BU401", searchStrings

    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter "H13:H401","English Language -1", 0
    CalculateStatusStudentsOFSubjectAfter "L13:L401","Information Tecnology", 18
    CalculateStatusStudentsOFSubjectAfter "P13:P401","Research and analysis skills ", 36
    CalculateStatusStudentsOFSubjectAfter "T13:T401","Mathematics", 54
    CalculateStatusStudentsOFSubjectAfter "X13:X401","Engineering Drawing & Projection", 72
    CalculateStatusStudentsOFSubjectAfter "AB13:AB401","Foundation Workshops ", 90
    CalculateStatusStudentsOFSubjectAfter "AF13:AF401","Occupational saftey and health", 108
    CalculateStatusStudentsOFSubjectAfter "AJ13:AJ401","Basics of electical and electronic tecnology", 126
    CalculateStatusStudentsOFSubjectAfter "AN13:AN401","Electronic & control system", 144
    CalculateStatusStudentsOFSubjectAfter "AS13:AS401","social issues ", 162
    CalculateStatusStudentsOFSubjectAfter "AW13:AW401"," Physics", 180
    CalculateStatusStudentsOFSubjectAfter "BA13:BA401","Technical reports ", 198
    CalculateStatusStudentsOFSubjectAfter "BE13:BE401","Main element of railway tracks", 216
    CalculateStatusStudentsOFSubjectAfter "BI13:BI401","Thermenal Machine ", 234
    CalculateStatusStudentsOFSubjectAfter "BM13:BM401","Brake systems ", 252
    CalculateStatusStudentsOFSubjectAfter "BQ13:BQ401","Principle of Negotiation ", 270
    CalculateStatusStudentsOFSubjectAfter "BU13:BU401","Field Training", 288
    '///////////////////////////////////////////////////////////////////////////
    '/////////////////////////// Remove the Circle of all cells ///////////////
    '/////////////////////////////////////////////////////////////////////////
    RemoveShapesFromCells "E13:AN401,AS13:BU401"

End Sub
sub ManageSecondRailwaySheet()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("sheet1")



    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "G16:G232","English technical language", 0
    CalculateStatusStudentsOFSubjectBefore "K16:K232","Motive power unit", 18
    CalculateStatusStudentsOFSubjectBefore "O16:O232","Rail joints and welding of Rails ", 36
    CalculateStatusStudentsOFSubjectBefore "S16:S232","Train Maintenance", 54
    CalculateStatusStudentsOFSubjectBefore "W16:W232","Train operations ", 72
    CalculateStatusStudentsOFSubjectBefore "AA16:AA232","Electrical and electromagnetic appliance ", 90
    CalculateStatusStudentsOFSubjectBefore "AE16:AE232","Enterpreneurship", 108
    CalculateStatusStudentsOFSubjectBefore "AI16:AI232","Track maintenance and Drainage", 126
    CalculateStatusStudentsOFSubjectBefore "AM16:AM232","Remove ,install and overhaul engine assemblise", 144
    CalculateStatusStudentsOFSubjectBefore "AQ16:AQ232","Electrical Machine Repair", 162
    CalculateStatusStudentsOFSubjectBefore "AU16:AU232","Profession Ethics", 180
    CalculateStatusStudentsOFSubjectBefore "AY16:AY232","Graduation project", 198
    


    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")
    HndelAbsentSubjectHaveFailedOrBassTotalScore_50 ws, "AR16:AU232", searchStrings
    HndelAbsentSubjectHaveTotalScore_50_Normal ws, "D16:G232", searchStrings
    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "H16:AA232,AF16:AM232", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "AB16:AE232,AN16:AQ232,AV16:AY232", searchStrings


    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("—«”»") ' Add more substrings as needed
    HndelSubjectHaveFailedOrBassTotalScore_50 ws, "AR16:AU232", searchStrings
    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    HandelSubjectHaveTotalScore_50_Normal ws, "D16:G232", searchStrings
    HandelSubjectHaveTotalScore_150_Normal ws, "H16:AA232,AF16:AM232", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "AB16:AE232,AN16:AQ232,AV16:AY232", searchStrings

    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter "G16:G232","English technical language", 0
    CalculateStatusStudentsOFSubjectAfter "K16:K232","Motive power unit", 18
    CalculateStatusStudentsOFSubjectAfter "O16:O232","Rail joints and welding of Rails ", 36
    CalculateStatusStudentsOFSubjectAfter "S16:S232","Train Maintenance", 54
    CalculateStatusStudentsOFSubjectAfter "W16:W232","Train operations ", 72
    CalculateStatusStudentsOFSubjectAfter "AA16:AA232","Electrical and electromagnetic appliance ", 90
    CalculateStatusStudentsOFSubjectAfter "AE16:AE232","Enterpreneurship", 108
    CalculateStatusStudentsOFSubjectAfter "AI16:AI232","Track maintenance and Drainage", 126
    CalculateStatusStudentsOFSubjectAfter "AM16:AM232","Remove ,install and overhaul engine assemblise", 144
    CalculateStatusStudentsOFSubjectAfter "AQ16:AQ232","Electrical Machine Repair", 162
    CalculateStatusStudentsOFSubjectAfter "AU16:AU232","Profession Ethics", 180
    CalculateStatusStudentsOFSubjectAfter "AY16:AY232","Graduation project", 198
    '///////////////////////////////////////////////////////////////////////////
    '/////////////////////////// Remove the Circle of all cells ///////////////
    '/////////////////////////////////////////////////////////////////////////
    RemoveShapesFromCells "D16:AY232"
End Sub




Sub ManageFabricSecond()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("sheet1")



    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "G13:G139","History oftechnology ", 0
    CalculateStatusStudentsOFSubjectBefore "K13:K139","Tech writingand oral pres ", 18
    CalculateStatusStudentsOFSubjectBefore "O13:O139","Spinning and yarn prod. Tech ", 36
    CalculateStatusStudentsOFSubjectBefore "S13:S139","Maintenance management and control", 54
    CalculateStatusStudentsOFSubjectBefore "W13:W139","Fabric structure (2) ", 72
    CalculateStatusStudentsOFSubjectBefore "AA13:AA139","knitting technology (1) ", 90
    CalculateStatusStudentsOFSubjectBefore "AE13:AE139","Industrialchemistry ", 108
    CalculateStatusStudentsOFSubjectBefore "AI13:AI139","Models of applied projects ", 126
    CalculateStatusStudentsOFSubjectBefore "AM13:AM139","profession Ethics  & Egyption labor law", 144
    CalculateStatusStudentsOFSubjectBefore "AQ13:AQ139","English Technical languge", 162
    CalculateStatusStudentsOFSubjectBefore "AU13:AU139","Weaving Machines 2 ", 180
    CalculateStatusStudentsOFSubjectBefore "AY13:AY139","Spinning Machines 1", 198
    CalculateStatusStudentsOFSubjectBefore "BC13:BC139","Reading Machines Catalog", 216
    CalculateStatusStudentsOFSubjectBefore "BG13:BG139","Computer Aided Drawing ", 234
    CalculateStatusStudentsOFSubjectBefore "BK13:BK139","Graduation Project", 252


   ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")
    HndelAbsentSubjectHaveFailedOrBassTotalScore_50 ws, "D13:G139,AJ13:AM139", searchStrings
    HndelAbsentSubjectHaveTotalScore_50_Normal ws, "H13:K139,P13:S139,AF13:AI139,AN13:AQ139", searchStrings
    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "L13:O139,T13:AE139,AR13:BG139", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "BH13:BK139", searchStrings


    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("—«”»") ' Add more substrings as needed
    HndelSubjectHaveFailedOrBassTotalScore_50 ws, "D13:G139,AJ13:AM139", searchStrings
    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    HandelSubjectHaveTotalScore_50_Normal ws, "H13:K139,P13:S139,AF13:AI139,AN13:AQ139", searchStrings
    HandelSubjectHaveTotalScore_150_Normal ws, "L13:O139,T13:AE139,AR13:BG139", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "BH13:BK139", searchStrings


    RemoveShapesFromCells "D13:BK139"
    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test After Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter "G13:G139","History oftechnology ", 0
    CalculateStatusStudentsOFSubjectAfter "K13:K139","Tech writingand oral pres ", 18
    CalculateStatusStudentsOFSubjectAfter "O13:O139","Spinning and yarn prod. Tech ", 36
    CalculateStatusStudentsOFSubjectAfter "S13:S139","Maintenance management and control", 54
    CalculateStatusStudentsOFSubjectAfter "W13:W139","Fabric structure (2) ", 72
    CalculateStatusStudentsOFSubjectAfter "AA13:AA139","knitting technology (1) ", 90
    CalculateStatusStudentsOFSubjectAfter "AE13:AE139","Industrialchemistry ", 108
    CalculateStatusStudentsOFSubjectAfter "AI13:AI139","Models of applied projects ", 126
    CalculateStatusStudentsOFSubjectAfter "AM13:AM139","profession Ethics  & Egyption labor law", 144
    CalculateStatusStudentsOFSubjectAfter "AQ13:AQ139","English Technical languge", 162
    CalculateStatusStudentsOFSubjectAfter "AU13:AU139","Weaving Machines 2 ", 180
    CalculateStatusStudentsOFSubjectAfter "AY13:AY139","Spinning Machines 1", 198
    CalculateStatusStudentsOFSubjectAfter "BC13:BC139","Reading Machines Catalog", 216
    CalculateStatusStudentsOFSubjectAfter "BG13:BG139","Computer Aided Drawing ", 234
    CalculateStatusStudentsOFSubjectAfter "BK13:BK139","Graduation Project", 252
    


End Sub

Sub ManageFabricFirst()
    Dim ws As Worksheet
    Dim searchStrings As Variant
    Dim searchString As Variant
    Set ws = ThisWorkbook.Sheets("sheet1")

    '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectBefore "H15:H110","English language  ", 0
    CalculateStatusStudentsOFSubjectBefore "L15:L110","Information Technology  ", 18
    CalculateStatusStudentsOFSubjectBefore "P15:P110","Mathematics  ", 36
    CalculateStatusStudentsOFSubjectBefore "T15:T110","Physics ", 54
    CalculateStatusStudentsOFSubjectBefore "X15:X110","Foundational Workshops  ", 72
    CalculateStatusStudentsOFSubjectBefore "AB15:AB110","Fabric Structure 1", 90
    CalculateStatusStudentsOFSubjectBefore "AF15:AF110","Weaving Preparation  ", 108
    CalculateStatusStudentsOFSubjectBefore "AJ15:AJ110","int.to Maintenance", 126
    CalculateStatusStudentsOFSubjectBefore "AN15:AN110","English Langauge 2", 144
    CalculateStatusStudentsOFSubjectBefore "AR15:AR110","Principles Of Electonic ", 162
    CalculateStatusStudentsOFSubjectBefore "AV15:AV110","Engineering Drawing", 180
    CalculateStatusStudentsOFSubjectBefore "AZ15:AZ110","Weaving Machines", 198
    CalculateStatusStudentsOFSubjectBefore "BD15:BD110","Prevaintive & corective maintenance", 216
    CalculateStatusStudentsOFSubjectBefore "BH15:BH110"," Textile Materials ", 234
    CalculateStatusStudentsOFSubjectBefore "BL15:BL110","Field Training", 252


   ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Absent Case //////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("€‹", "€")
    HndelAbsentSubjectHaveFailedOrBassTotalScore_50 ws, "E15:L110,AK15:AN110", searchStrings
    HndelAbsentSubjectHaveTotalScore_50_Normal ws, "U15:X110,AG15:AJ110,BA15:BD110", searchStrings
    HandelAbsentSubjectHaveTotalScore_150_Normal ws, "M15:T110,Y15:AF110,AO15:AZ110,BE15:BH110", searchStrings
    HndelAbsentSubjectHaveTotalScore_100_Normal ws, "BI15:BL110", searchStrings


    

    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("—«”»") ' Add more substrings as needed
    HndelSubjectHaveFailedOrBassTotalScore_50 ws, "E15:L110,AK15:AN110", searchStrings
    ' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////// Handel Normal Failed Student Case ///////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    searchStrings = Array("÷", "—·", "Ã‹ ÷") ' Add more substrings as needed
    HandelSubjectHaveTotalScore_50_Normal ws, "U15:X110,AG15:AJ110,BA15:BD110", searchStrings
    HandelSubjectHaveTotalScore_150_Normal ws, "M15:T110,Y15:AF110,AO15:AZ110,BE15:BH110", searchStrings
    HandelSubjectHaveTotalScore_100_Normal ws, "BI15:BL110", searchStrings


    RemoveShapesFromCells "E15:BL110"
     '////////////////////////////////////////////////////////////////////////////////
    '////////////////////// Report Test Before Change //////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////
    CalculateStatusStudentsOFSubjectAfter  "H15:H110","English language  ", 0
    CalculateStatusStudentsOFSubjectAfter  "L15:L110","Information Technology  ", 18
    CalculateStatusStudentsOFSubjectAfter  "P15:P110","Mathematics  ", 36
    CalculateStatusStudentsOFSubjectAfter  "T15:T110","Physics ", 54
    CalculateStatusStudentsOFSubjectAfter  "X15:X110","Foundational Workshops  ", 72
    CalculateStatusStudentsOFSubjectAfter  "AB15:AB110","Fabric Structure 1", 90
    CalculateStatusStudentsOFSubjectAfter  "AF15:AF110","Weaving Preparation  ", 108
    CalculateStatusStudentsOFSubjectAfter  "AJ15:AJ110","int.to Maintenance", 126
    CalculateStatusStudentsOFSubjectAfter  "AN15:AN110","English Langauge 2", 144
    CalculateStatusStudentsOFSubjectAfter  "AR15:AR110","Principles Of Electonic ", 162
    CalculateStatusStudentsOFSubjectAfter  "AV15:AV110","Engineering Drawing", 180
    CalculateStatusStudentsOFSubjectAfter  "AZ15:AZ110","Weaving Machines", 198
    CalculateStatusStudentsOFSubjectAfter  "BD15:BD110","Prevaintive & corective maintenance", 216
    CalculateStatusStudentsOFSubjectAfter  "BH15:BH110"," Textile Materials ", 234
    CalculateStatusStudentsOFSubjectAfter  "BL15:BL110","Field Training", 252
End Sub







' /////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ///////////////////////////////////////////////// Hndel Absent Student /////////////////////////////////////
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////



' the normal student that absent in  subject that have semester work and final exam and success is 60% 
Sub HandelAbsentSubjectHaveTotalScore_150_Normal( ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim counter As Integer
    Dim searchString As Variant
    Dim cell As Range
    Dim Flag As Boolean
    Flag = False
    counter = 0
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        counter = counter + 1
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                ws.Cells(cell.Row, col - 1).ClearContents
                ws.Cells(cell.Row, col + 1).ClearContents
                ws.Cells(cell.Row, col + 2).ClearContents
                ws.Cells(cell.Row, col - 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 2).Interior.Color = RGB(204, 102, 0)
                ' the class work
                Cell3 = ws.Cells(cell.Row, col - 1).Address
                ' the total score
                Cell1 = ws.Cells(cell.Row, col + 1).Address
                ' the final score
                Cell2 = ws.Cells(cell.Row, col).Address

                ' the final grade formula for subject thayt have 150
                ws.Cells(cell.Row, col + 2).Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<30,""—·"",IF(" & Cell1 & "<60,""÷ Ã‹"",IF(" & Cell1 & "<90,""÷"",IF(" & Cell1 & "<97.5,""·"",IF(" & Cell1 & "<112.5,""·"",IF(" & Cell1 & "<127.5,""·"",IF(" & Cell1 & "<=150,""·"",""-""))))))))))"
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col + 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),96)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                Flag =True
            End If
        Next searchString
        counter = counter Mod 4
        If Flag = True And counter = 0 Then
            ws.Cells(cell.Row, col).Value = "€‹"
            ws.Cells(cell.Row, col - 1).Value = 0
            Flag = False
        End If
        
    Next cell
End Sub
Sub HndelAbsentSubjectHaveTotalScore_100_Normal( ws As Worksheet ,range As String , searchStrings As Variant)
    Dim counter As Integer
    Dim searchString As Variant
    Dim cell As Range
    Dim Flag As Boolean
    Flag = False
    counter = 0
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        counter = counter + 1
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                ws.Cells(cell.Row, col - 1).ClearContents
                ws.Cells(cell.Row, col + 1).ClearContents
                ws.Cells(cell.Row, col + 2).ClearContents
                ws.Cells(cell.Row, col - 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 2).Interior.Color = RGB(204, 102, 0)
                ' the class work
                Cell3 = ws.Cells(cell.Row, col - 1).Address
                ' the total score
                Cell1 = ws.Cells(cell.Row, col + 1).Address
                ' the final score
                Cell2 = ws.Cells(cell.Row, col).Address

                ' the final grade formula for subject thayt have 100
                ws.Cells(cell.Row, col + 2).Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<20,""—·"",IF(" & Cell1 & "<40,""÷ Ã‹""," & _
                " IF(" & Cell1 & "<60,""÷"",IF(" & Cell1 & "<65,""·"",IF(" & Cell1 & "<75,""·"",IF(" & Cell1 & "<85,""·"",IF(" & Cell1 & "<=100,""·"",""-""))))))))))"
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col + 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),64)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"

                Flag=True
            End If
        Next searchString
        counter = counter Mod 4
        If Flag = True And counter = 0 Then
            ws.Cells(cell.Row, col).Value = "€‹"
            ws.Cells(cell.Row, col - 1).Value = 0
            Flag = False
        End If
        
    Next cell

End Sub
Sub HndelAbsentSubjectHaveTotalScore_50_Normal( ws As Worksheet ,range As String ,  searchStrings As Variant)
    Dim counter As Integer
    Dim searchString As Variant
    Dim cell As Range
    Dim Flag As Boolean
    Flag = False
    counter = 0
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        counter = counter + 1
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                ws.Cells(cell.Row, col - 1).ClearContents
                ws.Cells(cell.Row, col + 1).ClearContents
                ws.Cells(cell.Row, col + 2).ClearContents
                ws.Cells(cell.Row, col - 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 2).Interior.Color = RGB(204, 102, 0)
                ' the class work
                Cell3 = ws.Cells(cell.Row, col - 1).Address
                ' the total score
                Cell1 = ws.Cells(cell.Row, col + 1).Address
                ' the final score
                Cell2 = ws.Cells(cell.Row, col).Address

                
                ws.Cells(cell.Row, col + 2).Formula ="=IF(" & Cell1 & "=""€‹"",""€‹""," & _
                "IF(" & Cell1 & "=""€‘"",""€‘""," & _
                "IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·""," & _
                "IF(" & Cell2 & "<10,""—·""," & _
                "IF(" & Cell1 & "<20,""÷ Ã‹""," & _
                "IF(" & Cell1 & "<30,""÷""," & _
                "IF(" & Cell1 & "<32.5,""·""," & _
                "IF(" & Cell1 & "<37.5,""·""," & _
                "IF(" & Cell1 & "<42.5,""·""," & _
                "IF(" & Cell1 & "<=50,""·"",""-""))))))))))"
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col + 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),32)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"

                Flag = True
            End If
        Next searchString
        counter = counter Mod 4
        If Flag = True And counter = 0 Then
            ws.Cells(cell.Row, col).Value = "€‹"
            ws.Cells(cell.Row, col - 1).Value = 0
            Flag = False
        End If
        
    Next cell

End Sub


' the normal student that absent in subject the have semester work and final exam and success is 50% [ pass - fail ] only 
Sub HndelAbsentSubjectHaveFailedOrBassTotalScore_150(ws As Worksheet ,range As String  , searchStrings As Variant)

    Dim counter As Integer
    Dim searchString As Variant
    Dim cell As Range
    Dim Flag As Boolean
    
    Flag = False
    counter = 0
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        counter = counter + 1
        ' MsgBox "out side " & counter
        
        
        
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                
                Flag =True
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                ws.Cells(cell.Row, col - 1).ClearContents
                ws.Cells(cell.Row, col + 1).ClearContents
                ws.Cells(cell.Row, col + 2).ClearContents
                ws.Cells(cell.Row, col - 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 2).Interior.Color = RGB(204, 102, 0)
                ' the class work
                Cell3 = ws.Cells(cell.Row, col - 1).Address
                ' the total score
                Cell1 = ws.Cells(cell.Row, col + 1).Address
                ' the final score
                Cell2 = ws.Cells(cell.Row, col).Address

                ws.Cells(cell.Row, col + 2).Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<30,""—«”»"",IF(" & Cell1 & "<75,""—«”»"",IF(" & Cell1 & "<=150,""‰«ÃÕ"",""-""))))))"
                    
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col + 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),96)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                
            End If
        Next searchString
        
        counter =counter Mod 4
        If  counter = 0 And Flag = True  Then
            ws.Cells(cell.Row, col ).Value = "€‹"
            ws.Cells(cell.Row, col - 1).Value = 0
            Flag = False
        End If
        
    Next cell
End Sub
Sub HndelAbsentSubjectHaveFailedOrBassTotalScore_100(ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim counter As Integer
    Dim searchString As Variant
    Dim cell As Range
    Dim Flag As Boolean
    Flag = False
    counter = 0
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        counter = counter + 1
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                Flag =True
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                ws.Cells(cell.Row, col - 1).ClearContents
                ws.Cells(cell.Row, col + 1).ClearContents
                ws.Cells(cell.Row, col + 2).ClearContents
                ws.Cells(cell.Row, col - 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 2).Interior.Color = RGB(204, 102, 0)
                ' the class work
                Cell3 = ws.Cells(cell.Row, col - 1).Address
                ' the total score
                Cell1 = ws.Cells(cell.Row, col + 1).Address
                ' the final score
                Cell2 = ws.Cells(cell.Row, col).Address

                ws.Cells(cell.Row, col + 2).Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<20,""—«”»"",IF(" & Cell1 & "<50,""—«”»"",IF(" & Cell1 & "<=100,""‰«ÃÕ"",""-""))))))"
                    
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col + 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),64)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                Flag =True
            End If
        Next searchString
        counter = counter Mod 4
        If Flag = True And counter = 0 Then
            ws.Cells(cell.Row, col).Value = "€‹"
            ws.Cells(cell.Row, col - 1).Value = 0
            Flag = False
        End If
        
    Next cell
End Sub
Sub HndelAbsentSubjectHaveFailedOrBassTotalScore_50(ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim counter As Integer
    Dim searchString As Variant
    Dim cell As Range
    Dim Flag As Boolean
    Flag = False
    counter = 0
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        counter = counter + 1
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                Flag =True
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                ws.Cells(cell.Row, col - 1).ClearContents
                ws.Cells(cell.Row, col + 1).ClearContents
                ws.Cells(cell.Row, col + 2).ClearContents
                ws.Cells(cell.Row, col - 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 1).Interior.Color = RGB(204, 102, 0)
                ws.Cells(cell.Row, col + 2).Interior.Color = RGB(204, 102, 0)
                ' the class work
                Cell3 = ws.Cells(cell.Row, col - 1).Address
                ' the total score
                Cell1 = ws.Cells(cell.Row, col + 1).Address
                ' the final score
                Cell2 = ws.Cells(cell.Row, col).Address

                ws.Cells(cell.Row, col + 2).Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<10,""—«”»"",IF(" & Cell1 & "<25,""—«”»"",IF(" & Cell1 & "<=50,""‰«ÃÕ"",""-""))))))"
                    
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col + 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),32)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                Flag =True
            End If
        Next searchString
        counter = counter Mod 4
        If Flag = True And counter = 0 Then
            ws.Cells(cell.Row, col).Value = "€‹"
            ws.Cells(cell.Row, col - 1).Value = 0
            Flag = False
        End If
        
    Next cell
End Sub





' //////////////////////////////////////////////////////////////////////////////////////////////////////
' /////////////////////////////////////////// Hndel Failed Student ////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////////////////////////////

' the normal student that fail in subject that have semester work and final exam and success is 60%
Sub HandelSubjectHaveTotalScore_150_Normal(ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim cell As Range
    Dim searchString As Variant
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                For i = 1 To 3
                    ws.Cells(cell.Row, col - i).ClearContents
                    ws.Cells(cell.Row, col - i).Interior.Color = RGB(204, 102, 0)
                Next i
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
                Cell1 = ws.Cells(cell.Row, col - 1).Address
                Cell2 = ws.Cells(cell.Row, col - 2).Address
                Cell3 = ws.Cells(cell.Row, col - 3).Address
                cell.Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<30,""—·"",IF(" & Cell1 & "<60,""÷ Ã‹"",IF(" & Cell1 & "<90,""÷"",IF(" & Cell1 & "<97.5,""·"",IF(" & Cell1 & "<112.5,""·"",IF(" & Cell1 & "<127.5,""·"",IF(" & Cell1 & "<=150,""·"",""-""))))))))))"
                
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col - 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),97)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
           
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
            End If
        Next searchString
    Next cell
End SUB
Sub HandelSubjectHaveTotalScore_100_Normal(ws As Worksheet ,range As String  , searchStrings As Variant)
     Dim cell As Range
     Dim searchString As Variant
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                For i = 1 To 3
                    ws.Cells(cell.Row, col - i).ClearContents
                    ws.Cells(cell.Row, col - i).Interior.Color = RGB(204, 102, 0)
                Next i
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
                Cell1 = ws.Cells(cell.Row, col - 1).Address
                Cell2 = ws.Cells(cell.Row, col - 2).Address
                Cell3 = ws.Cells(cell.Row, col - 3).Address
               ' the final grade formula for subject thayt have 100
                cell.Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<20,""—·"",IF(" & Cell1 & "<40,""÷ Ã‹""," & _
                " IF(" & Cell1 & "<60,""÷"",IF(" & Cell1 & "<65,""·"",IF(" & Cell1 & "<75,""·"",IF(" & Cell1 & "<85,""·"",IF(" & Cell1 & "<=100,""·"",""-""))))))))))"
                
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col - 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),64)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
           
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
            End If
        Next searchString
    Next cell
End SUB
Sub HandelSubjectHaveTotalScore_50_Normal(ws As Worksheet ,range As String  , searchStrings As Variant)
     Dim cell As Range
     Dim searchString As Variant
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                For i = 1 To 3
                    ws.Cells(cell.Row, col - i).ClearContents
                    ws.Cells(cell.Row, col - i).Interior.Color = RGB(204, 102, 0)
                Next i
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
                Cell1 = ws.Cells(cell.Row, col - 1).Address
                Cell2 = ws.Cells(cell.Row, col - 2).Address
                Cell3 = ws.Cells(cell.Row, col - 3).Address
                cell.Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<10,""—·"",IF(" & Cell1 & "<20,""÷ Ã‹"",IF(" & Cell1 & "<30,""÷"",IF(" & Cell1 & "<32.5,""·"",IF(" & Cell1 & "<37.5,""·"",IF(" & Cell1 & "<42.5,""·"",IF(" & Cell1 & "<=50,""·"",""-""))))))))))"
                
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col - 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),32)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
           
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
            End If
        Next searchString
    Next cell
End SUB



' the normal student that fail in the subject that have sementer work and final exam and success is 50% [ fail - pass ] 
Sub HndelSubjectHaveFailedOrBassTotalScore_150(ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim cell As Range
    Dim searchString As Variant
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                For i = 1 To 3
                    ws.Cells(cell.Row, col - i).ClearContents
                    ws.Cells(cell.Row, col - i).Interior.Color = RGB(204, 102, 0)
                Next i
                
                Cell1 = ws.Cells(cell.Row, col - 1).Address
                Cell2 = ws.Cells(cell.Row, col - 2).Address
                Cell3 = ws.Cells(cell.Row, col - 3).Address
                
                
                
                ' the final grade formula for subject thayt have 150
                cell.Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<30,""—«”»"",IF(" & Cell1 & "<75,""—«”»"",IF(" & Cell1 & "<=150,""‰«ÃÕ"",""-""))))))"
                
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col - 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),97)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
            End If
        Next searchString
    Next cell
End Sub
Sub HndelSubjectHaveFailedOrBassTotalScore_100(ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim cell As Range
    Dim searchString As Variant
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                For i = 1 To 3
                    ws.Cells(cell.Row, col - i).ClearContents
                    ws.Cells(cell.Row, col - i).Interior.Color = RGB(204, 102, 0)
                Next i
                
                Cell1 = ws.Cells(cell.Row, col - 1).Address
                Cell2 = ws.Cells(cell.Row, col - 2).Address
                Cell3 = ws.Cells(cell.Row, col - 3).Address
                
                
                
                ' the final grade formula for subject thayt have 150
                cell.Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<20,""—«”»"",IF(" & Cell1 & "<50,""—«”»"",IF(" & Cell1 & "<=100,""‰«ÃÕ"",""-""))))))"
                
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col - 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),64)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
            End If
        Next searchString
    Next cell
End Sub
Sub HndelSubjectHaveFailedOrBassTotalScore_50(ws As Worksheet ,range As String  , searchStrings As Variant)
    Dim cell As Range
    Dim searchString As Variant
    For Each cell In ws.Range(range)
        ' Reset found flag
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(204, 102, 0)
                cell.ClearContents
                col = cell.Column
                For i = 1 To 3
                    ws.Cells(cell.Row, col - i).ClearContents
                    ws.Cells(cell.Row, col - i).Interior.Color = RGB(204, 102, 0)
                Next i
                
                Cell1 = ws.Cells(cell.Row, col - 1).Address
                Cell2 = ws.Cells(cell.Row, col - 2).Address
                Cell3 = ws.Cells(cell.Row, col - 3).Address
                
                
                
                ' the final grade formula for subject thayt have 150
                cell.Formula = "=IF(" & Cell1 & "=""€‹"",""€‹"",IF(" & Cell1 & "=""€‘"",""€‘"",IF(" & Cell1 & "=""⁄–—"",""„ƒÃ·"",IF(" & Cell2 & "<10,""—«”»"",IF(" & Cell1 & "<25,""—«”»"",IF(" & Cell1 & "<=50,""‰«ÃÕ"",""-""))))))"
                
                ' the total score grade formula  maximum 64%
                ws.Cells(cell.Row, col - 1).Formula = "=IF(ISNUMBER(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))" & _
                ",MIN(IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))),32)," & _
                "IF(" & Cell2 & "=""€"",""€‹"",IF(" & Cell2 & "="" "",""€‹"",IF(" & Cell2 & "=""⁄–—"",""⁄–—"",IF(" & Cell2 & "=""€‘"",""€‘""," & Cell3 & "+" & Cell2 & ")))))"
                
                ws.Cells(cell.Row, col - 2).Value = "€‹"
                ws.Cells(cell.Row, col - 3).Value = 0
            End If
        Next searchString
    Next cell
End Sub






' range is like D14:AY324







Sub RemoveShapesFromCells(rng As String)
    Dim ws As Worksheet
    Dim cell As Range
    Dim shape As shape
    Dim cellRange As Range
    Dim shapeToDelete As New Collection
    Dim shpName As Variant
    
    ' Set the worksheet (change "Sheet1" to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Define the range of cells from which to remove shapes (change "A1:A10" to your range)
    Set cellRange = ws.Range(rng)
    
    ' Loop through all shapes in the worksheet
    For Each shape In ws.Shapes
        ' Check if the shape intersects with any cell in the specified range
        For Each cell In cellRange
            If Not Intersect(cell, ws.Range(shape.TopLeftCell.Address & ":" & shape.BottomRightCell.Address)) Is Nothing Then
                shapeToDelete.Add shape.Name
                Exit For
            End If
        Next cell
    Next shape
    
    ' Delete the shapes from the collection
    For Each shpName In shapeToDelete
        ws.Shapes(shpName).Delete
    Next shpName
End Sub





Sub CalculateStatusStudentsOFSubjectBefore(range As String , subject As String ,location as Integer)
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    
    Dim cell As Range
    Dim countA As Long ' „ 
    Dim countB As Long ' Ã‹ Ã‹
    Dim countC As Long ' Ã‹ //
    Dim countD As Long ' ·
    Dim countF As Long ' ÷ 
    Dim countFF As Long ' ÷ Ã‹
    Dim countEF As Long ' —·
    Dim countAbsent As Long ' €‹
    Dim CountSuccess As Long ' ‰«ÃÕ
    Dim CountFailed As Long ' —«”»
    Dim SubjectName As Range

    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("Sheet2")


    countA =0 ' „ 
    countB =0' Ã‹ Ã‹
    countC =0 ' Ã‹ //
    countD =0 ' ·
    countF =0 ' ÷ 
    countFF = 0 ' ÷ Ã‹
    countEF = 0 ' —·
    countAbsent =0 ' €‹
    CountSuccess =0 ' ‰«ÃÕ
    CountFailed =0 ' —«”»
    
    For Each cell In wsSource.Range(range)
        Select Case cell.Value
            Case "„"
                countA = countA + 1
            Case "Ã‹ Ã‹"
                countB = countB + 1
            Case "//Ã‹"
                countC = countC + 1
            Case "·"
                countD = countD + 1
            Case "÷"
                countF = countF + 1
            Case "÷ Ã‹"
                countFF = countFF + 1
            Case "—·"
                countEF = countEF + 1
            Case "€‹"
                countAbsent = countAbsent + 1
            Case "‰«ÃÕ"
                CountSuccess = CountSuccess + 1
            Case "—«”»"
                CountFailed = CountFailed + 1
        End Select
    Next cell

    ' output result to sheet2 
    set SubjectName = wsDest.Range(wsDest.Cells(1+location,1),wsDest.Cells(1+location,4))
    SubjectName.Merge
    SubjectName.Value = subject
    SubjectName.Interior.Color = RGB(204, 102, 0)
    ' header of report 
    wsDest.Cells(2+location,1).Value = "Grade"
    wsDest.Cells(3+location,1).Value = "„"
    wsDest.Cells(4+location,1).Value = "Ã‹ Ã‹"
    wsDest.Cells(5+location,1).Value = "//Ã‹ "
    wsDest.Cells(6+location,1).Value = "·"
    wsDest.Cells(7+location,1).Value = "÷"
    wsDest.Cells(8+location,1).Value = "Ã‹ ÷"
    wsDest.Cells(9+location,1).Value = "—·"
    wsDest.Cells(10+location,1).Value = "€‹"
    wsDest.Cells(11+location,1).Value = "‰«ÃÕ"
    wsDest.Cells(12+location,1).Value = "—«”»"

    wsDest.Cells(2+location,2).Value = "Grade Before"
    wsDest.Cells(3+location,2).Value = countA
    wsDest.Cells(4+location,2).Value = countB
    wsDest.Cells(5+location,2).Value = countC
    wsDest.Cells(6+location,2).Value = countD
    wsDest.Cells(7+location,2).Value = countF
    wsDest.Cells(8+location,2).Value = countFF
    wsDest.Cells(9+location,2).Value = countEF
    wsDest.Cells(10+location,2).Value = countAbsent
    wsDest.Cells(11+location,2).Value = CountSuccess
    wsDest.Cells(12+location,2).Value = CountFailed
    



End Sub




Sub CalculateStatusStudentsOFSubjectAfter(range As String , subject As String , location As Integer)
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    
    Dim cell As Range
    Dim countA As Long ' „ 
    Dim countB As Long ' Ã‹ Ã‹
    Dim countC As Long ' Ã‹ //
    Dim countD As Long ' ·
    Dim countF As Long ' ÷ 
    Dim countFF As Long ' ÷ Ã‹
    Dim countEF As Long ' —·
    Dim countAbsent As Long ' €‹
    Dim CountSuccess As Long ' ‰«ÃÕ
    Dim CountFailed As Long ' —«”»
    Dim SumTotalStudentsFailed As Long 
    Dim ResultOfReport As Range 
    Dim ReportValue As Range
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("Sheet2")
    
    SumTotalStudentsFailed=0
    countA =0 ' „ 
    countB =0' Ã‹ Ã‹
    countC =0 ' Ã‹ //
    countD =0 ' ·
    countF =0 ' ÷ 
    countFF = 0 ' ÷ Ã‹
    countEF = 0 ' —·
    countAbsent =0 ' €‹
    CountSuccess =0 ' ‰«ÃÕ
    CountFailed =0 ' —«”»
    
    For Each cell In wsSource.Range(range)
        Select Case cell.Value
            Case "„"
                countA = countA + 1
            Case "Ã‹ Ã‹"
                countB = countB + 1
            Case "//Ã‹"
                countC = countC + 1
            Case "·"
                countD = countD + 1
            Case "÷"
                countF = countF + 1
            Case "÷ Ã‹"
                countFF = countFF + 1
            Case "—·"
                countEF = countEF + 1
            Case "€‹"
                countAbsent = countAbsent + 1
            Case "‰«ÃÕ"
                CountSuccess = CountSuccess + 1
            Case "—«”»"
                CountFailed = CountFailed + 1
        End Select
    Next cell


    wsDest.Cells(2+location,3).Value = "Grade After"
    wsDest.Cells(3+location,3).Value = countA
    wsDest.Cells(4+location,3).Value = countB
    wsDest.Cells(5+location,3).Value = countC
    wsDest.Cells(6+location,3).Value = countD
    wsDest.Cells(7+location,3).Value = countF
    wsDest.Cells(8+location,3).Value = countFF
    wsDest.Cells(9+location,3).Value = countEF
    wsDest.Cells(10+location,3).Value = countAbsent
    wsDest.Cells(11+location,3).Value = CountSuccess
    wsDest.Cells(12+location,3).Value = CountFailed

    wsDest.Cells(2+location,4).Value = "Status"
    wsDest.Cells(3+location,4).Formula= "=IF(" & wsDest.Cells(3+location,3).Address & "=" & wsDest.Cells(3+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(4+location,4).Formula= "=IF(" & wsDest.Cells(4+location,3).Address & "=" & wsDest.Cells(4+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(5+location,4).Formula= "=IF(" & wsDest.Cells(5+location,3).Address & "=" & wsDest.Cells(5+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(6+location,4).Formula= "=IF(" & wsDest.Cells(6+location,3).Address & "=" & wsDest.Cells(6+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(7+location,4).Formula= "=IF(" & wsDest.Cells(7+location,3).Address & "=" & wsDest.Cells(7+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(8+location,4).Formula= "=IF(" & wsDest.Cells(8+location,3).Address & "=" & wsDest.Cells(8+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(9+location,4).Formula= "=IF(" & wsDest.Cells(9+location,3).Address & "=" & wsDest.Cells(9+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(10+location,4).Formula= "=IF(" & wsDest.Cells(10+location,3).Address & "=" & wsDest.Cells(10+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(11+location,4).Formula= "=IF(" & wsDest.Cells(11+location,3).Address & "=" & wsDest.Cells(11+location,2).Address & ",""True"",""False"")"
    wsDest.Cells(12+location,4).Formula= "=IF(" & wsDest.Cells(12+location,3).Address & "=" & wsDest.Cells(12+location,2).Address & ",""True"",""False"")"

    
    set ResultOfReport= wsDest.Range(wsDest.Cells(13+location,1),wsDest.Cells(14+location,2))
    ResultOfReport.Merge
    ResultOfReport.Value=" Report Result "
    ResultOfReport.Interior.Color = RGB(204, 102, 0)

    SumTotalStudentsFailed =countAbsent+countEF+countFF+countF+CountFailed
    set ReportValue = wsDest.Range(wsDest.Cells(13+location,3),wsDest.Cells(14+location,4))
    ReportValue.Merge
    If SumTotalStudentsFailed = wsDest.Cells(10+location,3).Value Then 
        ReportValue.Value = "True"
        ReportValue.Interior.Color =RGB(102,255 ,102)

    Else 
        ReportValue.Value = "False"
        ReportValue.Interior.Color =RGB(255, 51, 51)
    End IF 


    For Each cell In wsDest.Range(wsDest.Cells(3+location,4),wsDest.Cells(12+location,4))
        If cell.Value = "True" Then
            cell.Interior.Color = RGB(102,255 ,102)
        Else 
            cell.Interior.Color = RGB(255, 51, 51)
        End If
    Next cell
    ' ceter the result data 
    With wsDest.Range(wsDest.Cells(1+location,1),wsDest.Cells(12+location,4))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub


Sub BlockAllCellsNotUsed(ws As Worksheet ,range As String  , searchStrings As Variant, password As String)
    Dim searchString As Variant
    Dim cell As Range
    For Each cell In ws.Range(range)
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                col = cell.Column
                cell.Locked = True
                ws.Cells(cell.Row, col - 1).Locked = True
                ws.Cells(cell.Row, col - 2).Locked = True
                ws.Cells(cell.Row, col - 3).Locked = True
            Else 
                col = cell.Column
                cell.Locked = False
                ws.Cells(cell.Row, col - 1).Locked = False
                ws.Cells(cell.Row, col - 2).Locked = False
                ws.Cells(cell.Row, col - 3).Locked = False
            End If
        Next searchString
    Next cell

    ws.Protect Password:=password , AllowFiltering:=True

            

        
End Sub

Sub UnblockAllCellsNotUsed(ws As Worksheet ,range As String  , searchStrings As Variant,password As String)
    Dim searchString As Variant
    ws.Unprotect Password:=password
    ' Dim cell As Range
    ' For Each cell In ws.Range(range)
    '     For Each searchString In searchStrings
    '         If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
    '             col = cell.Column
    '             cell.Locked = False
    '             ws.Cells(cell.Row, col - 1).Locked = False
    '             ws.Cells(cell.Row, col - 2).Locked = False
    '             ws.Cells(cell.Row, col - 3).Locked = False
    '         End If
    '     Next searchString
    ' Next cell
End Sub


Sub  InsertCircleBasedOnSubstring()
    Dim ws As Worksheet
    Dim cell As Range
    Dim shape As shape
    Dim cellWidth As Single
    Dim cellHeight As Single
    Dim diameter As Single
    Dim searchStrings As Variant
    Dim searchString As Variant
    Dim found As Boolean
    Dim shapeTop As Single
    
    ' Set the worksheet (change "Sheet1" to your sheet's name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Define the Arabic substrings to search for (as an array)
    searchStrings = Array("—·", "÷", "Ã‹ ÷", "€‹","—«”»") ' Add more substrings as needed
    
    ' Loop through each cell in the specified range (change "A1:A10" to your range)
    For Each cell In ws.Range("H13:H118,L13:L118,P13:P118,T13:T118,X13:X118,AB13:AB118,AF13:AF118,AJ13:AJ118,AN13:AN118,AR13:AR118,AV13:AV118,AZ13:AZ118,BD13:BD118,BH13:BH118,BL13:BL118")
        ' Reset found flag
        
        
        ' Check if any of the substrings are found in the cell
        For Each searchString In searchStrings
            If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                ' Get the cell dimensions
                cellWidth = cell.Width
                cellHeight = cell.Height
                
                ' Determine the diameter of the circle to fit within the cell
                diameter = Application.Min(cellWidth, cellHeight)
                
              
                
                ' Add a circle shape to the worksheet
                Set shape = ws.Shapes.AddShape(msoShapeOval, cell.Left, cell.Top, cellWidth, cellHeight)
                
                ' Format the circle
                shape.Fill.Transparency = 1 ' No fill color
                shape.Line.ForeColor.RGB = RGB(0, 0, 0) ' Black border
                shape.Line.Weight = 2 ' Border weight (optional)
                
                ' Set found flag to true
                found = True
            End If
        Next searchString
        ' If any substring was found, exit the loop for this cell
    Next cell
End Sub
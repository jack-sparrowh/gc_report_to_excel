# gc_report_to_excel
Code that takes GC_reports from GC-MS where i work and combines files with the same name, aggregates them and saves as the sheet in excel. Below is an example of a pdf file that is being converted while running this code.

"
Data File file_path
Sample Name: 265.17 (2)

=====================================================================
Acq. Operator   : ###                             Seq. Line :   8
Acq. Instrument : Instrument 1                    Location : Vial 107
Injection Date  : arbitraty_date                  Inj :   1
                                                Inj Volume : 0.5 µl
Acq. Method     : method_path
Last changed    : arbitraty_date
                  (modified after loading)
Analysis Method : method_path
Last changed    : arbitraty_date
                  (modified after loading)
Additional Info : Peak(s) manually integrated
=====================================================================
                         ISTD Percent Report                        
=====================================================================

Sorted By             :      Signal
Calib. Data Modified  :      arbitrary_date AM
Multiplier:                   :      1.0000
Dilution:                     :      1.0000
Sample Amount:               :   2038.00000  [g]
Do not use Multiplier & Dilution Factor with ISTDs
Sample ISTD Information:
ISTD  ISTD Amount   Name
  #      [g]    
----|-------------|-------------------------
  1    490.00000   G1E Me

Signal 1: FID1 A, Front Signal
Uncalibrated peaks RF :      1.00000
ISTD for Uncal. peaks :      G1E Me

RetTime  Type  ISTD    Area     Amt/Area    Amount   Grp   Name
 [min]         used  [pA*s]      ratio        %    
-------|------|----|----------|----------|----------|--|------------------
  8.576 MM   I    1  425.32822    1.00000  24.043180    G1E Me                                            
 10.404 MM        1   33.64039    1.00000   1.901642    ?                                                
 12.463 MM        1  616.31982 8.08000e-1  28.150441    kwas mlekowy                                      
 15.681 MM R      1  498.48685 4.98000e-1  14.033007    Gliceryna                                        
 15.853 MM T      1    2.48454    1.00000   0.140448    ?                                                
 17.379 MM        1   92.09738 8.08000e-1   4.206553    kwas mlekowy 2                                    
 19.856 MM        1   68.03019    1.00000   3.845647    mleczan glicerolu                                
 20.135 MM        1  523.92346    1.00000  29.616624    mleczan glicerolu                                
 21.270 MM        1   15.76477 8.08000e-1   0.720057    kwas mlekowy 3                                    
 23.126 MM        1    4.70342    1.00000   0.265878    ?                                                
 23.438 MM        1   75.11490    1.00000   4.246135    dimleczan glicerolu                              
 23.603 MM        1   92.02913    1.00000   5.202272    dimleczan glicerolu                              
 24.619 MM        1    1.79380 8.08000e-1   0.081932    kwas mlekowy 4                                    
 26.753 MM        1   26.48589    1.00000   1.497209    trimleczan glicerolu                              
 31.097 MM        1    3.88209    1.00000   0.219449    ?                                                

Instrument 1 arbitrary_date AM

Page  1 of 2

Data File file_path
Sample Name: 265.17 (2)

Totals without ISTD(s) :              94.127292

1 Warnings or Errors :

Warning : Calibration warnings (see calibration table listing)

=====================================================================
                          *** End of Report ***

Instrument 1 arbitrary_date AM

Page  2 of 2
"

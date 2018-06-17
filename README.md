# Vincenty Direct and Inverse Solution of Geodesics on the Ellipsoid
algorithms by Thaddeus Vincenty (1975)
---
Based on implementation in Java Script by © Chris Veness 2002-2017 MIT License  
https://www.movable-type.co.uk/scripts/latlong-vincenty.html  
https://github.com/chrisveness/geodesy

To make the long story short, I as looking for a way to calculate coordinates, distance and azimuth in Excel.
I checked several available solutions but they were either incomplete, did not work or results were inacurate.
That is how I ended up developing my own, complete Vincenty Direct and Inverse formulae implementation.

Solution contains 6 functions implementing **Vincenty Direct** and **Vincenty Inverse** calculations as well as 2 functions for Decimal ↔ Degrees/Minutes/Seconds format conversion and uses WGS84 model. Functions are available in Excel files:
+ [Vincenty.xlsm](../../raw/master/Vincenty.xlsm) - Excel Macro-Enabled Workbook
+ [Vincenty.xlam](../../raw/master/Vincenty.xlam) - Excell Add-in
+ [Vincenty.xls](../../raw/master/Vincenty.xls) - Excel 97-2003 Add-in
+ [Vincenty.xla](../../raw/master/Vincenty.xla) - Excel 97-2003 Workbook

For better change tracking source code has been placed separately in [Vincenty.bas](Vincenty.bas), [InvParams.cls](InvParams.cls), [DirParams.cls](DirParams.cls) files.

Calculation results have been validated using 1200 test cases generated for 6 range clusters and distance between 10 m and 30,000 km 
against **Geoscience Australia** website:
+ http://www.ga.gov.au/geodesy/datums/vincenty_direct.jsp
+ http://www.ga.gov.au/geodesy/datums/vincenty_inverse.jsp  
and **GeodSolve library** by Charles Karney:
+ https://geographiclib.sourceforge.io/cgi-bin/GeodSolve
+ https://geographiclib.sourceforge.io/scripts/geod-google.html
+ https://link.springer.com/article/10.1007%2Fs00190-012-0578-z  

A|Geoscience Australia|GeodSolve Library|
---|---:|---:
VincentyDirLat       |0.0000005%          |0.0000000%
VincentyDirLon       |0.0000002%          |0.0000001%
VincentyDirRevAzimuth|0.0000833%          |0.0000000%
VincentyInvDistance  |0.0024183%          |0.0001269%
VincentyInvFwdAzimuth|0.0008098%          |0.0003928%
VincentyInvRevAzimuth|0.0005245%          |0.0003928%

For complete test results refer to [VincentTest.xlsm](../../raw/master/VincentyTest.xlsm) file.

**References:**
+ [Wikipedia, Vincenty's formulae](https://en.wikipedia.org/wiki/Vincenty%27s_formulae)
+ https://www.ngs.noaa.gov/PUBS_LIB/inverse.pdf
+ https://geographiclib.sourceforge.io/geodesic-papers/vincenty75b.pdf
+ [Wikipedia, Geodesics on an ellipsoid](https://en.wikipedia.org/wiki/Geodesics_on_an_ellipsoid)
+ [Wikipedia, Great-circle distance](https://en.wikipedia.org/wiki/Great-circle_distance)
+ [Wikipedia, Haversine formula](https://en.wikipedia.org/wiki/Haversine_formula)

## ExcelVBA
### SAS Add-in for Microsoft Office, install package, 2019-11-06
64 bit Download in [here](http://ftp.sas.com/techsup/download/hotfix/HF2/BIRD/msofficeint__94220__wx6__en__sp0__1/)    
32 bit Download in [here](http://ftp.sas.com/techsup/download/hotfix/HF2/BIRD/msofficeint__94220__win__en__sp0__1/)  
  
*ps: version should be match with your office's version, not SAS. *  

### SAS.OfficeAdd-in.txt  
By batch import multiple *sas7bdat* or *xpt* into Excel workbook through file system location.  

### SAS.OfficeAdd-in 2.txt  
By batch import multiple *sas7bdat* into Excel workbook through SAS library, **which need run the SAS code in SAS program tab manually** *at first*. This method can also solve the external SAS format catalog loading issue.

### colADJ.txt
Clear the unexpect formats cross all worksheets inside SDTM specifcation.  

### About set VBA project reference by program  
User need select SAS Add-in in VBA reference before running VBA macro. So first step we need do that in the beginning of the program. I just find that after SAS Add-in version changed, the GUID also changed for its VBA reference.   
So when using GUID, in version 8.3, it changed to {9E9CE404-E32F-4DEC-BC01-292916642B95} .   
See more detail info in [here](https://stackoverflow.com/questions/9879825/how-to-add-a-reference-programmatically-using-vba#:~:text=There%20are%20two%20ways%20to%20add%20references%20using%20VBA.%20.,to%20add%20a%20reference%20to.)   
About how to find GUID, in previous I used this [method](https://www.thespreadsheetguru.com/vba/2014/3/16/display-object-library-reference-guid-information), but somehow it does not work now in my Excel.
So I have to search in my regedit.exe , with key words "SAS.OfficeAddin.tlb", and then a path with the GUID format will be found,  as for major and minor value, they are still same.    


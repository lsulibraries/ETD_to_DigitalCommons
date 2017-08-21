
================================
history/context of spreadsheets:
================================



1) do we expect the primary source to have changed since data pulled? 


     `data may be requested changed after ingest into Digital Commons -- the graduate school will be responsible for this`


2) Is there a way to catch in the future any item that was submitted but is not yet a candidate for migration?


3) How to bring in new dissertations from Fall graduates, etc future graduates? 


     `second pull from database planned for after migration, then isolating changes & ETL'ing new items`


4) Do lingering candidates for graduation get switched automatically or will they be merged into digital commons manually?

5) what is the loading pathway for Digital Commons?  and what form will the data need be, in order to be loaded?


6) Which category gets published in Digital Commons:  Available, Submitted, Withheld? 

     `all three`


7) is any curating necessary, other than merging with 2nd source? 

     `not unless Rene or Justin say so`

8) are any fields in the exported data superfluous?

     `migrate all`


Data Transformation:
*********************

1)  what is other data source into which we are merging?

     `catalog data`

2) is there any other data to pull, other than the metadata/pdf (or other binary)?

     `Requirement is that we do fields renaming very conservatively.`
     `Degree name and Department Names have changed -- we must remap the old names to new names.  Or we may keep the names unchanged and link to multiple names in same search.`

3) Plans
     `ETL a hundred or so items, then consult with Justin if things look right.`


Outliers in ETL database records
*********************************

1) ODD FILETYPES: {'wmv', 'Megan_McVay_cataloguing_abstract', 'ThesisTaraBSmithsonpdf', 'docx', 'doc', 'gz', 'aif', 'DissertationWaychoff', 'PDF', 'pdf', 'Jenkins_thesis', 'swf', 'tif'}  

     It's ok for a main document to have supporting documents of an odd filetype.  It's not ok to have a file missing the period separating the extension from the filename -- or a file missing an extension altogether.

2) MORE THAN ONE BINARY:  

     It's ok to have one main document & many supporting documents.  It's not ok to have a main document split by chapter into many documents:  (E.g., etd-0409103-184148, etd-09012004-114224, etd-09012004-114224, etd-04152004-142117, etd-0830102-145811, etd-0707103-142120, etd-1112103-221719, etd-0327102-091522, etd-0710102-054039)

3) SAME BINARY UPLOADED 4 TIMES:  

     we have only one actual binary.  Must be careful to not upload it to Digital Commons 4 times.  etd-08272008-000329,  etd-06092010-230857, etd-06092010-221901

4) SAME ADVISOR UPLOADED 10 times:  

     Same situation as above.

5) 9 ADVISORS:  etd-01042012-135713

6) 178 KEYWORDS:  etd-0415102-110441    

     Six keywords seems to be the most common limit set by other DigitalCommons installations - based on 30 minutes of googling.

7) etd-06092008-192351  has the main record availability as withheld, but the file is listed as available to all.  Contradiction.

8) etd-06062010-192030 has a record available item, but no file.

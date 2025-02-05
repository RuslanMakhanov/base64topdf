# base64topdf

## Scripts that I wrote to convert base64 codes to pdf files.

*The first script reads all lines from big csv file than takes id and base64 code from each line.*
*After it opens excel file to find the name of pdf file using id it copied from csv file and than creates a file with base64 and name taken from excel file.*

*The second script does the same except for it reads second base64 code in line (If it exists) and than search for file name using the same logic but tracking it in another table in excel document*

### Prehistory of why I created those scripts

While working as a Software Administrator coding on python doesn't include my responsibilities We had a problem in my department. One of the administrators decided to delete all files that one of the services of my org has given to people. 
so we lost all pdf files of the services given to people from 1st of september to 31th of December. 
The head of department came to me asking for help as I used to work as Django Developer before and asked if I could write the simple script.

I had 4 big csv files for each month the data were deleted from. Those files contained all the data I needed including id of base64 files which was on every line of code. So I created basic logic and described how the first code works before. 

***After everything was done a person who was responsible for deleting all files bought me 2 pizzas cause if i did not do this script the head of deparment would force him to copy base64 cod manually***



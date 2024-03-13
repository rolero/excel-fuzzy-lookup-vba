# excel-fuzzy-lookup-vba
VBA code to do fuzzy lookup though functions in VBA in Microsoft Excel

Function to Fuzzy match LookupValue with entries in                        
column 1 of table specified by TableArray.                                 
TableArray must specify the top left cell of the range to be searched      
The function stops scanning the table when an empty cell in column 1       
is found. 

For each entry in column 1 of the table, FuzzyPercent is called to match LookupValue with the Table entry.                                          

'Rank' is an optional parameter which may take any value > 0 (default 1) and causes the function to return the 'nth' best match (where 'n' is defined by 'Rank' parameter). If the 'Rank' match percentage< NFPercent (Default 5%), #N/A is returned. 

IndexNum is the column number of the entry in TableArray required to be returned, as follows:                                                      

If IndexNum > 0 and the 'Rank' percentage match is >= NFPercent (Default 5%) the column entry indicated by IndexNum is returned.                                                  

if IndexNum = 0 and the 'Rank' percentage match is >= NFPercent (Default 5%) the offset row (starting at 1) is returned.This value can be used directly in the 'Index' function.   
                                                                            
Algorithm can take one of the following values:                            

Algorithm = 1:                                                             
     This algorithm is best suited for matching mis-spellings.              
     For each character in 'String1', a search is performed on 'String2'.   
     The search is deemed successful if a character is found in 'String2'   
     within 3 characters of the current position.                           
     A score is kept of matching characters which is returned as a          
     percentage of the total possible score.                                

Algorithm = 2:                                                             
     This algorithm is best suited for matching sentences, or               
     'firstname lastname' compared with 'lastname firstname' combinations   
     A count of matching pairs, triplets, quadruplets etc. in 'String1' and 
     'String2' is returned as a percentage of the total possible.           

Algorithm = 3: Both Algorithms 1 and 2 are performed.                      

sources: 
https://www.mrexcel.com/board/threads/fuzzy-matching-new-version-plus-explanation.195635/
https://docs.google.com/document/d/1oYVHJcim6POTMtqtHwjq8bwxpMijwtBZhOSUfIZzhmM/edit?pli=1

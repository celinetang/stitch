So this is the file that helps me stitch together all the sheets (sheet name : Detail_5_2_1_0_1) of the excel files in the folder CT1.
The output are two excel sheets : CT1_stitched.xlsx and CT1_extract.xlsx


My problem now is : I want to create a loop and have no idea how.

Context : 
in the root folder, I have multiple folders called CT[i] where [i]=1 to 15 (ie. my folders are called CT1, CT2....CT15)
insite those folders I have multiple excel files, and I want to read one sheet of all these excel files called 'Detail_5_2_X_0_1' where X is the experimental numerotation of the track my battery is on, X is between 1 to 10 (I have 10 tracks in total)
then I want to do my above stated code on it so I can have the output of CT[i]_stitched.xlsx and CT[i]_extract.xlsx

I've managed to read all the excel files in my root folder with :
all_files = glob.glob('./CT*/20*.xls')
however now, I have no idea how to create the loop with i index... 
Do you have any idea ?

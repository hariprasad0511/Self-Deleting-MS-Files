# Self Destructing MS-Office files

This tools helps to enforce Information security by embedding security controls that self deletes any MS-Office based fiels (Wrod, Excel and Powerpoint). This tool is originally planned as a web application that exposes an API, which will consume an MS-Office file as input and will provide a downloadable MS-Office file with embedded security control.

The underlying concept is the usage of Macros to control the MS-Office files. This tool can prevent the voilation of Data Retention Policy uder the new GDPR Security norms. 

### Example Usage Scenarios:

- If you are sending a customer with a price list that is valid for just 3 months, then this tools will auto-delete the price list file after the stipulated time inorder to prevent the customer form using the price list with ill intentions.

- If you are sending a file with Personally Identifiable Information (PII) to an internal division of your company for some data research operations, then this tools will ensure that file is deleted after stipulated time to avoid breaking the GDPR constraints.

## Sample files' usage notes:

- The downloadable files provided here are only for demo purposes. Users are strictly instucted not to use these fiels in a production environments.
- The files deleted by using this Macros cannot be recoverd, as they are deleteld permanantly from the system and is not moved to the Recycle bin.

# INSTRUCTIONS

1. Clone the repository using the following command
  - ```
    git clone
    ```
2. Open the ** macro.txt ** file 
  - In the place of ** Date ** present in the * macro.txt * enter and expiration date for the file
  - Copy the contents of the file 
  
3. Open your choice of MS-Office applicaion (Word/Excel/PPT) and enable developer mode form the options
  - Navigate to the ** Visual Basic ** section under the * Developer * tab
  - Select the ** ThisWorkbook ** option from the * project tree *
  - In the workbook make sure ** Object : Workbook ** and ** ** Procedure : Open ** by selecting the appropriate drop-down buttons
4. Paste the copied code from ** Step: 3 **
5. Save the MS-Office file as ** Macro enable file ** and close the VB_Script editor and the MS-Ofice file.
6. Now open the MS-Office files to see the magic

### Note:

Two reference MS-Excel fiels with following files have been atached for reference.

```
Sample file postdate.xlsm
Sample file predate.xlsm
```


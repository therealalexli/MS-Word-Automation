# MS-Word-Automation

# MS-Word-Automation

This program helps in automating document generation. I have developed it to aid in generating loan documents for a mortgage or similar house loan and it is formatted as such. The main purpose of this readme is to provide supplementary explanation as well as recommendations and reasoning for using and applying this code. 

Modules Used:
-Pandas (to read in and store a dataframe data source)
-Unidecode (to deal with an issue outputting unicode with the MailMerge module
-MailMerge (for populating and creating Merge Fields)
-Datetime (for date)

The code itself is well commented and step by step explanations are provided in the .ipynb file.

We begin by importing the relevant python modules. 

The next step is to read in our datasource into a pandas dataframe:

    df = pd.read_excel('datasource.xlsx', sheet_name='DataSource', header=0)

In my code I have specified a sheet as well which I highly recommend if you are going to be in a corporate work environment. You will most likely be working with excel spreadsheets which if not formatted correctly can be a pain to read into a pandas dataframe. At the same time it may be impractical, unrealistic, or even impossible depending on your company to alter the data source in a meaningful way to work with your code. In my case I was using a very hard to interpret excel spreadsheet. I was able to remedy the situation by creating an additional excel sheet that references the harder-to-interpret first sheet. This way I could organize the information into 2 columns that could easily be converted later on into a python dictionary.

After reading in the data source into a pandas dataframe, we need to convert the values in the dataframe from unicode into string so that they can be outputted later on using the MailMerge module. As of 02-01-2019 there is no support on MS Word or the MailMerge module for dealing with unicode characters and converting to string will be necessary if you have any special characters in your dataframe.

In my dataframe, there are two columns: Labels (to be used as dictionary keys) and Values (the corresponding values). To avoid confusion, I will refer to the column names with a capital L or V and use a lower case l or v when referring to labels and values in the general sense.

We begin by converting values into strings:

    string_values  = infosheet['Values'].astype(str)

Keep in mind that if your values have special characters you will HAVE to follow the steps below I will outline for Labels before converting to string:

    nouni = unidecode(infosheet.Labels[1])
    nouni2 = unidecode(infosheet.Labels[4])

    infosheet.Labels[1] = nouni
    infosheet.Labels[4] = nouni2

    string_labels = infosheet['Labels'].astype(str) 
    
Above, I use the unidecode module to basically hardcode two of my Label values into non-special characters. In my specific case this was done to get rid of a special character apostrophe that appeared in the 2nd and 5th item in my dataframe. As of 02/01/2019 I have not been able to find a better way other than hardcoding to get rid of unwanted values.

After converting the values in each column to strings, we can output a dictionary with string_labels as keys and string_values as values. I also manipulate the keys and values to get rid of all punctuation and spaces with list comprehension. This is to ensure compatibility with the MailMerge module. For example, when trying to pull merge fields from a formatted MS Word document, MailMerge will stop if it encounters a space or a punctuation mark - consequently you will be unable to pull many merge fields using the get_merge_fields function and your code will not work as intended.

    dd = dict( zip( string_labels, string_values)) 

    dictionary = {k.replace(" ",""): v for k,v in dd.items()}   

    dictionary = {k.replace(":",""): v for k,v in dictionary.items()}  

    dictionary = {k.replace("#",""): v for k,v in dictionary.items()} 

    dictionary = {k.replace('(',""): v for k,v in dictionary.items()} 

    dictionary = {k.replace(')',""): v for k,v in dictionary.items()} 

    dictionary ={k.replace('.',''): v for k,v in dictionary.items()} 

    dictionary ={k.replace(',',''): v for k,v in dictionary.items()} 
 
 I recommend checking your dictionary after manipulation to ensure that the keys are consistent with the names of the merge fields you inputted into the word document templates. 
 
 After creating a dictionary we begin Part 2 which is reading in our previously formatted merge field documents. This will be different depending on what you are using the program for but I have included an example below from my code for a privacy policy template:
 
    PrivacyPolicyTemplate = "Privacy Policy Template.docx"
    PrivacyPolicyDocument = MailMerge(PrivacyPolicyTemplate)
    
Above, we first read the template into a variable and then use that variable to create a MailMerge object which we store into another variable called PrivacyPolicyDocument.

Now we can populate the merge fields of our stored variable with our dictionary using our explode operator:

    PrivacyPolicyDocument.merge(**dictionary)
    
It is CRITICAL that the keys in our dictionary match corresponding merge fields down to the case of the letter. It does not matter if the document we are populating does not use all the keys in the dictionary - the explode operator will allow us to access all of the dictionary keys and the merge function will automatically match the keys to the merge fields and populate the corresponding merge field with  the value associated with the key.

Now we finally write our document and name it:

    PrivacyPolicyDocument.write(propertyaddres + " Privacy Policy.docx")

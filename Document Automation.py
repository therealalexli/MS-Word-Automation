#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd   # we use pandas for reading in an excel datasource

from unidecode import unidecode  #  we use unidecode to handle special characters in unicode

from pandas import ExcelWriter
from pandas import ExcelFile

from __future__ import print_function
from mailmerge import MailMerge  # we use MailMerge to populate the fields 
from datetime import date  

infosheet = pd.read_excel('df.xlsx', sheet_name='DataSource', header=0)  ## read the data source into a pandas dataframe

string_values  = infosheet['Values'].astype(str)  ## convert the values from unicode to string for output in MailMerge

nouni = unidecode(infosheet.Labels[1])  ## hardcoded exception for one of the labels
nouni2 = unidecode(infosheet.Labels[4]) ## hardcoded exception for one of the labels

infosheet.Labels[1] = nouni  # replace the Labels value with our hardcoded non-unicode variable
infosheet.Labels[4] = nouni2 # replace the Labels value with our hardcoded non-unicode variable

string_labels = infosheet['Labels'].astype(str) # now that we do not have unicode issues we can convert from unicode to string like we did with Values


# In[ ]:


"""Below we are creating a dictionary and manipulating it so that there are no spaces and punctuation. We
do this to avoid problems with the MailMerge module and MS Word Merge Fields"""


dd = dict( zip( string_labels, string_values)) 

dictionary = {k.replace(" ",""): v for k,v in dd.items()}   

dictionary = {k.replace(":",""): v for k,v in dictionary.items()}  

dictionary = {k.replace("#",""): v for k,v in dictionary.items()} 

dictionary = {k.replace('(',""): v for k,v in dictionary.items()} 

dictionary = {k.replace(')',""): v for k,v in dictionary.items()} 

dictionary ={k.replace('.',''): v for k,v in dictionary.items()} 

dictionary ={k.replace(',',''): v for k,v in dictionary.items()} 


# In[ ]:


"""Below we template the documents into variables"""

LoanAgreementTemplate = "Loan Agreement Template.docx"  
PromissoryNoteTemplate = "Promissory Note Template.docx"
DeedofTrustTemplate = "Deed of Trust Template.docx"
PrivacyPolicyTemplate = "Privacy Policy Template.docx"
FederalEqualCreditOpportunityActNoticeTemplate = "Federal Equal Credit Opportunity Act Notice Template.docx"
EnvironmentalIndemnityAgreementTemplate = "Environmental Indemnity Agreement Template.docx"
LoanAgreementMultipleDrawsTemplate = "Loan Agreement Multiple Draws Template.docx"
LoanAgreementConstructionTemplate = "Loan Agreement Construction Template.docx"
HazardInsuranceDisclosureTemplate = "Hazard Insurance Disclosure Template.docx"
ArbitrationAgreementTemplate = "Arbitration Agreement Template.docx"
CFLNoticeTemplate = 'CFL Notice Template.docx'

"""Now we use the variables and create MailMerge objects"""

LoanAgreementDocument = MailMerge(LoanAgreementTemplate) 
PromissoryNoteDocument = MailMerge(PromissoryNoteTemplate)
DeedofTrustDocument = MailMerge(DeedofTrustTemplate)
PrivacyPolicyDocument = MailMerge(PrivacyPolicyTemplate)
FederalEqualCreditOpportunityActNoticeDocument = MailMerge(FederalEqualCreditOpportunityActNoticeTemplate)
EnvironmentalIndemnityAgreementDocument = MailMerge(EnvironmentalIndemnityAgreementTemplate)
LoanAgreementMultipleDrawsDocument = MailMerge(LoanAgreementMultipleDrawsTemplate)
LoanAgreementConstructionDocument = MailMerge(LoanAgreementConstructionTemplate)
HazardInsuranceDisclosureDocument = MailMerge(HazardInsuranceDisclosureTemplate)
ArbitrationAgreementDocument = MailMerge(ArbitrationAgreementTemplate)
CFLNoticeDocument = MailMerge(CFLNoticeTemplate)


# In[ ]:


propertyaddress = dictionary['PropertyAddress']   ## Since the application of this program is real estate, we store the property address as a variable for naming


# In[ ]:


"""Populate the MailMerge object with the dictionary. We use the explode operator so the dictionary can automatically match and populate the Merge Fields.
This will only work if the merge fields EXACTLY MATCH the dictionary keys."""

LoanAgreementDocument.merge(**dictionary)
PromissoryNoteDocument.merge(**dictionary)
DeedofTrustDocument.merge(**dictionary)
PrivacyPolicyDocument.merge(**dictionary)
FederalEqualCreditOpportunityActNoticeDocument.merge(**dictionary)
EnvironmentalIndemnityAgreementDocument.merge(**dictionary)
LoanAgreementMultipleDrawsDocument.merge(**dictionary)
LoanAgreementConstructionDocument.merge(**dictionary)
HazardInsuranceDisclosureDocument.merge(**dictionary)
ArbitrationAgreementDocument.merge(**dictionary)
CFLNoticeDocument.merge(**dictionary)
##ConditionalLoanApprovalDocument.merge(**dictionary)


"""Finally, we create new documents using the write function. They should appear in the same folder as your program."""

PromissoryNoteDocument.write(propertyaddress +' Promissory Note.docx')
LoanAgreementDocument.write(propertyaddress +' Loan Agreement LEGAL.docx') 
DeedofTrustDocument.write(propertyaddress +' Deed of Trust.docx')
PrivacyPolicyDocument.write(propertyaddress +' Privacy Policy.docx')
FederalEqualCreditOpportunityActNoticeDocument.write(propertyaddress +' Federal Equal Credit Opportunity Act Notice .docx')
EnvironmentalIndemnityAgreementDocument.write(propertyaddress +' Environmental Indemnity Agreement.docx')
LoanAgreementMultipleDrawsDocument.write(propertyaddress +' Loan Agreement Multiple Draws.docx')
LoanAgreementConstructionDocument.write(propertyaddress +' Loan Agreement Construction Document.docx')
HazardInsuranceDisclosureDocument.write(propertyaddress +' Hazard Insurance Disclosure Document.docx')
ArbitrationAgreementDocument.write(propertyaddress +' Arbitration Agreement.docx')
CFLNoticeDocument.write(propertyaddress+' CFL Notice.docx')
###ConditionalLoanApprovalDocument.write(propertyaddress+ ' Conditional Loan Approval.docx')


# In[ ]:





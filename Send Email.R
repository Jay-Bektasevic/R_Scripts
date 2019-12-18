

#########################################################################################################
#                                                                                                       #
# Purpose:       SEND EMAIL                                                                             #
#                                                                                                       #
# Author:        Jay Bektasevic                                                                         #
# Contact:       jbektasevic@clearpathmutual.com                                                        #
# Client:        Accounting                                                                             #
#                                                                                                       #
# Code created:  2019-05-16                                                                             #
# Last updated:  2019-05-16                                                                             #
# Source:        K:/Departments/Accounting/Jay Bektasevic/Book of Business Analysis/Class Code Analysis #
#                                                                                                       #
# Comment:       Send Outlook email from R.                                                             #
#                                                                                                       #
#########################################################################################################


library(RDCOMClient)

############################################# Init COM API ##############################################
OutApp <- COMCreate("Outlook.Application")

########################################### Create an Email  ############################################
outMail = OutApp$CreateItem(0)

##################################### Configure  Email Parameters  ######################################
outMail[["To"]] = "jbektasevic@clearpathmutual.com"
outMail[["subject"]] = "TEST!!!"
outMail[["body"]] = "TEST!!!"

########################################### Add Attachments  ############################################
outMail[["Attachments"]]$Add("K:\\Departments\\Accounting\\Jay Bektasevic\\Book of Business Analysis\\Class Code Analysis\\class_code_loss_ratio.xlsx")

############################################## Send Email  ##############################################                   
outMail$Send()
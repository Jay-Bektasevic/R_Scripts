
####################################################################################################
#                                                                                                  #
# Purpose:       Code Comment Format                                                               #
#                                                                                                  #
# Author:        Jay Bektasevic                                                                    #
# Contact:       jbektasevic@clearpathmutual.com                                                   #
# Client:        Jay Bektasevic                                                                    #
#                                                                                                  #
# Code created:  2019-03-29                                                                        #
# Last updated:  2019-03-29                                                                        #
# Source:        K:/Departments/Accounting/Jay Bektasevic/R Helper Scripts                         #
#                                                                                                  #
# Comment:       To be used in formating of R code                                                 #
#                                                                                                  #
####################################################################################################

library(commentr)

# Header comment

header_comment(
    "ALL CLAIMS DUMP",
    "This script is used in ...",
    author = "Jay Bektasevic",
    contact = "jbektasevic@clearpathmutual.com",
    client = "Accounting",
    width = 105
    )

####################################################################################################
#                                                                                                  #
#                                      A small block comment                                       #
#                                                                                                  #
####################################################################################################

block_comment("Here we will re-run the SQL query to get the policies for the rest of the year but, this time 
              we will use the previous year's rating detail and apply the new rates to it. ",
              width = 105)


####################################### Comment on one line ########################################

line_comment("PLOT",  width = 105) 

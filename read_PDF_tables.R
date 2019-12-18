
#     =================================================================================================    
#     Script Name:  Read PDF tables
#     =================================================================================================    
#
#     Create Date:  2018-12-19
#     Author:       Jay Bektasevic
#     Author Email: jbektasevic@clearpathmutual.com
# 
#     Description: Excellent script to help with the extraction of tables from PDF files via the "tabulizer" package.
# 
# 
#     Call by:      [Application Name]
#                   [Job]
#                   [Interface]
# 
# 
#     Used By:      General Purpose
#                 
#     =================================================================================================
#                                       SUMMARY OF CHANGES
#     -------------------------------------------------------------------------------------------------
#     Date:                   Author:           Comments:
#     ----------------------- ----------------- -------------------------------------------------------
#     2018-12-19   
#     =================================================================================================



#   Increase the heap size before the rJava package is loaded - it's a dependency for "xlsx" package.
    options(java.parameters = "-Xmx16000m")

#   List of packages to load:
    packages <- c('tabulizer', 'dplyr', 'xlsx', 'tidyr', 'stringr' )

    new_packages <-
        packages[!(packages %in% installed.packages()[, "Package"])]
    
    if (length(new_packages))
        install.packages(new_packages)

#   Load Neccessary Packages
    sapply(packages, require, character.only = TRUE)
    
#   NOTE
#   In case you have issues installing "tabulizer" package - use the following line of code 
#   (you may need to install "remotes" package beforehand). 
#   remotes::install_github(c("ropensci/tabulizerjars", "ropensci/tabulizer"), INSTALL_opts = "--no-multiarch", dependencies = c("Depends", "Imports"))

#   Location of pdf file
    setwd("C:/Users/jbektasevic/Downloads")
    
#   PDF file
    pdf.file <- "KEMI - KY - Rate- Eff 1-1-2020a.pdf"

#   Next, we will use the extract_tables() function from tabulizer to extract table from pdf.
#   I am using the default parameters for extract_tables. These are guess and method. 
#   I'll leave guess set to TRUE, which tells tabulizer that we want it to figure out the 
#   locations of the tables on its own. We could set this to FALSE if we want to have more granular 
#   control, but for this application we don't need to. We leave the method argument set to "matrix", 
#   which will return a list of matrices (one for each pdf page). This could also be set to 
#   return data frames instead.
#   utilize function locate_areas() to get area coordinates.
    
   
    out <- extract_tables(pdf.file, 
                          # pages = c(29:32),
                          # area = list(
                          #     c( 97.30286,  66.13714, 676.59429, 543.60000 ),
                          #     c(115.4057,  68.4000, 681.1200, 548.1257),
                          #     c(115.40571,  66.13714,681.12000 ,543.60000),
                          #     c(115.40571, 61.61143, 676.59429, 536.81143)
                          #     
                          # ),
                          guess = FALSE
            )

#   Now we have a list object called out, with each element a matrix representation of a page of the 
#   pdf table. We want to combine these into a single data matrix containing all of the data. We can 
#   do so most elegantly by combining do.call and rbind, passing it our list of matrices.
    
    final <- do.call(rbind, out)

#   After doing so, the first few rows of the matrix may contain the headers, which have not been 
#   formatted well since as they can take up multiple rows of the pdf table. To fix it, we'll have 
#   to convert the matrix into a data.frame dropping the first row (in this case). Then I create a 
#   character vector containing the formatted headers and use that as the column names.

    final <- as.data.frame(final)

#   Column names form the first row.
    headers <- as.vector(t(as.vector(final[1,])))

#   Apply custom column names
    names(final) <- headers

#   Export files to excel workbook         

    write.xlsx(
        final,
        file = "KEMI - KY - Rate-Rule - Eff 1-1-2019.xlsx",
        sheetName = "KEMI",
        row.names = FALSE,
        append = TRUE
    )
###############################################################
 
    
#   Location of pdf file
    setwd("C:/Users/jbektasevic/Downloads")
    
#   PDF file
    pdf.file <- "KEMI - KY - Rate- Eff 1-1-2020a.pdf"    
    
    out <- extract_tables(pdf.file, guess = FALSE )
    
    col_names <- c(
        'Class',
        'Class Code Name',
        'Current Standard Rate',
        'Current Prefered Rate',
        'Current Minimum Premium',
        'Proposed Standard Rate',
        'Proposed Preferred Rate',
        'Proposed Minimum Premium',
        'Standard Change',
        'Preferred Change'
        
    )
    
    x1 <- as.data.frame(out[[1]][5:68,])
    x1 <- separate(x1, V7, sep = " ", into = c("New", "New1"), remove = TRUE)
    x1 <- separate(x1, V8, sep = " ", into = c("New2", "New3"), remove = TRUE)
    names(x1)[] <- col_names
    
    
    x2 <- as.data.frame(out[[2]][1:68,])
    x2[43,4]<- 5.42 # fix a rate in that particular cell
    x2 <- x2[,-3] # remove blank column
    names(x2)[]  <- col_names
    
    
    x3 <- as.data.frame(out[[3]][1:68,])
    names(x3)[]  <- col_names
    
    x4 <- as.data.frame(out[[4]][1:68,])
    names(x4)[]  <- col_names
    
    x5 <- as.data.frame(out[[5]][1:68,])
    x5$V3 <- str_extract(x5$V2, '[^ ]+$') # extracts last characters after the space in a string, which in this case is the rate
    x5[9,3] <- 11.93
    x5[18,3] <- 1.23
    x5[19,3] <- 5.98
    x5[27,3] <- 8.05 
    x5[39,3] <- 5.15
    x5[41,3] <- 1.35
        
    names(x5)[]  <- col_names
    
    x6 <- as.data.frame(out[[6]][1:68,])
    x6 <- x6[,-4]
    x6$V3 <- as.character(x6$V3)  # convert to character instead of factor to manipulate the values.
    x6[49,3] <- 1.29
    x6[51,3] <- 1.81
    x6[68,3] <- 17.10
    names(x6)[]  <- col_names

    x7 <- as.data.frame(out[[7]][1:68,])    
    names(x7)[]  <- col_names
    
    
    x8 <- as.data.frame(out[[8]][1:68,]) 
    x8$V3 <- str_extract(x8$V2, '[^ ]+$') # Left Here
    
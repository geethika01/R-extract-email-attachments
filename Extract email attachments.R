# 
# Title : Download outlook email attachments and consolidate into one table
# Description : The script 
#   1) extracts emails from outlook inbox by a given subject after a specified 
#       received date
#   2) Save the excel attachement by 
#           1. File name of the attachment
#           2. Name of the sender
#   3) Read the saved attachements and consolidate data into one final table
#
# Written by : Geethika Wijewardena
# Date : 11/10/2019
#------------------------------------------------
# Setup
#------------------------------------------------
rm(list=ls())
rm(list = ls())
library(RDCOMClient)
library(dplyr)
library(stringr)

working_dir<-"C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments/"

#--------------------------------------------
# Extract emails from outlook
#--------------------------------------------
# Create a new instance of Outlook COM server class
outlook_app <- COMCreate("Outlook.Application")
# Create a search object to search the mail box by given criteria (e.g. subject)
search <- outlook_app$AdvancedSearch(
  "Inbox",
  "urn:schemas:httpmail:subject = 'REA0001 - Measurements'"
)
# Allow some time for the search to complete
Sys.sleep(5)
results <- search$Results()

#-------------------------------------------------------------------------------
# Approach 1
# Extract emails and save the attachment by the name of the attachment
#-------------------------------------------------------------------------------

# Filter search results by receive date
for (i in 1:results$Count()){
  receive_date <- as.Date("1899-12-30") + floor(results$Item(i)$ReceivedTime())
  if(receive_date >= as.Date("2019-10-09")) {
    # Get the attachment of each email and save it by the name of the attachment
    #   in a given file path
    email <- results$Item(i)
    attachment_file <- paste0(working_dir,email$Attachments(1)[['DisplayName']])
    email$Attachments(1)$SaveAsFile(attachment_file)
    
    # Read each attachment and assign data into a variable (which is the filename)
    #   generated dynamically, 
    df_name <- str_sub(email$Attachments(1)[['DisplayName']],1,-6)
    data <- readxl::read_excel(attachment_file, col_types =c("date", "numeric"),
                               col_names = T) %>% 
      rename(!!df_name := "Case")%>% 
      mutate(Hour = str_sub(as.character(Hour),11,nchar(as.character(Hour))))
    assign(df_name, data)
  }
}

# Consolidate all dataframes into one
dat <- lapply(ls(pattern="REA"), function(x) get(x)) %>% 
  purrr::reduce(full_join, by = "Hour") 

#-------------------------------------------------------------------------------
# Approach 2
# Extract emails and save the attachment by the name of the sender
#-------------------------------------------------------------------------------
getDataFromEmailAtt<- function(results, i){
  # Function to extract data from email attachement, save it it a specified
  # directory by the name of the sender, read the saved excel file and return 
  # a dataframe with a given colum named by the sender's name.
  # Args: results - object returned by search$Results() of RDCOMClient for 
  #                 outlook applications.
  #       i - order number of the extracted emails in the results object 
  # Returns: Dataset of the email attachment with given column renamed by the 
  #          sender's name
  # 
  receive_date <- as.Date("1899-12-30") + floor(results$Item(i)$ReceivedTime())
  if(receive_date >= as.Date("2019-10-09")) {
    # Get the attachment of each email and save it by the name of the attachment
    #   in a given file path
    email <- results$Item(i)
    attachment_file <- paste0(working_dir,email[['SenderName']],'.xlsx')
    email$Attachments(1)$SaveAsFile(attachment_file)
    
    # Read each attachment and assign data into a variable (which is the filename)
    # generated dynamically, 
    df_name <- email[['SenderName']]
    data <- readxl::read_excel(attachment_file, col_names = T) %>% 
      rename(!!df_name := "Measurement")%>% 
      mutate(Hour = str_sub(as.character(Hour),11,nchar(as.character(Hour))))
  return(data)
  }
}

# Get the first dataset
dat <- getDataFromEmailAtt(results, i=1)

# Append datasets of the other emails
for (i in 2:results$Count()){
  data <- getDataFromEmailAtt(results, i)
  dat <- dat %>% inner_join(data, by=c('Hour'))
}

#-----------------------------------------------
# End
#-----------------------------------------------

# 
# Title : Download outlook email attachments and consolidate into one table
# Written by : Geethika Wijewardena
# Date : 11/10/2019
#---------------------------------------------------------------------------------

rm(list=ls())
rm(list = ls())
library(RDCOMClient)
library(dplyr)
library(stringr)

# Create COM Object to access Outlook Application
outlook_app <- COMCreate("Outlook.Application")
# Create a search object to search the mail box by given criteria (e.g. subject)
search <- outlook_app$AdvancedSearch(
  "Inbox",
  "urn:schemas:httpmail:subject = 'REA0001 - Measurements'"
)
# Allow some time for the search to complete
Sys.sleep(5)
results <- search$Results()

# Filter search results by receive date
for (i in 1:results$Count()){
  receive_date <- as.Date("1899-12-30") + floor(results$Item(i)$ReceivedTime())
  if(receive_date >= as.Date("2019-10-09")) {
    
    # Get the attachment of each email and save it by the name of the attachment in a given file path
    email <- results$Item(i)
    attachment_file <- paste0("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments/",email$Attachments(1)[['DisplayName']])
    email$Attachments(1)$SaveAsFile(attachment_file)
    
    # Read each attachment and assign data into a variable (which is the filename) generated dynamically, 
    df_name <- str_sub(email$Attachments(1)[['DisplayName']],1,-6)
    data <- data <- readxl::read_excel(attachment_file, col_types = c("date", "numeric"), col_names = T) %>% 
      rename(!!df_name := "Measurement")%>% 
      mutate(Hour = str_sub(as.character(Hour),11,nchar(as.character(Hour))))
    assign(df_name, data)
  }
}

# Consolidate all dataframes into one
dat <- lapply(ls(pattern="REA"), function(x) get(x)) %>% purrr::reduce(full_join, by = "Hour") 

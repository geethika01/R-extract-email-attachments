rm(list=ls())
library(RDCOMClient)
library(openxlsx)
library(dplyr)
library(stringr)
library(tidyverse)

setwd("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments")
outlook_app <- COMCreate("Outlook.Application")
search <- outlook_app$AdvancedSearch(
  "Inbox",
  "urn:schemas:httpmail:subject = 'GEO0001 - Marks'"
)
results <- search$Results()

for (i in 1:results$Count()){
  receive_date <- as.Date("1899-12-30") + floor(results$Item(i)$ReceivedTime())
  if(receive_date == as.Date("2019-10-09")) {
    email <- results$Item(i)
    attachment_file <- paste0("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments/",email$Attachments(1)[['DisplayName']])
    email$Attachments(1)$SaveAsFile(attachment_file)
    df_name <- str_sub(email$Attachments(1)[['DisplayName']],1,-6)
    data <- read.xlsx(attachment_file) %>% rename(!!df_name := "Score")
    assign(df_name, data)
  }
}

# Join scores
dat <- lapply(ls(pattern="marker"), function(x) get(x)) %>% reduce(full_join, by = "Question") 

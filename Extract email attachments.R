library(RDCOMClient)
library(openxlsx)
library(dplyr)
library(stringr)

setwd("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments")
outlook_app <- COMCreate("Outlook.Application")
search <- outlook_app$AdvancedSearch(
  "Inbox",
  "urn:schemas:httpmail:subject = 'GEOM0001 - Marks'"
)
results <- search$Results()

for (i in 1:results$Count()){
  email <- results$Item(i)
  attachment_file <- paste0("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments/",email$Attachments(1)[['DisplayName']])
  email$Attachments(1)$SaveAsFile(attachment_file)
  df_name <- str_sub(email$Attachments(1)[['DisplayName']],1,-6)
  data <- read.xlsx(attachment_file) %>% rename(!!df_name := "Score")
  assign(df_name, data)
  
}

# Join scores
dat <- marker_1 %>% inner_join(marker_2, by="Question")

library(RDCOMClient)
library(openxlsx)
library(dplyr)

setwd("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments")
outlook_app <- COMCreate("Outlook.Application")
search <- outlook_app$AdvancedSearch(
  "Inbox",
  "urn:schemas:httpmail:subject = 'GEOM0001 - Marks'"
)
results <- search$Results()

# Read the first file
email <- results$Item(1)
attachments_obj <- email[['attachments']]
attachment_file
email$Attachments(1)$SaveAsFile(paste0("C:/Users/geethika.wijewardena/Workspace/R-extract-email-attachments/", email$Attachments(1)[['DisplayName']] ))
data <- read.xlsx(attachment_file)



attachments.obj <- email[['attachments']] # Gets the attachment object
attachments <- character() # Create an empty vector for attachment names

if(attachments.obj$Count() > 0){ # Check if there are attachments
  for(i in c(1:attachments.obj$Count())){ # Loop through attachments
    attachments <- append(attachments, attachments.obj$Item(i)[['DisplayName']]) # Add attachment name to vector
  }
}

print(attachments)

# Read other attachments
for (i in 2:results$Count()) {
  email <- results$Item(i)
  attachment_file <- tempfile()
  email$Attachments(1)$SaveAsFile(attachment_file)
  dat <- read.xlsx(attachment_file)
    
  }
}
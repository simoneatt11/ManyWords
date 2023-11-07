# Simone Attanasio - 20231107
# the output will be a word document

# packages needed
install.packages("officer")
install.packages("fs")
install.packages("readtext")

# activate packages
library(officer)
library(fs)
library(readtext)
library(dplyr)

#set as working directory the folder containing all your doc files
setwd("path/to/folder")

# Replace with the actual path to your folder
folder_path <- "/path/to/folder"

doc <- read_docx()
# create an empty variable in R - the loop will append here

file_list <- list.files(folder_path, pattern = "\\.(docx|doc)$", full.names = TRUE)
# create a variable containing the list of files present in the directory "folder_path"

# go through every file in the file_list and append everything in doc
for (file_path in file_list) {
  content <- readtext(file_path)

  file_name <- basename(file_path)

  doc <- doc %>%
    body_add("Title: ") %>%
    body_add(file_name) %>%
    body_add("\nFile Content:\n") %>%
    body_add(content$text) %>%
    body_add("\n\n")
}

# create a new document in a file path with a new file name
final_doc_path <- "/path/to/folder/Final_doc.docx"  #Replace with the desired file path and filename

# print doc in the final document
print(doc, target = final_doc_path)

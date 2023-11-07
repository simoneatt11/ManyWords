# Simone Attanasio - 20231107
# doublecheck packages that are needed
# doublecheck steps that are correct - write a good documentation

install.packages("officer")
install.packages("fs")
install.packages("readtext")

library(officer)
library(fs)
library(readtext)
library(dplyr)

folder_path <- "/Users/simoneattanasio/Downloads/Scritti/test"
# Replace with the actual path to your folder

doc <- read_docx()
# create an empty variable in R - the loop will append here

file_list <- list.files(folder_path, pattern = "\\.(docx|doc)$", full.names = TRUE)
# create a variable containing the list of files present in the directory "folder_path"

# explain the loop ...
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

# try this tomorrow -> doc <- doc %>%
# -> body_add("File Name: ", file_name, "\nFile Content:\n", content$text, "\n\n")

final_doc_path <- "Tutto_2018.docx"  # Replace with the desired file path
print(doc, target = final_doc_path)

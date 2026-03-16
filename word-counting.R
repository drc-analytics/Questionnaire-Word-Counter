library(openxlsx)
library(dplyr)
library(stringr)



df_labs<-read.xlsx("Data/Musk+FG+2+Day+2_March+10,+2026_16.46.xlsx")
df_vals<-read.xlsx("Data/Musk+FG+2+Day+2_March+10,+2026_16.46 (1).xlsx")
#read.xlsx("Musk+FG+2+Day+2_March+10,+2026_16.46 (1).xlsx")

#not great, but right now cleaning out any irrelevant columns by hand. 
#there is probably a way to do this in order to not have to update this per survey
#or per project, but that will take more workarounds
remov_col<-c("StartDate",             "EndDate",               "Status",                "IPAddress",            
 "Progress",              "Duration.(in.seconds)", "Finished",              "RecordedDate",         
 "ResponseId",            "RecipientLastName",     "RecipientFirstName",    "RecipientEmail",       
 "ExternalReference",     "LocationLatitude",      "LocationLongitude",     "DistributionChannel" , 
 "UserLanguage", "Q1", "Q68", "Q69", "Q109","Q110",
 "Q113",
 "Q157",
 "Q97",
 "Q98",
 "Q99",
 "Q3_1",
 "Q3_2",
 "Q3_3",
 "Q3_4",
 "Q3_5",
 "Q3_6",
 "Q3_7",
 "Q4_1",
 "Q71_1",
 "Q71_2",
 "Q71_3",
 "Q71_4",
 "Q71_5",
 "Q71_6",
 "Q71_7",
 "Q72_1",
 "Q77_1",
 "Q77_2",
 "Q77_3",
 "Q77_4",
 "Q77_5",
 "Q77_6",
 "Q77_7",
 "Q78_1",
 "Q83_1",
 "Q83_2",
 "Q83_3",
 "Q83_4",
 "Q83_5",
 "Q83_6",
 "Q83_7",
 "Q84_1",
 "Q90_1",
 "Q89_1",
 "Q89_2",
 "Q89_3",
 "Q89_4",
 "Q89_5",
 "Q89_6",
 "Q89_7",
 "Q160_1",
 "Q160_2",
 "Q160_3",
 "Q160_4",
 "Q160_5",
 "Q160_6",
 "Q160_7",
 "Q161_1",
 "Q101",
 "Q154",
 "Q163",
 "Q74",
 "Q80",
 "Q86",
 "Q92"
 )


df_vals<-df_vals[,!(colnames(df_vals)%in% remov_col)]




df_vals[] <- lapply(df_vals, function(x) {
  if (is.character(x)) gsub('xml:space="preserve">', '', x, fixed = TRUE) else x
})




df_vals <- df_vals %>%
  mutate(
    across(
      where(is.character),
      ~ .x |>
        gsub('xml:space="preserve">', '', x = _, fixed = TRUE) |>
        gsub('&#xa;', ' ', x = _, fixed = TRUE)
    )
  )





# convert blank values to NA
df_vals <- df_vals %>%
  mutate(
    across(
      where(is.character),
      ~ ifelse(str_trim(.x) == "", NA, .x)
    )
  )
df_vals$total_words<-0
df_vals$words_per_Q<-0



# exclude summary columns and Q2 from counted totals
question_cols<-setdiff(names(df_vals), c("total_words", "words_per_Q", "Q2"))

#thanks Claude for the code assist
#Basically we're setting all of the question columns we want counted to characters
#then, we're making sure the column is not NA, trimming entries by spaces and counting all words
df_vals <- df_vals %>%
  mutate(across(all_of(question_cols), as.character)) %>%
  rowwise() %>%
  mutate(
    answered_values = list(c_across(all_of(question_cols))),
    answered_values = list(answered_values[!is.na(answered_values) & str_trim(answered_values) != ""]),
    n_answered = length(answered_values),
    total_words = sum(str_count(answered_values, "\\S+")),
    words_per_Q = ifelse(n_answered > 0, total_words / n_answered, NA_real_)
  ) %>%
  ungroup() %>%
  select(-answered_values)


# set column names from first row
#colnames(df_vals) <- as.character(df_vals[1, ])
#above line works, but makes it impossible to output as CSV file because  of columns with same name
df_vals$total_words[1]<-"total_words"
df_vals$words_per_Q[1]<-"words_per_Q"
df_vals$n_answered[1]<-"n_answered"

# remove the first row
#df_vals <- df_vals[-1, ]
#just run above if not saving as CSV and cleaning for optics

write.csv(df_vals, "Output/df_Qs_vals.csv")

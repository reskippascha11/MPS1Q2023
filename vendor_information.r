library(tidyverse)
#library(readxl)
#install.packages("writexl")
#install.packages("xlsx")
library(xlsx)
library(writexl)

order_totals_1 <-
  read.xlsx("1st Grade Order Totals.xlsx", sheetIndex = 1) %>%
  rename(Item_Number = 2, Total_Requested = 5) %>%
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)
order_totals_2 <-
  read.xlsx("2nd Grade Order Totals.xlsx", sheetIndex = 1) %>%
  rename(Item_Number = 2, Total_Requested = 5) %>%
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)
order_totals_3 <-
  read.xlsx("3rd Grade Order Totals.xlsx", sheetIndex = 1) %>%
  rename(Item_Number = 2, Total_Requested = 5) %>%
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)
order_totals_4 <-
  read.xlsx("4th Grade Order Totals.xlsx", sheetIndex = 1) %>%
  rename(Item_Number = 2, Total_Requested = 5) %>%
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)
order_totals_5 <-
  read.xlsx("5th Grade Order Totals.xlsx", sheetIndex = 1) %>%
  rename(Item_Number = 2, Total_Requested = 5) %>%
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)
order_totals_K <-
  read.xlsx("Kindergarten Order Totals.xlsx", sheetIndex = 1) %>%
  rename(Item_Number = 2, Total_Requested = 5) %>%
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_rbind_all_grades_1 <- rbind(
  order_totals_1,
  order_totals_2,
  order_totals_3,
  order_totals_4,
  order_totals_5,
  order_totals_K
)

#View(order_totals_1)
#View(order_totals_2)
#View(order_totals_3)
#View(order_totals_4)
#View(order_totals_5)
#View(order_totals_K)

#View(order_totals_rbind_all_grades_1)

order_totals_rbind_all_grades_2 <-
  order_totals_rbind_all_grades_1[!is.na(order_totals_rbind_all_grades_1$Total_Requested), ] %>%
  arrange(Vendor, Item_Number, Description)

#View(order_totals_rbind_all_grades_2)

sum_total_requested <- 0

vendor_info <- tibble(
  Vendor = character(),
  Item_Number = character(),
  Description = character(),
  Pkg = character(),
  Total_Requested = numeric()
)
#View(vendor_info)

current_Vendor <-  character()
current_Item_Number <-  character()
current_Description <-  character()
current_Pkg <-  character()
current_Total_Requested <-  numeric()

for (i in 1:nrow(order_totals_rbind_all_grades_2)) {
  
  if (i == 1) {
    current_Vendor = order_totals_rbind_all_grades_2$Vendor[[i]]
    current_Item_Number = order_totals_rbind_all_grades_2$Item_Number[[i]]
    current_Description = order_totals_rbind_all_grades_2$Description[[i]]
    current_Pkg = order_totals_rbind_all_grades_2$Pkg[[i]]
    sum_total_requested <-
      order_totals_rbind_all_grades_2$Total_Requested[[i]]
    next # first row has been processed so skip to next iteration
  }
  
  if (i == nrow(order_totals_rbind_all_grades_2)) {
    
    print(paste0("order_totals_rbind_all_grades_2$Item_Number[[i]] = ",order_totals_rbind_all_grades_2$Item_Number[[i]]))
    print(paste0("current_Item_Number =  ", current_Item_Number))
        
    if (order_totals_rbind_all_grades_2$Item_Number[[i]] == current_Item_Number) {
      sum_total_requested <-
        sum_total_requested  + order_totals_rbind_all_grades_2$Total_Requested[[i]]
      vendor_info <- vendor_info %>%  add_row(
        Vendor = order_totals_rbind_all_grades_2$Vendor[[i]],
        Item_Number = order_totals_rbind_all_grades_2$Item_Number[[i]],
        Description = order_totals_rbind_all_grades_2$Description[[i]],
        Pkg = order_totals_rbind_all_grades_2$Pkg[[i]],
        Total_Requested = sum_total_requested)      
    } else {
      vendor_info <- vendor_info %>%  add_row(
        Vendor = current_Vendor,
        Item_Number = current_Item_Number,
        Description = current_Description,
        Pkg = current_Pkg,
        Total_Requested = sum_total_requested)      
      vendor_info <- vendor_info %>%  add_row(
        Vendor = order_totals_rbind_all_grades_2$Vendor[[i]],
        Item_Number = order_totals_rbind_all_grades_2$Item_Number[[i]],
        Description = order_totals_rbind_all_grades_2$Description[[i]],
        Pkg = order_totals_rbind_all_grades_2$Pkg[[i]],
        Total_Requested = order_totals_rbind_all_grades_2$Total_Requested[[i]])
    }
    break # last row has been processed, so leave the for loop
  }
  
  if (order_totals_rbind_all_grades_2$Item_Number[[i]] == current_Item_Number) {
    sum_total_requested <-
      sum_total_requested  + order_totals_rbind_all_grades_2$Total_Requested[[i]]
  } else {
    vendor_info <- vendor_info %>%  add_row(
      Vendor = current_Vendor,
      Item_Number = current_Item_Number,
      Description = current_Description,
      Pkg = current_Pkg,
      Total_Requested = sum_total_requested)
    current_Vendor = order_totals_rbind_all_grades_2$Vendor[[i]]
    current_Item_Number = order_totals_rbind_all_grades_2$Item_Number[[i]]
    current_Description = order_totals_rbind_all_grades_2$Description[[i]]
    current_Pkg = order_totals_rbind_all_grades_2$Pkg[[i]]
    sum_total_requested <-
      order_totals_rbind_all_grades_2$Total_Requested[[i]]
  }
  
}

write_xlsx(vendor_info, "vendor_info_v6.xlsx")

  

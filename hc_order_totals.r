library(tidyverse)
library(xlsx)

order_totals <-read.xlsx("Kindergarten Order Totals.xlsx", sheetIndex = 1) %>% 
  rename(Item_Number = 2, Total_Requested = 5) %>% 
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_2 <- order_totals[!is.na(order_totals$Total_Requested), ]

order_totals_3 <- order_totals_2 %>% 
  mutate(Total_Requested_Num = as.numeric(Total_Requested))

for (i in 1:nrow(order_totals_3)) {
  if (is.na(order_totals_3$Total_Requested_Num[[i]])) {
    #print(order_totals_3$Item_Number[[i]])
    #print(order_totals_3$Description[[i]])
    #print(order_totals_3$Total_Requested_Num[[i]])
    
    print("Kindergarten")
    print(order_totals_3$Item_Number[[i]])
    print(order_totals_3$Description[[i]])
    print(order_totals_3$Total_Requested[[i]])   
    
    #print(str_c(order_totals_3$Item_Number[[i]],
    #      order_totals_3$Description[[i]],
    #      order_totals_3$Total_Requested[[i]],
    #      sep = " "
    #      ))
  }
}

order_totals <-read.xlsx("1st Grade Order Totals.xlsx", sheetIndex = 1) %>% 
  rename(Item_Number = 2, Total_Requested = 5) %>% 
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_2 <- order_totals[!is.na(order_totals$Total_Requested), ]

order_totals_3 <- order_totals_2 %>% 
  mutate(Total_Requested_Num = as.numeric(Total_Requested))

for (i in 1:nrow(order_totals_3)) {
  if (is.na(order_totals_3$Total_Requested_Num[[i]])) {
    #print(order_totals_3$Item_Number[[i]])
    #print(order_totals_3$Description[[i]])
    #print(order_totals_3$Total_Requested_Num[[i]])
    
    print("1st Grade")    
    print(str_c(order_totals_3$Item_Number[[i]],
                order_totals_3$Description[[i]],
                order_totals_3$Total_Requested[[i]],
                sep = " "
    ))
  }
}


order_totals <-read.xlsx("2nd Grade Order Totals.xlsx", sheetIndex = 1) %>% 
  rename(Item_Number = 2, Total_Requested = 5) %>% 
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_2 <- order_totals[!is.na(order_totals$Total_Requested), ]

order_totals_3 <- order_totals_2 %>% 
  mutate(Total_Requested_Num = as.numeric(Total_Requested))

for (i in 1:nrow(order_totals_3)) {
  if (is.na(order_totals_3$Total_Requested_Num[[i]])) {
    #print(order_totals_3$Item_Number[[i]])
    #print(order_totals_3$Description[[i]])
    #print(order_totals_3$Total_Requested_Num[[i]])
    
    print("2nd Grade")    
    print(str_c(order_totals_3$Item_Number[[i]],
                order_totals_3$Description[[i]],
                order_totals_3$Total_Requested[[i]],
                sep = " "
    ))
  }
}

order_totals <-read.xlsx("3rd Grade Order Totals.xlsx", sheetIndex = 1) %>% 
  rename(Item_Number = 2, Total_Requested = 5) %>% 
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_2 <- order_totals[!is.na(order_totals$Total_Requested), ]

order_totals_3 <- order_totals_2 %>% 
  mutate(Total_Requested_Num = as.numeric(Total_Requested))

for (i in 1:nrow(order_totals_3)) {
  if (is.na(order_totals_3$Total_Requested_Num[[i]])) {
    #print(order_totals_3$Item_Number[[i]])
    #print(order_totals_3$Description[[i]])
    #print(order_totals_3$Total_Requested_Num[[i]])
    
    print("3rd Grade")
    print(order_totals_3$Item_Number[[i]])
    print(order_totals_3$Description[[i]])
    print(order_totals_3$Total_Requested[[i]])
    
    #print(str_c(order_totals_3$Item_Number[[i]],
    #            order_totals_3$Description[[i]],
    #            order_totals_3$Total_Requested[[i]],
    #            sep = " "
    #))
  }
}

order_totals <-read.xlsx("4th Grade Order Totals.xlsx", sheetIndex = 1) %>% 
  rename(Item_Number = 2, Total_Requested = 5) %>% 
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_2 <- order_totals[!is.na(order_totals$Total_Requested), ]

order_totals_3 <- order_totals_2 %>% 
  mutate(Total_Requested_Num = as.numeric(Total_Requested))

for (i in 1:nrow(order_totals_3)) {
  if (is.na(order_totals_3$Total_Requested_Num[[i]])) {
    #print(order_totals_3$Item_Number[[i]])
    #print(order_totals_3$Description[[i]])
    #print(order_totals_3$Total_Requested_Num[[i]])
    
    print("4th Grade")    
    print(str_c(order_totals_3$Item_Number[[i]],
                order_totals_3$Description[[i]],
                order_totals_3$Total_Requested[[i]],
                sep = " "
    ))
  }
}

order_totals <-read.xlsx("5th Grade Order Totals.xlsx", sheetIndex = 1) %>% 
  rename(Item_Number = 2, Total_Requested = 5) %>% 
  select(Vendor, Item_Number, Description, Pkg, Total_Requested)

order_totals_2 <- order_totals[!is.na(order_totals$Total_Requested), ]

order_totals_3 <- order_totals_2 %>% 
  mutate(Total_Requested_Num = as.numeric(Total_Requested))

for (i in 1:nrow(order_totals_3)) {
  if (is.na(order_totals_3$Total_Requested_Num[[i]])) {
    #print(order_totals_3$Item_Number[[i]])
    #print(order_totals_3$Description[[i]])
    #print(order_totals_3$Total_Requested_Num[[i]])
    
    print("5th Grade")    
    print(str_c(order_totals_3$Item_Number[[i]],
                order_totals_3$Description[[i]],
                order_totals_3$Total_Requested[[i]],
                sep = " "
    ))
  }
}


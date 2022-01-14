####Combined report

#load libraries####
library(readxl)
library(googlesheets4)
library(lubridate)
#library(reshape2)
#library(plyr)
library(dplyr)
library(stringr)
library(data.table)



###Input data####

#commercial info

commercialInfo <- read_sheet("https://docs.google.com/spreadsheets/d/1JrT1sy5QGvSR6GaYiVEn2WrCS6y72ke0sJnYaw-VYAg",
                             range = "Updated Sellers List!A:K", col_names = T, na = "") %>%
                    select(c("Shop Name", "Mother Company", "KAM Name", "BDM Name",
                             "Category Head /Team Lead Name", "Category")) %>%
                    dplyr::rename(ShopName = 'Shop Name', MotherCompany = 'Mother Company',
                                  KAM = 'KAM Name', BDM = 'BDM Name',
                                  CatHead = "Category Head /Team Lead Name")
#commercialInfo <- commercialInfo[grepl("T10|Priority", ShopName, ignore.case = T), ]
commercialInfo <- data.table(commercialInfo)
commercialInfo$ShopName <- tolower(commercialInfo$ShopName)
commercialInfo$ShopName <- str_squish(commercialInfo$ShopName)


#seller payment

corporateTrackerUrl <- read.csv("Data/Input Data/PaymentTrackers/CorporateTrackers.csv",
                                stringsAsFactors = F, header = T, comment.char = "")

corporateTracker <- read_sheet(corporateTrackerUrl[1,3],
                               range = "Bill Submission Summary!A4:AR4", col_names = T, na = "") %>%
                    select(c("Bill Serial No", "Campaign Name", "Payment Date", "Submitted Bill", "Approved Amount",
                             "Amount", "Total Due for this bill")) %>%
                    dplyr::rename(Serial = 'Bill Serial No', ShopName = 'Campaign Name',
                                  PaymentDate = 'Payment Date', Submitted = 'Submitted Bill', Approved = 'Approved Amount',
                                  Paid = 'Amount', Due = 'Total Due for this bill')

for (i in 1:length(corporateTrackerUrl$URL)) {
                    temp <- read_sheet(corporateTrackerUrl[i,3],
                                       range = "Bill Submission Summary!A4:AR", col_names = T, na = "") %>%
                                        select(c("Bill Serial No", "Campaign Name", "Payment Date", "Submitted Bill", "Approved Amount",
                                                 "Amount", "Total Due for this bill")) %>%
                                        dplyr::rename(Serial = 'Bill Serial No', ShopName = 'Campaign Name',
                                                      PaymentDate = 'Payment Date', Submitted = 'Submitted Bill', Approved = 'Approved Amount',
                                                      Paid = 'Amount', Due = 'Total Due for this bill')
                    temp <- temp[!is.na(temp$Submitted), ]
                    corporateTracker <- rbind(corporateTracker, temp)}

corporateTracker$Serial <- as.character(corporateTracker$Serial)
corporateTracker$ShopName <- as.character(corporateTracker$ShopName)
corporateTracker$Submitted <- as.numeric(gsub(",","",corporateTracker$Submitted))
corporateTracker$Approved <- as.numeric(gsub(",","",corporateTracker$Approved))
corporateTracker$Paid <- as.numeric(gsub(",","",corporateTracker$Paid))
corporateTracker$Due <- as.numeric(gsub(",","",corporateTracker$Due))

corporateTracker <- data.table(corporateTracker)

#seller payment
CorporateSellerPayment <- corporateTracker[, .(CorporateSubmitted = sum(Submitted, na.rm = T),
                                               CorporateApproved = sum(Approved, na.rm = T),
                                               CorporatePaid = sum(Paid, na.rm = T),
                                               CorporateDue = sum(Due, na.rm = T)), by = .(ShopName)]

CorporateSellerPayment$ShopName <- tolower(CorporateSellerPayment$ShopName)
CorporateSellerPayment$ShopName <- str_squish(CorporateSellerPayment$ShopName)

rm(corporateTrackerUrl, corporateTracker, i, temp)


#sales data
file.list <- list.files(path = "Data/Input Data/SystemReport/SystemReport", pattern="*.xlsx", full.names = T)
sales <- lapply(file.list, read_excel) %>% bind_rows(.id = "id")

rm(file.list)

# add month column, remove comma from last update time and add last update date column,

sales <-  cbind(sales, "LastUpdatedate" = gsub(pattern = ",.*", "", sales$`Last Update`)) %>%
                    dplyr::select(c("Invoice No", "Order Date", "Shop Name", "Order Status", "Order Items",
                                    "Order Quantity", "Order Price", "Payment Method", "Note", "LastUpdatedate", "Payment Status")) %>%
                    dplyr::rename(Invoice = 'Invoice No', OrderDate = 'Order Date', ShopName = 'Shop Name',
                                  OrderStatus = 'Order Status', OrderItem = 'Order Items', PaymentMethod = 'Payment Method',
                                  OrderQuantity = 'Order Quantity', unit.OrderPrice = 'Order Price', LastUpdate = 'Note',
                                  PaymentStatus = 'Payment Status') %>%
                    dplyr::filter(PaymentStatus == "paid")

sales$unit.OrderPrice <- as.numeric(sales$unit.OrderPrice)
sales$OrderQuantity <- as.numeric(sales$OrderQuantity)

sales$OrderItem <- gsub(pattern= "\r|\n.*" , "", sales$OrderItem)
sales$OrderItem <- str_squish(sales$OrderItem)
sales$OrderItem <- tolower(sales$OrderItem)
sales$ShopName <- tolower(sales$ShopName)
sales$ShopName <- str_squish(sales$ShopName)

##common derivatives####
#unique invoice sales data
unqSales <- data.table(sales[!duplicated(sales$`Invoice`), ]) %>%
                    cbind(count = 1)

#unique shops in sales data
unqShop <- data.table(unqSales[!duplicated(unqSales$ShopName), ]) %>%
                    select(c("ShopName", "OrderDate")) %>%
                    rename(CampaignDate = OrderDate)

#add campaign date in the sales data
sales <- data.table(merge(sales, unqShop, all = T))


#Product data
#load sales data with product from Jan to Jul 16
file.list <- list.files(path = "Data/Input Data/SystemReport/SystemReportWithProd", pattern="*.xlsx", full.names = T)
salesProd <- lapply(file.list, read_excel) %>% bind_rows(.id = "id")

#sales and product data
salesProd <-  dplyr::select(salesProd, c("Shop Name", "Order Items", "Seller Price", '...25')) %>%
                    dplyr::rename(ShopName = 'Shop Name', OrderItem = 'Order Items',
                                  unit.mrp = 'Seller Price', unit.seller.price = '...25')

salesProd <- mutate(salesProd, code = paste(ShopName, OrderItem, sep = ""))
salesProd <- salesProd[!duplicated(salesProd$code), ] %>% select(-c("code"))

salesProd$OrderItem <- gsub(pattern= "\r|\n.*" , "", salesProd$OrderItem)
salesProd$OrderItem <- str_squish(salesProd$OrderItem)
salesProd$OrderItem <- tolower(salesProd$OrderItem)
salesProd$ShopName <- tolower(salesProd$ShopName)
salesProd$ShopName <- str_squish(salesProd$ShopName)

salesProd$unit.OrderPrice <- as.numeric(salesProd$unit.OrderPrice)
salesProd$OrderQuantity <- as.numeric(salesProd$OrderQuantity)
salesProd$unit.mrp <- as.numeric(salesProd$unit.mrp)
salesProd$unit.seller.price <- as.numeric(salesProd$unit.seller.price)

##Product data
prodFiles <- list.files("Data/Input Data/ProductData/New Product Data", pattern="*.csv", full.names = T)

productData <- lapply(prodFiles, fread, sep = "\t", header = TRUE, data.table = TRUE) %>%
                    bind_rows(.id = "id") %>%
                    select(c('shop name', 'order item', 'seller price', 'mrp')) %>%
                    rename(ShopName = 'shop name', OrderItem = 'order item',
                           unit.seller.price = 'seller price', unit.mrp = 'mrp')

rm(file.list, prodFiles)

#the 3 operation takes about 21 seconds in i9
productData$OrderItem <- str_squish(productData$OrderItem)
productData$OrderItem <- tolower(productData$OrderItem)
productData$ShopName <- str_squish(productData$ShopName)
productData$ShopName <- tolower(productData$ShopName)

productData <- rbind(productData, salesProd)
productData <- mutate(productData, code = paste(ShopName, OrderItem, sep = ""))
productData <- productData[!duplicated(productData$code), ] %>% select(-c("code"))

evalyPrice <- evalyPrice[!grepl('for t10 july 30', ShopName) & !grepl('for t10 july 23', ShopName), ]

evalyPrice <- merge(sales, productData, all.x = T, all.y = F, by = c("ShopName", "OrderItem")) %>%
                    group_by(ShopName, OrderItem) %>%
                    mutate(itemCount = row_number()) %>%
                    filter(itemCount == 1) %>%
                    select(c("ShopName", "OrderItem", "unit.OrderPrice", "unit.seller.price",
                             "unit.mrp"))

tempProd <- read.csv("Data/Input Data/ProductData/T102330.csv") %>%
                    rename(ShopName = 'ï..ShopName')

evalyPrice <- rbind(evalyPrice, tempProd)

evalyPrice$unit.OrderPrice <- as.numeric(evalyPrice$unit.OrderPrice)
evalyPrice$unit.seller.price <- as.numeric(evalyPrice$unit.seller.price)
evalyPrice$unit.mrp <- as.numeric(evalyPrice$unit.mrp)

evalyPrice <- filter(evalyPrice, unit.mrp > 0)
evalyPrice <- filter(evalyPrice, unit.seller.price > 0)
evalyPrice <- filter(evalyPrice, unit.OrderPrice > 0)
evalyPrice <- filter(evalyPrice, unit.mrp >= unit.seller.price)

#shop-item merge - shop specific info (campaign date), item status, price info, commercial info, delivery time info

totalInv <- unqSales[OrderStatus != "cancel", .(Total.Invoice = sum(count, na.rm = T)), by = .(ShopName, OrderItem)]
processingInv <- unqSales[OrderStatus == "processing",
                          .(Processing.Invoice = sum(count, na.rm = T)), by = .(ShopName, OrderItem)]
pickedInv <- unqSales[OrderStatus == "picked",
                      .(Picked.Invoice = sum(count, na.rm = T)), by = .(ShopName, OrderItem)]
refundedInv <- unqSales[grepl("refund", LastUpdate, ignore.case = T) & (OrderStatus == "shipped" | OrderStatus == "delivered"),
                        .(Refunded.Invoice = sum(count, na.rm = T)), by = .(ShopName, OrderItem)]
deliveredInv <- unqSales[!grepl("refund", LastUpdate, ignore.case = T) & (OrderStatus == "shipped" | OrderStatus == "delivered"),
                         .(Delivered.And.Shipped.Inv = sum(count, na.rm = T)), by = .(ShopName, OrderItem)]
undeliveredInv <- unqSales[OrderStatus == "picked" | OrderStatus == "processing", .(Undelivered.Invoice = sum(count, na.rm = T)),
                           by = .(ShopName, OrderItem)]

InvoiceStatus <- merge(merge(merge(merge(merge(totalInv, processingInv, all = T),
                                         pickedInv, all = T),
                                   deliveredInv, all = T),
                             refundedInv, all = T),
                       undeliveredInv, all = T) %>%
                    mutate(refundRatio = Refunded.Invoice/Total.Invoice)


#Order status, month, shop, item wise item count. All can take about 20 minutes.
total <- sales[OrderStatus != "cancel", .(Total.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
processing <- sales[OrderStatus == "processing", .(Processing.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
picked <- sales[OrderStatus == "picked", .(Picked.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
refunded <- sales[grepl("refund", LastUpdate, ignore.case = T) & (OrderStatus == "shipped" | OrderStatus == "delivered"), .(Refunded.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
delivered <- sales[!grepl("refund", LastUpdate, ignore.case = T) & (OrderStatus == "shipped" | OrderStatus == "delivered"),
                   .(Delivered.And.Shipped.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
undelivered <- sales[OrderStatus == "picked" | OrderStatus == "processing", .(Undelivered.Quantity = sum(OrderQuantity, na.rm = T)),
                     by = .(ShopName, OrderItem)]

ItemStatus <- merge(merge(merge(merge(merge(total, processing, all = T),
                                      picked, all = T),
                                delivered, all = T),
                          refunded, all = T),
                    undelivered, all = T)

#last date difference
shopDelivery <- dplyr::filter(sales, OrderStatus == 'shipped' | OrderStatus == 'delivered') %>%
                    dplyr::filter(!grepl("refund", LastUpdate, ignore.case = T)) %>%
                    mutate(deliDays = dmy(LastUpdatedate) - mdy(CampaignDate))
shopDelivery <- shopDelivery[, .(shopItemDeliTime = sum(deliDays, na.rm = T)), by = .(ShopName, OrderItem)]

shopDelivery$shopItemDeliTime <- as.numeric(shopDelivery$shopItemDeliTime, units="days")

#Nagad impact: prior to 16/7 10% up to 2000
nagad <- merge((select(sales, -c("unit.OrderPrice"))), evalyPrice, all = F, by = c("ShopName", "OrderItem"))
nagad <- nagad[grepl("nagad", PaymentMethod, ignore.case = T),
               .(NagadDiscountValue =  sum(ifelse((OrderQuantity*unit.OrderPrice*0.1)>2000,
                                                  2000, (OrderQuantity * unit.OrderPrice * .1)),
                                           na.rm = T)),
               by = .(ShopName, OrderItem)]

shopItemReport <- merge(merge(merge(merge(merge(merge(merge(unqShop, CorporateSellerPayment, all.x = T, all.y = F, by = c("ShopName")),
                                                      ItemStatus, all = F, by = c("ShopName")),
                                                evalyPrice, all = F, by = c("ShopName", "OrderItem")),
                                          commercialInfo, all.x = T, all.y = F, by = c("ShopName")),
                                    shopDelivery, all.x = T, all.y = F, by = c("ShopName", "OrderItem")),
                              InvoiceStatus, all = F, by = c("ShopName", "OrderItem")),
                        nagad, all.x = T, all.y = F, by = c("ShopName", "OrderItem")) %>%
                    mutate(TotalOrderValue = Total.Quantity * unit.OrderPrice,
                           ProcessingValue = Processing.Quantity * unit.OrderPrice,
                           PickedValue = Picked.Quantity * unit.OrderPrice,
                           Delivered.And.Shipped.Value = Delivered.And.Shipped.Quantity * unit.OrderPrice,
                           Undelivered.Value = Undelivered.Quantity * unit.OrderPrice,
                           refundValue = Refunded.Quantity*unit.seller.price,
                           
                           TotalMarketValue = Total.Quantity * unit.mrp,
                           
                           TotalSellerPrice = Total.Quantity * unit.seller.price,
                           ProcessingSellerPrice = Processing.Quantity * unit.seller.price,
                           PickedSellerPrice = Picked.Quantity * unit.seller.price,
                           Delivered.And.Shipped.SellerPrice = Delivered.And.Shipped.Quantity * unit.seller.price,
                           Undelivered.SellerPrice = Undelivered.Quantity * unit.seller.price,
                           
                           comissionValue = (unit.mrp-unit.seller.price)*Total.Quantity, 
                           discountValue = (unit.mrp - unit.OrderPrice)*Total.Quantity,
                           burnValue = (unit.seller.price - unit.OrderPrice)*Total.Quantity) %>%
                    cbind(count = 1)

shopItemReport <- data.table(shopItemReport)
cost <- shopItemReport[, .(Cost = sum(Delivered.And.Shipped.SellerPrice, Undelivered.SellerPrice, refundValue, na.rm = T)),
                       by = .(ShopName, OrderItem)]

shopItemReport <- merge(shopItemReport, cost, by = c("ShopName", "OrderItem")) %>%
                    mutate(Month = month(mdy(CampaignDate)),
                           Week = week(mdy(CampaignDate)),
                           
                           comissionRate = (TotalMarketValue - TotalSellerPrice)/TotalMarketValue,
                           burnRate = (TotalSellerPrice - (TotalOrderValue-NagadDiscountValue))/TotalMarketValue,
                           discountRate = (TotalMarketValue - (TotalOrderValue-NagadDiscountValue))/TotalMarketValue,
                           
                           costRatio = Cost/TotalOrderValue,
                           avgDeliTime = shopItemDeliTime/Delivered.And.Shipped.Inv)

shopItemReport <- shopItemReport[!is.na(shopItemReport$ShopName), ]
shopItemReport <- shopItemReport[!is.na(shopItemReport$Total.Quantity), ]
shopItemReport <- shopItemReport[!is.na(shopItemReport$Total.Invoice), ]

shopItemReport <- shopItemReport[, c("MotherCompany", "KAM", "BDM", "Category", "CatHead",
                                     "ShopName", "CampaignDate", "Month", "Week",
                                     
                                     "OrderItem", "unit.OrderPrice", "unit.seller.price", "unit.mrp",
                                     
                                     "Total.Invoice", "Processing.Invoice", "Picked.Invoice", "Delivered.And.Shipped.Inv",
                                     "Refunded.Invoice", "Undelivered.Invoice", "refundRatio",
                                     
                                     "Total.Quantity", "Processing.Quantity", "Picked.Quantity", "Delivered.And.Shipped.Quantity",
                                     "Refunded.Quantity", "Undelivered.Quantity",
                                     
                                     "TotalMarketValue",
                                     
                                     "TotalSellerPrice", "ProcessingSellerPrice", "PickedSellerPrice", "Delivered.And.Shipped.SellerPrice", "Undelivered.SellerPrice",
                                     
                                     "TotalOrderValue", "ProcessingValue", "PickedValue", "Delivered.And.Shipped.Value", "Undelivered.Value", "refundValue",
                                     
                                     "comissionRate", "burnRate", "discountRate", "comissionValue", "discountValue", "burnValue",
                                     
                                     "NagadDiscountValue", "Cost", "costRatio", "avgDeliTime", "shopItemDeliTime", "count",
                                     
                                     "CorporateSubmitted", "CorporateApproved", "CorporatePaid", "CorporateDue")]

write.csv(shopItemReport, "Data/Output Data/CommercialReport/ShopItemReport.csv")

rm(total, processing, picked, delivered, refunded, undelivered, cost, totalInv, processingInv,
   pickedInv, deliveredInv, refundedInv, undeliveredInv, nagad)


##output variations####
#read data
shopItemReport <- fread("Data/Output Data/CommercialReport/ShopItemReport.csv")


#T10 campaign comparison
campaign1 <- shopItemReport[grepl('for t10 july 2', ShopName) & !grepl('for t10 july 23', ShopName), ] %>%
                     cbind(CampaignName = c("T10 July 2"))
campaign2 <- shopItemReport[grepl('for t10 july 9 2021', ShopName), ] %>%
                     cbind(CampaignName = c("T10 July 9"))
campaign3 <- shopItemReport[grepl('for t10 july 16', ShopName), ] %>%
                     cbind(CampaignName = c("T10 July 16"))
campaign4 <- shopItemReport[grepl('for t10 july 23', ShopName), ] %>%
                     cbind(CampaignName = c("T10 July 23"))
campaign5 <- shopItemReport[grepl('for t10 july 30', ShopName), ] %>%
                     cbind(CampaignName = c("T10 July 30"))

shopItemReport <- merge(merge(merge(merge(campaign1, campaign2, all = T),
                                     campaign3, all = T),
                               campaign4, all = T),
                         campaign5, all = T)

#priority
#shopItemReport <- shopItemReport[grepl('for priority', ShopName), ]

#management report

TotalValue <- shopItemReport[,.(TotalValue = sum(TotalOrderValue, na.rm = T)),
                             #by = .(Category, BDM, KAM, MotherCompany)] # monthless mother company
                             #by = .(Category, BDM, KAM, MotherCompany, Month)] # monthly mother company
                             #by = .(Category, BDM, KAM, MotherCompany, Week)] # weekly mother company
                             #by = .(Category, BDM, KAM, MotherCompany, CampaignName)] # Campaign MC Comparison
                             by = .(CampaignName)] # Campaign Comparison
                             #by = .(Category, BDM, KAM, Month)] # monthly kam report
                             #by = .(Category, BDM, Month)] #monthly BDM
                             #by = .(Category)] #overall category
                             #by = .(Category, Month)] #monthly category
                             #by = .(Category, Week)] #weekly category
                             #by = .(Month)] #overall monthly
                             #] #overall company

shopItemReport <- merge(shopItemReport, TotalValue, all = T,
                        #by = c("Category", "BDM", "KAM", "MotherCompany")) # monthless mother company
                        #by = c("Category", "BDM", "KAM", "MotherCompany", "Month")) # monthly mother company
                        #by = c("Category", "BDM", "KAM", "MotherCompany", "Week")) # weekly mother company
                        #by = c("Category", "BDM", "KAM", "MotherCompany", "CampaignName")) #campaign MC comparison
                        by = c("CampaignName")) #campaign comparison
                        #by = c("Category", "BDM", "KAM", "Month")) # monthly kam report
                        #by = c("Category", "BDM", "Month")) #monthly BDM
                        #by = c("Category")) #overall category
                        #by = c("Category", "Month")) #monthly category
                        #by = c("Category", "Week")) #weekly category
                        #by = c("Month")) #overall monthly

#shopItemReport <- cbind(shopItemReport, TotalValue) #for overall company

rm(TotalValue, campaign1, campaign2, campaign3, campaign4, campaign5)

#Report
Report <- shopItemReport[, .(TotalInvoice = sum(Total.Invoice, na.rm = T),
                             ProcessingInvoice = sum(Processing.Invoice, na.rm = T),
                             PickedInvoice = sum(Picked.Invoice, na.rm = T),
                             DeliveredAndShippedInvoice = sum(Delivered.And.Shipped.Inv, na.rm = T),
                             RefundedInvoice = sum(Refunded.Invoice, na.rm = T),
                             
                             Submitted = sum(CorporateSubmitted, na.rm = T),
                             Approved = sum(CorporateApproved, na.rm = T),
                             Paid = sum(CorporatePaid, na.rm = T),
                             Due = sum(CorporateDue, na.rm = T),
                             
                             Revenue = sum(TotalOrderValue, na.rm = T),
                             Cost = sum(Cost, na.rm = T),
                             
                             Delivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T),
                             UndeliveredValue = sum(Undelivered.Value, na.rm = T),
                             
                             Delivered.And.Shipped.SellerPrice = sum(Delivered.And.Shipped.SellerPrice, na.rm = T),
                             Undelivered.SellerPrice = sum(Undelivered.SellerPrice, na.rm = T),
                             
                             shopItemDeliTime = sum(shopItemDeliTime, na.rm = T),
                             
                             TotalMarketValue = sum(TotalMarketValue, na.rm = T),
                             TotalSellerPrice = sum(TotalSellerPrice, na.rm = T),
                             NagadDiscountValue = sum(NagadDiscountValue, na.rm = T)),
                         
                         #by = .(Category, BDM, KAM, MotherCompany)] # monthless mother company
                         #by = .(Category, BDM, KAM, MotherCompany, Month)] # monthly mother company
                         #by = .(Category, BDM, KAM, MotherCompany, Week)] # weekly mother company
                         #by = .(Category, BDM, KAM, MotherCompany, CampaignName)] # Campaign MC Comparison
                         by = .(CampaignName)] # Campaign Comparison
                         #by = .(Category, BDM, KAM, Month)] # monthly kam report
                         #by = .(Category, BDM, Month)] #monthly BDM
                         #by = .(Category)] #overall category
                         #by = .(Category, Month)] #monthly category
                         #by = .(Category, Week)] #weekly category
                         #by = .(Month)] #overall monthly
                         #] #overall company


Report <- mutate(Report,
                 creditPerformance = Delivered.And.Shipped.SellerPrice/Paid,
                 avgRefundRatio = RefundedInvoice/TotalInvoice,
                 avgDeliTime = shopItemDeliTime/DeliveredAndShippedInvoice,
                 avgCostRatio = Cost/Revenue,
                 
                 avgComissionRate = ((TotalMarketValue - TotalSellerPrice)/TotalMarketValue),
                 avgBurnRate = ((TotalSellerPrice - (Revenue-NagadDiscountValue))/TotalMarketValue),
                 avgDiscountRate = ((TotalMarketValue - (Revenue-NagadDiscountValue))/TotalMarketValue))


Report <- Report[, c(#"Category", "BDM", "KAM", "MotherCompany", "Month", "Week", "CampaignName",
                    "CampaignName",
                    "TotalInvoice", "ProcessingInvoice", "PickedInvoice", "DeliveredAndShippedInvoice" , "RefundedInvoice",
                    
                    "Revenue", "Delivered.And.Shipped.Value", "UndeliveredValue", "Cost",
                    
                    "Delivered.And.Shipped.SellerPrice", "Undelivered.SellerPrice",
                    
                    "TotalMarketValue", "TotalSellerPrice", "NagadDiscountValue",
                    
                    "avgComissionRate", "avgBurnRate", "avgDiscountRate",
                    
                    "creditPerformance", "avgRefundRatio", "avgDeliTime", "avgCostRatio",
                    
                    "Submitted", "Approved", "Paid", "Due")]

write.csv(Report, "Data/Output Data/CommercialReport/T10.csv")#WhatReport.csv")



##Operational reports####

# commercialReportUrl <- read.csv("Data/Input Data/CommercialReport/CommercialReportUrl.csv", header = T, comment.char = "")
# invoice = NA
# item = NA
# 
# for (i in 1:length(commercialReportUrl$URL)){
#                     item = filter(shopItemReport, BDM == commercialReportUrl[i, 1])
#                     
#                     sheet_write(item, ss = commercialReportUrl[i, 2], sheet = "ItemData")
#                     sheet_write(as.data.frame(Sys.Date()), ss = commercialReportUrl[i, 2], sheet = "Update Date")
# }
# 
# rm(i, invoice, item)

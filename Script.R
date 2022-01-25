#load libraries####
library(readxl)
library(googlesheets4)
library(lubridate)
library(reshape2)
library(plyr)
library(dplyr)
library(stringr)
library(data.table)


#Part 1: IMPORT####


## Commercial info####

commercialInfo <- read_sheet("https://docs.google.com/spreadsheets/d/1JrT1sy5QGvSR6GaYiVEn2WrCS6y72ke0sJnYaw-VYAg",
                             range = "Updated Sellers List!A:K", col_names = T, na = "") %>%
     select(c("Shop Name", "Mother Company", "KAM Name", "BDM Name",
              "Category Head /Team Lead Name", "Category")) %>%
     dplyr::rename(ShopName = 'Shop Name', MotherCompany = 'Mother Company',
                   KAM = 'KAM Name', BDM = 'BDM Name',
                   CatHead = "Category Head /Team Lead Name")

commercialInfo <- data.table(commercialInfo)
#commercialInfo <- commercialInfo[grepl("T10|Priority", ShopName, ignore.case = T), ]
commercialInfo$ShopName <- tolower(commercialInfo$ShopName)
commercialInfo$ShopName <- str_squish(commercialInfo$ShopName)

##Product data####
prodFiles <- list.files("Data/Input Data/ProductData/New Product Data", pattern="*.csv", full.names = T)

productData <- lapply(prodFiles, fread, sep = "\t", header = TRUE, data.table = TRUE) %>%
     bind_rows(.id = "id") %>%
     select(-c("id", "product id", "product specifications",
               'enlisted shop code', "e price", "category name")) %>%
     rename(OrderItem = 'order item', ShopName = 'shop name', Brand = 'brand name',
            unit.seller.price = 'seller price', unit.mrp = 'mrp')

#the 3 operation takes about 21 seconds in i9
productData$OrderItem <- str_squish(productData$OrderItem)
productData$OrderItem <- tolower(productData$OrderItem)
productData$ShopName <- str_squish(productData$ShopName)
productData$ShopName <- tolower(productData$ShopName)

productData <- filter(productData, productData$unit.mrp > 0)
productData <- filter(productData, productData$unit.seller.price > 0)
productData <- filter(productData, productData$unit.mrp > productData$unit.seller.price)

##System data####

file.list <- list.files(path = "Data/Input Data/SystemReport/System Report Till May", pattern="*.xlsx", full.names = T)

#load the excel files, rename the last column as shopCode. Takes about 5 minutes for data since November.
sales <- lapply(file.list, read_excel) %>%
     bind_rows(.id = "id")

rm(file.list, prodFiles)

# add month column, remove comma from last update time and add last update date column,

sales <-  mutate(sales, Month = month(mdy(sales$`Order Date`)))%>%
     cbind("LastUpdatedate" = gsub(pattern = ",.*", "", sales$`Last Update`)) %>%
     dplyr::select(c("Invoice No", "Order Date", "Shop Name", "Order Status", "Order Items",
                     "Order Quantity", "Order Price", "Note", "LastUpdatedate", "Payment Status",
                     "Month")) %>%
     dplyr::rename(Invoice = 'Invoice No', OrderDate = 'Order Date', ShopName = 'Shop Name',
                   OrderStatus = 'Order Status', OrderItem = 'Order Items',
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

#unique invoice sales data
unqSales <- data.frame(sales[!duplicated(sales$`Invoice`), ]) %>%
     cbind(count = 1)

#unique shops in sales data
unqShop <- data.frame(unqSales[!duplicated(unqSales$ShopName), ]) %>%
     select(c("ShopName", "OrderDate")) %>%
     rename(CampaignDate = OrderDate)

#add campaign date in the sales data
sales <- merge(sales, unqShop, all = T)

sales <- data.table(sales)
unqSales <- data.table(unqSales)

##seller payment####

corporateTrackerUrl <- read.csv("Data/Input Data/PaymentTrackers/CorporateTrackers.csv",
                                stringsAsFactors = F, header = T, comment.char = "")


corporateTracker <- read_sheet(corporateTrackerUrl[1,3],
                               range = "Bill Submission Summary!A4:AR4", col_names = T, na = "") %>%
     select(c("Bill Serial No", "Campaign Name", "Payment Date", "Submitted Bill", "Approved Amount",
              "Amount", "Total Due for this bill")) %>%
     dplyr::rename(Serial = 'Bill Serial No', ShopName = 'Campaign Name',
                   PaymentDate = 'Payment Date', Submitted = 'Submitted Bill', Approved = 'Approved Amount',
                   Paid = 'Amount', Due = 'Total Due for this bill')


system.time(
     for (i in 1:length(corporateTrackerUrl$URL)) {
          temp <- read_sheet(corporateTrackerUrl[i,3],
                             range = "Bill Submission Summary!A4:AR", col_names = T, na = "") %>%
               select(c("Bill Serial No", "Campaign Name", "Payment Date", "Submitted Bill", "Approved Amount",
                        "Amount", "Total Due for this bill")) %>%
               dplyr::rename(Serial = 'Bill Serial No', ShopName = 'Campaign Name',
                             PaymentDate = 'Payment Date', Submitted = 'Submitted Bill', Approved = 'Approved Amount',
                             Paid = 'Amount', Due = 'Total Due for this bill')
          temp <- temp[!is.na(temp$Submitted), ]
          corporateTracker <- rbind(corporateTracker, temp)
     })

corporateTracker$Serial <- as.character(corporateTracker$Serial)
corporateTracker$ShopName <- as.character(corporateTracker$ShopName)
corporateTracker$Submitted <- as.numeric(gsub(",","",corporateTracker$Submitted))
corporateTracker$Approved <- as.numeric(gsub(",","",corporateTracker$Approved))
corporateTracker$Paid <- as.numeric(gsub(",","",corporateTracker$Paid))
corporateTracker$Due <- as.numeric(gsub(",","",corporateTracker$Due))

#Add shop specific data with corporate tracker data
corporateTracker <- merge(corporateTracker, unqShop, all.x = T, all.y = F)
corporateTracker <- data.table(corporateTracker)

#seller payment

CorporateSellerPayment <- corporateTracker[, .(CorporateSubmitted = sum(Submitted, na.rm = T),
                                               CorporateApproved = sum(Approved, na.rm = T),
                                               CorporatePaid = sum(Paid, na.rm = T),
                                               CorporateDue = sum(Due, na.rm = T)), by = .(ShopName)]

CorporateSellerPayment$ShopName <- tolower(CorporateSellerPayment$ShopName)
CorporateSellerPayment$ShopName <- str_squish(CorporateSellerPayment$ShopName)

rm(corporateTrackerUrl, corporateTracker, i, temp)


#Part 2: PROCESSING####


##Invoice Status####

totalInv <- unqSales[, .(Total.Invoice = sum(count, na.rm = T)), by = .(ShopName)]
processingInv <- unqSales[OrderStatus == "processing",
                          .(Processing.Invoice = sum(count, na.rm = T)), by = .(ShopName)]
pickedInv <- unqSales[OrderStatus == "picked",
                      .(Picked.Invoice = sum(count, na.rm = T)), by = .(ShopName)]
refundedInv <- unqSales[grepl("refund", LastUpdate, ignore.case = T),
                        .(Refunded.Invoice = sum(count, na.rm = T)), by = .(ShopName)]
deliveredInv <- unqSales[!grepl("refund", LastUpdate, ignore.case = T) & (OrderStatus == "shipped" | OrderStatus == "delivered"),
                         .(Delivered.And.Shipped.Inv = sum(count, na.rm = T)), by = .(ShopName)]
undeliveredInv <- unqSales[OrderStatus == "picked" | OrderStatus == "processing", .(Undelivered.Invoice = sum(count, na.rm = T)),
                           by = .(ShopName)]

InvoiceStatus <- merge(merge(merge(merge(merge(totalInv, processingInv, all = T),
                                         pickedInv, all = T),
                                   deliveredInv, all = T),
                             refundedInv, all = T),
                       undeliveredInv, all = T) %>%
     mutate(refundRatio = Refunded.Invoice/Total.Invoice)


#merge InvoiceStatus report with commercial info
shopInvoiceReport <- merge(merge(merge(unqShop, InvoiceStatus, all = T, by = c("ShopName")),
                                 commercialInfo, all.x = T, all.y = F, by = c("ShopName")),
                           CorporateSellerPayment, all.x = T, all.y = F, by = c("ShopName")) %>%
     mutate(Month = month(mdy(CampaignDate)))

shopInvoiceReport <- shopInvoiceReport[!is.na(shopInvoiceReport$ShopName), ]


#export shop report
write.csv(shopInvoiceReport, "Data/Output Data/CommercialReport/shopInvoiceReport.csv")

#stop auto running here

rm(totalInv, processingInv, pickedInv, deliveredInv, refundedInv, undeliveredInv)


##shop-item report####

#Order status, month, shop, item wise item count. All can take about 20 minutes.
total <- sales[, .(Total.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
processing <- sales[OrderStatus == "processing", .(Processing.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
picked <- sales[OrderStatus == "picked", .(Picked.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
refunded <- sales[grepl("refund", LastUpdate, ignore.case = T), .(Refunded.Quantity = sum(OrderQuantity, na.rm = T)), by = .(ShopName, OrderItem)]
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
deliveredInv <- unqSales[!grepl("refund", LastUpdate, ignore.case = T) & (OrderStatus == "shipped" | OrderStatus == "delivered"),
                         .(Delivered.And.Shipped.Inv = sum(count, na.rm = T)), by = .(ShopName, OrderItem)]

shopDeliTime <- dplyr::filter(sales, OrderStatus == 'shipped' | OrderStatus == 'delivered') %>%
     dplyr::filter(!grepl("refund", LastUpdate, ignore.case = T)) %>%
     mutate(deliDays = dmy(LastUpdatedate) - mdy(CampaignDate))
shopDeliTime <- shopDeliTime[, .(shopItemDeliTime = sum(deliDays, na.rm = T)), by = .(ShopName, OrderItem)]


#avg delivery time of shops
shopDelivery <- merge(deliveredInv, shopDeliTime, all = T) %>%
     mutate(avgDeliTime = shopItemDeliTime/Delivered.And.Shipped.Inv) %>%
     select(-c(shopItemDeliTime, Delivered.And.Shipped.Inv))

shopDelivery$avgDeliTime <- as.numeric(shopDelivery$avgDeliTime, units="days")


#item e price
ePrice <- sales %>% group_by(ShopName, OrderItem) %>%
     mutate(itemCount = row_number()) %>%
     filter(itemCount == 1) %>%
     select(c("ShopName", "OrderItem", "unit.OrderPrice"))
ePrice$unit.OrderPrice <- as.numeric(ePrice$unit.OrderPrice)


shopItemReport <- merge(merge(merge(merge(merge(unqShop, ItemStatus, all = T),
                                          productData, all.x = T, all.y = F),
                                    ePrice, all.x = T, all.y = F),
                              commercialInfo, all.x = T, all.y = F),
                        shopDelivery, all.x = T, all.y = F)%>%
     mutate(TotalSellerPrice = Total.Quantity * unit.seller.price, ProcessingSellerPrice = Processing.Quantity * unit.seller.price,
            PickedSellerPrice = Picked.Quantity * unit.seller.price, Delivered.And.Shipped.SellerPrice = Delivered.And.Shipped.Quantity * unit.seller.price,
            Undelivered.SellerPrice = Undelivered.Quantity * unit.seller.price,
            TotalOrderValue = Total.Quantity * unit.OrderPrice, ProcessingValue = Processing.Quantity * unit.OrderPrice,
            PickedValue = Picked.Quantity * unit.OrderPrice, Delivered.And.Shipped.Value = Delivered.And.Shipped.Quantity * unit.OrderPrice,
            Undelivered.Value = Undelivered.Quantity * unit.OrderPrice,
            refundValue = Refunded.Quantity*unit.seller.price,
            comission = (unit.mrp - unit.seller.price)/unit.mrp) %>%
     cbind(count = 1)

shopItemReport <- data.table(shopItemReport)
cost <- shopItemReport[, .(cost = sum(Delivered.And.Shipped.SellerPrice, Undelivered.SellerPrice, refundValue, na.rm = T)), by = .(ShopName, OrderItem)]

shopItemReport <- merge(shopItemReport, cost) %>%
     mutate(Month = month(mdy(CampaignDate))) %>%
     mutate(costRatio = cost/TotalOrderValue)

write.csv(shopItemReport, "Data/Output Data/CommercialReport/ShopItemReport.csv")

rm(total, processing, picked, delivered, refunded, undelivered, shopDeliTime, shopDelivery, ePrice, cost, deliveredInv)


#Part 3: Export####


shopInvoiceReport <- read.csv("Data/Output Data/CommercialReport/shopInvoiceReport.csv",
                              stringsAsFactors = F, header = T, comment.char = "")
shopItemReport <- read.csv("Data/Output Data/CommercialReport/shopItemReport.csv",
                           stringsAsFactors = F, header = T, comment.char = "")

shopInvoiceReport <- data.table(shopInvoiceReport)
shopItemReport <- data.table(shopItemReport)


shopInvoiceReport <- shopInvoiceReport[, c("Category", "CatHead", "BDM", "KAM",
                                           "MotherCompany", "ShopName", "CampaignDate", "Month",
                                           "Total.Invoice", "Processing.Invoice", "Picked.Invoice", "Delivered.And.Shipped.Inv",
                                           "Refunded.Invoice", "Undelivered.Invoice", "refundRatio",
                                           "CorporateSubmitted", "CorporateApproved", "CorporatePaid", "CorporateDue")]

shopItemReport <- shopItemReport[, c("MotherCompany", "KAM", "BDM", "Category", "CatHead",
                                     "ShopName", "CampaignDate", "Month",
                                     "OrderItem", "unit.OrderPrice", "unit.seller.price", "unit.mrp",
                                     "Total.Quantity", "Processing.Quantity", "Picked.Quantity", "Delivered.And.Shipped.Quantity",
                                     "Refunded.Quantity", "Undelivered.Quantity",
                                     "TotalSellerPrice", "ProcessingSellerPrice", "PickedSellerPrice", "Delivered.And.Shipped.SellerPrice", "Undelivered.SellerPrice",
                                     "TotalOrderValue", "ProcessingValue", "PickedValue", "Delivered.And.Shipped.Value", "Undelivered.Value", "refundValue",
                                     "comission", "cost", "costRatio", "avgDeliTime", "count")]


##BDM/cat head teams####

commercialReportUrl <- read.csv("Data/Input Data/CommercialReport/CommercialReportUrl.csv", header = T, comment.char = "")
invoice = NA
item = NA

for (i in 1:length(commercialReportUrl$URL)){
     invoice = filter(shopInvoiceReport, BDM == commercialReportUrl[i, 1])
     item = filter(shopItemReport, BDM == commercialReportUrl[i, 1])
     
     sheet_write(invoice, ss = commercialReportUrl[i, 2], sheet = "InvoiceData")
     sheet_write(item, ss = commercialReportUrl[i, 2], sheet = "ItemData")
     sheet_write(as.data.frame(Sys.Date()), ss = commercialReportUrl[i, 2], sheet = "Update Date")
}

rm(i, invoice, item)

##Management report####

MCTotalValue <- shopItemReport[,.(MCTotalValue = sum(TotalOrderValue, na.rm = T)), by = .(CatHead, BDM, KAM, MotherCompany, Month)]
KAMTotalValue <- shopItemReport[,.(KAMTotalValue = sum(TotalOrderValue, na.rm = T)), by = .(CatHead, BDM, KAM, Month)]
BDMTotalValue <- shopItemReport[,.(BDMTotalValue = sum(TotalOrderValue, na.rm = T)), by = .(CatHead, BDM, Month)]
CatTotalValue <- shopItemReport[,.(CatHeadTotalValue = sum(TotalOrderValue, na.rm = T)), by = .(CatHead, Month)]
MCSDValue <- shopItemReport[,.(MCDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T)), by = .(CatHead, BDM, KAM, MotherCompany, Month)]
KAMSDValue <- shopItemReport[,.(KAMDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T)), by = .(CatHead, BDM, KAM, Month)]
BDMSDValue <- shopItemReport[,.(BDMDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T)), by = .(CatHead, BDM, Month)]
CatSDValue <- shopItemReport[,.(CatHeadDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T)), by = .(CatHead, Month)]

shopItemReport <- merge(merge(merge(merge(merge(merge(merge(merge(shopItemReport, MCTotalValue, all = T, by = c("CatHead", "BDM", "KAM", "MotherCompany", "Month")),
                                                            KAMTotalValue, all = T, by = c("CatHead", "BDM", "KAM", "Month")),
                                                      BDMTotalValue, all = T, by = c("CatHead", "BDM", "Month")),
                                                CatTotalValue, all = T, by = c("CatHead", "Month")),
                                          MCSDValue, all = T, by = c("CatHead", "BDM", "KAM", "MotherCompany", "Month")),
                                    KAMSDValue, all = T, by = c("CatHead", "BDM", "KAM", "Month")),
                              BDMSDValue, all = T, by = c("CatHead", "BDM", "Month")),
                        CatSDValue, all = T, by = c("CatHead", "Month"))

rm(MCTotalValue, KAMTotalValue, BDMTotalValue, CatTotalValue, MCSDValue, KAMSDValue, BDMSDValue, CatSDValue)


###Mother company report####

MCInvoiceReport <- shopInvoiceReport[ , .(MCTotalInvoice = sum(Total.Invoice, na.rm = T),
                                          MCProcessingInvoice = sum(Processing.Invoice, na.rm = T),
                                          MCPickedInvoice = sum(Picked.Invoice, na.rm = T),
                                          MCDeliveredAndShippedInvoice = sum(Delivered.And.Shipped.Inv, na.rm = T),
                                          MCRefundedInvoice = sum(Refunded.Invoice, na.rm = T),
                                          MCSubmitted = sum(CorporateSubmitted, na.rm = T),
                                          MCApproved = sum(CorporateApproved, na.rm = T),
                                          MCPaid = sum(CorporatePaid, na.rm = T),
                                          MCDue = sum(CorporateDue, na.rm = T)),
                                      by = .(CatHead, BDM, KAM, MotherCompany, Month)] %>%
     mutate(MCavgRefundRatio = MCRefundedInvoice/MCTotalInvoice)


MCReport <- shopItemReport[, .(MCRevenue = sum(TotalOrderValue, na.rm = T),
                               MCDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T),
                               MCUndeliveredValue = sum(Undelivered.Value, na.rm = T),
                               MCDelivered.And.Shipped.SellerPrice = sum(Delivered.And.Shipped.SellerPrice, na.rm = T),
                               MCUndelivered.SellerPrice = sum(Undelivered.SellerPrice, na.rm = T),
                               MCCost = sum(cost, na.rm = T),
                               MCavgComission = sum(comission*(TotalOrderValue/MCTotalValue), na.rm = T),
                               MCavgCostRatio = sum(costRatio*(TotalOrderValue/MCTotalValue), na.rm = T),
                               MCavgDeliTime = sum(avgDeliTime*(TotalOrderValue/MCDelivered.And.Shipped.Value), na.rm = T)),
                           by = .(CatHead, BDM, KAM, MotherCompany, Month)]

MCReport <- merge(MCInvoiceReport, MCReport, all = T)
rm(MCInvoiceReport)

MCReport <- MCReport[, c("CatHead", "BDM", "KAM", "MotherCompany", "Month",
                         "MCTotalInvoice", "MCProcessingInvoice", "MCPickedInvoice", "MCDeliveredAndShippedInvoice", "MCRefundedInvoice",
                         "MCRevenue", "MCDelivered.And.Shipped.Value", "MCUndeliveredValue",
                         "MCDelivered.And.Shipped.SellerPrice", "MCUndelivered.SellerPrice",
                         "MCCost", "MCavgComission", "MCavgCostRatio", "MCavgRefundRatio", "MCavgDeliTime",
                         "MCSubmitted", "MCApproved", "MCPaid", "MCDue")]

write.csv(MCReport, "Data/Output Data/CommercialReport/MCReport.csv")
sheet_write(MCReport, ss = "https://docs.google.com/spreadsheets/d/1cmlWpMrJ1jn85A6e7xhAXN3Xus5-y7bqzvp6dF1Amuk/", sheet = "MotherCompanyData")


###KAM Report####
KAMInvoiceReport <- shopInvoiceReport[ , .(KAMTotalInvoice = sum(Total.Invoice, na.rm = T),
                                           KAMProcessingInvoice = sum(Processing.Invoice, na.rm = T),
                                           KAMPickedInvoice = sum(Picked.Invoice, na.rm = T),
                                           KAMDeliveredAndShippedInvoice = sum(Delivered.And.Shipped.Inv, na.rm = T),
                                           KAMRefundedInvoice = sum(Refunded.Invoice, na.rm = T),
                                           KAMSubmitted = sum(CorporateSubmitted, na.rm = T),
                                           KAMApproved = sum(CorporateApproved, na.rm = T),
                                           KAMPaid = sum(CorporatePaid, na.rm = T),
                                           KAMDue = sum(CorporateDue, na.rm = T)),
                                       by = .(CatHead, BDM, KAM, Month)] %>%
     mutate(KAMavgRefundRatio = KAMRefundedInvoice/KAMTotalInvoice)


KAMReport <- shopItemReport[, .(KAMRevenue = sum(TotalOrderValue, na.rm = T),
                                KAMDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T),
                                KAMUndeliveredValue = sum(Undelivered.Value, na.rm = T),
                                KAMDelivered.And.Shipped.SellerPrice = sum(Delivered.And.Shipped.SellerPrice, na.rm = T),
                                KAMUndelivered.SellerPrice = sum(Undelivered.SellerPrice, na.rm = T),
                                KAMCost = sum(cost, na.rm = T),
                                KAMavgComission = sum(comission*(TotalOrderValue/KAMTotalValue), na.rm = T),
                                KAMavgCostRatio = sum(costRatio*(TotalOrderValue/KAMTotalValue), na.rm = T),
                                KAMavgDeliTime = sum(avgDeliTime*(TotalOrderValue/KAMDelivered.And.Shipped.Value), na.rm = T)),
                            by = .(CatHead, BDM, KAM, Month)]

KAMReport <- merge(KAMInvoiceReport, KAMReport, all = T)
rm(KAMInvoiceReport)

KAMReport <- KAMReport[, c("CatHead", "BDM", "KAM", "Month",
                           "KAMTotalInvoice", "KAMProcessingInvoice", "KAMPickedInvoice", "KAMDeliveredAndShippedInvoice", "KAMRefundedInvoice",
                           "KAMRevenue", "KAMDelivered.And.Shipped.Value", "KAMUndeliveredValue",
                           "KAMDelivered.And.Shipped.SellerPrice", "KAMUndelivered.SellerPrice",
                           "KAMCost", "KAMavgComission", "KAMavgCostRatio", "KAMavgRefundRatio", "KAMavgDeliTime",
                           "KAMSubmitted", "KAMApproved", "KAMPaid", "KAMDue")]

write.csv(KAMReport, "Data/Output Data/CommercialReport/KAMReport.csv")
sheet_write(KAMReport, ss = "https://docs.google.com/spreadsheets/d/1cmlWpMrJ1jn85A6e7xhAXN3Xus5-y7bqzvp6dF1Amuk/", sheet = "KAMData")

###BDM report####
BDMInvoiceReport <- shopInvoiceReport[ , .(BDMTotalInvoice = sum(Total.Invoice, na.rm = T),
                                           BDMProcessingInvoice = sum(Processing.Invoice, na.rm = T),
                                           BDMPickedInvoice = sum(Picked.Invoice, na.rm = T),
                                           BDMDeliveredAndShippedInvoice = sum(Delivered.And.Shipped.Inv, na.rm = T),
                                           BDMRefundedInvoice = sum(Refunded.Invoice, na.rm = T),
                                           BDMSubmitted = sum(CorporateSubmitted, na.rm = T),
                                           BDMApproved = sum(CorporateApproved, na.rm = T),
                                           BDMPaid = sum(CorporatePaid, na.rm = T),
                                           BDMDue = sum(CorporateDue, na.rm = T)),
                                       by = .(CatHead, BDM, Month)] %>%
     mutate(BDMavgRefundRatio = BDMRefundedInvoice/BDMTotalInvoice)


BDMReport <- shopItemReport[, .(BDMRevenue = sum(TotalOrderValue, na.rm = T),
                                BDMDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T),
                                BDMUndeliveredValue = sum(Undelivered.Value, na.rm = T),
                                BDMDelivered.And.Shipped.SellerPrice = sum(Delivered.And.Shipped.SellerPrice, na.rm = T),
                                BDMUndelivered.SellerPrice = sum(Undelivered.SellerPrice, na.rm = T),
                                BDMCost = sum(cost, na.rm = T),
                                BDMavgComission = sum(comission*(TotalOrderValue/BDMTotalValue), na.rm = T),
                                BDMavgCostRatio = sum(costRatio*(TotalOrderValue/BDMTotalValue), na.rm = T),
                                BDMavgDeliTime = sum(avgDeliTime*(TotalOrderValue/BDMDelivered.And.Shipped.Value), na.rm = T)),
                            by = .(CatHead, BDM, Month)]

BDMReport <- merge(BDMInvoiceReport, BDMReport, all = T)
rm(BDMInvoiceReport)

BDMReport <- BDMReport[, c("CatHead", "BDM", "Month",
                           "BDMTotalInvoice", "BDMProcessingInvoice", "BDMPickedInvoice", "BDMDeliveredAndShippedInvoice", "BDMRefundedInvoice",
                           "BDMRevenue", "BDMDelivered.And.Shipped.Value", "BDMUndeliveredValue",
                           "BDMDelivered.And.Shipped.SellerPrice", "BDMUndelivered.SellerPrice",
                           "BDMCost", "BDMavgComission", "BDMavgCostRatio", "BDMavgRefundRatio", "BDMavgDeliTime",
                           "BDMSubmitted", "BDMApproved", "BDMPaid", "BDMDue")]

write.csv(BDMReport, "Data/Output Data/CommercialReport/BDMReport.csv")
sheet_write(BDMReport, ss = "https://docs.google.com/spreadsheets/d/1cmlWpMrJ1jn85A6e7xhAXN3Xus5-y7bqzvp6dF1Amuk/", sheet = "BDMData")

###Cat Head report####
CatHeadInvoiceReport <- shopInvoiceReport[ , .(CatHeadTotalInvoice = sum(Total.Invoice, na.rm = T),
                                               CatHeadProcessingInvoice = sum(Processing.Invoice, na.rm = T),
                                               CatHeadPickedInvoice = sum(Picked.Invoice, na.rm = T),
                                               CatHeadDeliveredAndShippedInvoice = sum(Delivered.And.Shipped.Inv, na.rm = T),
                                               CatHeadRefundedInvoice = sum(Refunded.Invoice, na.rm = T),
                                               CatHeadSubmitted = sum(CorporateSubmitted, na.rm = T),
                                               CatHeadApproved = sum(CorporateApproved, na.rm = T),
                                               CatHeadPaid = sum(CorporatePaid, na.rm = T),
                                               CatHeadDue = sum(CorporateDue, na.rm = T)),
                                           by = .(CatHead, Month)] %>%
     mutate(CatHeadavgRefundRatio = CatHeadRefundedInvoice/CatHeadTotalInvoice)


CatHeadReport <- shopItemReport[, .(CatHeadRevenue = sum(TotalOrderValue, na.rm = T),
                                    CatHeadDelivered.And.Shipped.Value = sum(Delivered.And.Shipped.Value, na.rm = T),
                                    CatHeadUndeliveredValue = sum(Undelivered.Value, na.rm = T),
                                    CatHeadDelivered.And.Shipped.SellerPrice = sum(Delivered.And.Shipped.SellerPrice, na.rm = T),
                                    CatHeadUndelivered.SellerPrice = sum(Undelivered.SellerPrice, na.rm = T),
                                    CatHeadCost = sum(cost, na.rm = T),
                                    CatHeadavgComission = sum(comission*(TotalOrderValue/CatHeadTotalValue), na.rm = T),
                                    CatHeadavgCostRatio = sum(costRatio*(TotalOrderValue/CatHeadTotalValue), na.rm = T),
                                    CatHeadavgDeliTime = sum(avgDeliTime*(TotalOrderValue/CatHeadDelivered.And.Shipped.Value), na.rm = T)),
                                by = .(CatHead, Month)]

CatHeadReport <- merge(CatHeadInvoiceReport, CatHeadReport, all = T, by = c("CatHead", "Month"))
rm(CatHeadInvoiceReport)

CatHeadReport <- CatHeadReport[, c("CatHead", "Month",
                                   "CatHeadTotalInvoice", "CatHeadProcessingInvoice", "CatHeadPickedInvoice", "CatHeadDeliveredAndShippedInvoice", "CatHeadRefundedInvoice",
                                   "CatHeadRevenue", "CatHeadDelivered.And.Shipped.Value", "CatHeadUndeliveredValue",
                                   "CatHeadDelivered.And.Shipped.SellerPrice", "CatHeadUndelivered.SellerPrice",
                                   "CatHeadCost", "CatHeadavgComission", "CatHeadavgCostRatio", "CatHeadavgRefundRatio", "CatHeadavgDeliTime",
                                   "CatHeadSubmitted", "CatHeadApproved", "CatHeadPaid", "CatHeadDue")]

write.csv(CatHeadReport, "Data/Output Data/CommercialReport/CatHeadReport.csv")
sheet_write(CatHeadReport, ss = "https://docs.google.com/spreadsheets/d/1cmlWpMrJ1jn85A6e7xhAXN3Xus5-y7bqzvp6dF1Amuk/", sheet = "CatHeadData")

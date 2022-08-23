## install and load packages
install.packages("openxlsx")
install.packages("tidyverse")
install.packages("lubridate")
install.packages("viridis")
library(openxlsx)
library(tidyverse)
library(lubridate)
library(viridis)

## import datasets
addstocount <- read_csv("DataAnalyst_Ecom_data_addsToCart.csv")
sessioncounts <- read_csv("DataAnalyst_Ecom_data_sessionCounts.csv")

## explore the dimensions of the datasets
glimpse(addstocount)
head(addstocount)
glimpse(sessioncounts)
head(sessioncounts)

## recharacterize columns in sessioncounts for analysis
sessioncounts$dim_browser <- as.factor(sessioncounts$dim_browser)
sessioncounts$dim_deviceCategory <- as.factor(sessioncounts$dim_deviceCategory)
sessioncounts$dim_date <- mdy(sessioncounts$dim_date)

## extract year and month from sessioncounts
sessioncounts <- sessioncounts |>
        mutate(dim_year = year(dim_date), dim_month = month(dim_date)) |>
        select(dim_browser:dim_date, dim_year, dim_month, everything())

## Sheet 1: Month_by_Device aggregation
## summarize sessioncounts into a new dataframe/tibble
sescts <- sessioncounts |>
        group_by(dim_year, dim_month, dim_deviceCategory) |>
        summarize(Sessions = sum(sessions),
                  Transactions = sum(transactions),
                  QTY = sum(QTY))

## add ECR column to summary dataset: sescts
## This output will create the data for Sheet 1
sescts <- sescts |>
        mutate(ECR = Transactions/Sessions)

## Sheet 2: Month_over_Month metrics
## subset sessioncounts to the last three months and summarize
sescts02 <- sessioncounts |>
        filter(dim_month == c(5:6)) |>
        group_by(dim_month) |>
        summarize(Sessions = sum(sessions),
                  Transactions = sum(transactions),
                  QTY = sum(QTY),
                  ECR = Transactions / Sessions)

## create a new datasframe: add addsToCart column to sescts02
monthly <- left_join(
        sescts02,
        addstocount) |>
        select(-dim_year)

## Transpose dataset
monthly_long <- as.data.frame(t(monthly))

## Rename transposed columns
colnames(monthly_long) <- c("last_month", "this_month")

## change row names to named column "metrics"
monthly02 <- rownames_to_column(monthly_long, "metrics")

## Calculate delta and abs(delta) columns
## This (?) is the desired output for sheet 2 (?), long instead of wide
monthly02 <- monthly02 |>
        mutate(
                delta = this_month - last_month,
                delta_abs = abs(delta)
                )

## Create the .xlsx file using openxlsx
## Sheet 1 will be the month-by-device aggregation
## additional code will open the workbook just created,
## add a new sheet,
## append the delta data for the most recent two months' metrics,
## and save the workbook with the new data
write.xlsx(sescts, "ixis_ref_tables.xlsx", asTable = FALSE,
           sheetName = "Month_by_Device")
wb <- (loadWorkbook("ixis_ref_tables.xlsx"))
addWorksheet(
        wb,
        sheetName = "Month_over_Month")
writeData(wb, "Month_over_Month", monthly02)
saveWorkbook(wb, "ixis_ref_tables.xlsx", overwrite = TRUE)

## Now to create some graphs for the presentation
## Start with the session counts by device, and join addstocount data
sescts_tot <- left_join(
        sescts,
        addstocount
)


## This function will create new factor var and levels
## write function fct_case_when to convert to factor AND specify levels in order of case_when statements
fct_case_when <- function(...) {
        args <- as.list(match.call())
        levels <- sapply(args[-1], function(f) f[[3]])  # extract RHS of formula
        levels <- levels[!is.na(levels)]
        factor(dplyr::case_when(...), levels=levels)
}

## now create a new variable for month/year combo
sescts_tot <- sescts_tot |>
        mutate(month = fct_case_when(
                dim_month == "7" ~ "July 2012",
                dim_month == "8" ~ "Aug 2012",
                dim_month == "9" ~ "Sept 2012",
                dim_month == "10" ~ "Oct 2012",
                dim_month == "11" ~ "Nov 2012",
                dim_month == "12" ~ "Dec 2012",
                dim_month == "1" ~ "Jan 2013",
                dim_month == "2" ~ "Feb 2013",
                dim_month == "3" ~ "Mar 2013",
                dim_month == "4" ~ "Apr 2013",
                dim_month == "5" ~ "May 2013",
                dim_month == "6" ~ "June 2013"
        ))

## Verify dataset is a dataframe, just in case!
sescts_tot <- as.data.frame(sescts_tot)

## rearrange the device levels in order of their ECR:
## desktop, then tablet, then mobile
sescts_tot$dim_deviceCategory <- fct_relevel(sescts_tot$dim_deviceCategory, 
                                             "desktop", "tablet", "mobile")

## Create a bar cart comparing ECR among devices over time
sescts_tot |>
        ggplot(aes(x = month, y = ECR, 
                   fill = dim_deviceCategory)) + 
        geom_col(position = "dodge") +
        scale_fill_viridis(discrete = TRUE,
                           name = "Device",
                           labels = c("Desktop", "Tablet", "Mobile")) +
        theme_classic() +
        theme(
                axis.text.x = element_text(angle = -30, hjust = 0, vjust = 1)
        ) +
        labs(
                x = NULL,
                y = "E-commerce Conversion Rate (ECR)"
        )
        
## Create a bar cart comparing Sessions among devices over time
sescts_tot |>
        ggplot(aes(x = month, y = Sessions, 
                   fill = dim_deviceCategory)) + 
        geom_col(position = "dodge") +
        scale_fill_viridis(discrete = TRUE,
                           name = "Device",
                           labels = c("Desktop", "Tablet", "Mobile")) +
        scale_y_continuous(labels = scales::label_comma(accuracy = 1)) +
        theme_classic() +
        theme(
                axis.text.x = element_text(angle = -30, hjust = 0, vjust = 1)
        ) +
        labs(
                x = NULL,
                y = "Number of sessions"
        )

## Create a bar cart comparing Transactions among devices over time
sescts_tot |>
        ggplot(aes(x = month, y = Transactions, 
                   fill = dim_deviceCategory)) + 
        geom_col(position = "dodge") +
        scale_fill_viridis(discrete = TRUE,
                           name = "Device",
                           labels = c("Desktop", "Tablet", "Mobile")) +
        scale_y_continuous(labels = scales::label_comma(accuracy = 1)) +
        theme_classic() +
        theme(
                axis.text.x = element_text(angle = -30, hjust = 0, vjust = 1)
        ) +
        labs(
                x = NULL,
                y = "Number of transactions"
        )

## scatterplot of Session vs Transactions
sescts_tot |>
        ggplot(aes(x = Sessions, y = Transactions, 
                   color = dim_deviceCategory)) +
        geom_point(size = 2, alpha = 0.75) +
        scale_color_viridis(discrete = TRUE,
                           name = "Device",
                           labels = c("Desktop", "Tablet", "Mobile")) +
        scale_y_continuous(labels = scales::label_comma(accuracy = 1)) +
        scale_x_continuous(labels = scales::label_comma(accuracy = 1)) +        
        theme_classic() +
        theme(
                axis.text.x = element_text(angle = -30, hjust = 0, vjust = 1)
        ) +
        labs(
                x = "Number of sessions",
                y = "Number of transactions"
        )
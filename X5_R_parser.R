library(data.table)
library(dplyr)

options(java.parameters = "-Xmx8g")
library(XLConnect)

esxtop_convert <- function() {
  w_list <- c("MSK-DPRO-FNR004", "MSK-DPRO-FNR024", "MSK-DPRO-FNR026",
              "MSK-DPRO-FNR028", "MSK-DPRO-ORA111", "MSK-DPRO-ORA112",
              "MSK-DPRO-ORA109", "MSK-DPRO-ORA084", "MSK-DPRO-FNR019",
              "MSK-DPRO-FNR009", "MSK-DPRO-HPO002", "MSK-DPRO-HPO003",
              "MSK-DPRO-ORA122", "MSK-DPRO-ORA123", "MSK-DPRO-ORA149",
              "MSK-DPRO-HPO004", "MSK-DPRO-ORA134", "MSK-DPRO-ORA135")
  fl <- list.files(pattern="^msk.+csv", all.files=F, full.names=F, recursive=F)
  
  for(i in fl) {
    print(i)
    esxtop <- esxtop_read(i) 
    
    if (length(esxtop) > 0)
      write.csv(select(esxtop, 1, matches(paste(w_list, collapse="|"))), 
                paste0("out_", i), sep = ",", dec = ".", row.names = F)
  }
}

esxtop_read <- function(f) {
  match_str <- c("Used", "Run", "System", "% Wait", "Ready", "CoStop",
                 "Members", "Overlap")
  
  period <- 3 # seconds
  time_frame <- 13 # hours
  s_time <- "04/26/2017 19:00"
  
  header <- colnames(fread(f, nrows = 0, header = T))
  
  s_string <- paste0("\\\\", 
                     paste(unlist(strsplit(f, ".", fixed = T))[-4], 
                           collapse = "."), 
                     "\\Group Cpu")
  
  sel <- substring(header, 1, nchar(s_string)) %chin% s_string
  
  out <- fread(f, sep = ",", header = F, dec = ".", stringsAsFactors = F, 
               skip = s_time, nrows = time_frame*3600/period, 
               col.names = c("ts_string", header[sel]), 
               select = c(1, which(sel)))
  colnames(out) <- gsub("^.+Group\\s(.+)\\((\\d+):(.+)\\)\\\\(.+$)", 
                        "\\3.\\1.\\4", colnames(out))
  
  out[, names(out) := lapply(.SD, type.convert, as.is = TRUE)] %>% 
    select(1, matches(paste(match_str, collapse="|")))
}

esxtop_read2 <- function(f) {
  match_str <- c("\\Commands/sec")
  
  period <- 3 # seconds
  time_frame <- 1 # hours
  s_time <- "04/26/2017 19:00"
  
  header <- colnames(fread(f, nrows = 0, header = T))
  
  s_string <- paste0("\\\\", 
                     paste(unlist(strsplit(f, ".", fixed = T))[-4], 
                           collapse = "."), 
                     "\\Physical Disk Per-Device-Per-World")
  
  sel <- substring(header, 1, nchar(s_string)) %chin% s_string
  
  out <- fread(f, sep = ",", header = F, dec = ".", stringsAsFactors = F, 
               skip = s_time, nrows = time_frame*3600/period, 
               col.names = c("ts_string", header[sel]), 
               select = c(1, which(sel)))
  
  #  pdisk <- melt(out, id.vars = 1) %>% 
  #    mutate(lun = gsub("^.+\\((.+\\d+:C\\d+:T\\d+:L\\d+).+$", "\\1", variable),
  #           world = gsub("^.+-(\\d+)\\).+$", "\\1", variable),
  #           measure = gsub("^.+\\\\(.+)$", "\\1", variable)) %>%
  #    filter(measure %in% pdisk_filter) %>%
  #    select(ts = 1, world, lun, measure, value) %>% arrange(ts, world, lun)
  
  out[, names(out) := lapply(.SD, type.convert, as.is = TRUE)] %>% 
    select(1, matches(paste(match_str, collapse="|")))
}

outToExcel <- function(x) {
  FILENAME <- "out.xlsx"
  
  if(file.exists(FILENAME)) file.remove(FILENAME)
  
  wb <- loadWorkbook(FILENAME, create = TRUE)
  createSheet(wb, name = "Raw")
  createName(wb, name = "raw_data", formula = "Raw!$A$1")
  
  csDate = createCellStyle(wb)
  setDataFormat(csDate, format = "mm/dd/yyyy HH:mm:ss")
  csNumber = createCellStyle(wb)
  setDataFormat(csNumber, format = "0.00")
  
  writeNamedRegion(wb, x, name = "raw_data", header = T)
  
  setColumnWidth(wb, sheet = "Raw", col = seq(1, ncol(x)))
  setCellStyle(wb, sheet = "Raw", col = 1, row = seq(2, nrow(x) + 1), 
               cellstyle = csDate)
  
  rc = expand.grid(row = seq(2, nrow(x) + 1), col = seq(3, ncol(x)))
  
  setCellStyle(wb, sheet = "Raw", col = rc$col, row = rc$row, 
               cellstyle = csNumber)
  
  saveWorkbook(wb)
  
  xlcFreeMemory()
}
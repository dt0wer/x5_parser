library(data.table)
library(dplyr)

options(java.parameters = "-Xmx8g")
library(XLConnect)

esxtop_convert <- function() {
#    w_list <- c("MSK-DPRO-FNR004", "MSK-DPRO-FNR024", "MSK-DPRO-FNR026",
#                "MSK-DPRO-FNR028", "MSK-DPRO-ORA111", "MSK-DPRO-ORA112",
#                "MSK-DPRO-ORA109", "MSK-DPRO-ORA084", "MSK-DPRO-FNR019",
#                "MSK-DPRO-FNR009", "MSK-DPRO-HPO002", "MSK-DPRO-HPO003",
#                "MSK-DPRO-ORA122", "MSK-DPRO-ORA123", "MSK-DPRO-ORA149",
#                "MSK-DPRO-HPO004", "MSK-DPRO-ORA134", "MSK-DPRO-ORA135")

      w_list <- c("MSK-DPRO-APP199", "MSK-DPRO-ORA155", "MSK-DPRO-ORA085",
                  "MSK-DPRO-FNR002", "MSK-DPRO-FNR003", "MSK-DPRO-FNR014",
                  "MSK-DPRO-FNR015")
    
  fl <- list.files(pattern="^msk.+csv", all.files=F, full.names=F, recursive=F)
  
  for(i in fl) {
    print(i)
    esxtop <- esxtop_read2(i, w_list) 
    
    if (length(esxtop) > 0)
      write.csv(esxtop, paste0("full_", i), sep = ",", dec = ".", 
                row.names = F)
  }
}

esxtop_read <- function(f) {
  match_str <- c("Used", "Run", "System", "% Wait", "Ready", "CoStop",
                 "Members", "Overlap")
  
  period <- 3 # seconds
  time_frame <- 13 # hours
  s_time <- "04/27/2017 01:05"
  
  header <- colnames(fread(f, nrows = 0, header = T))
  
  s_string <- paste0("\\\\", 
                     paste(unlist(strsplit(f, ".", fixed = T))[1], 
                           collapse = "."), 
                     "\\Group Cpu")
  
  sel <- substring(header, 1, nchar(s_string)) %chin% s_string
  
  out <- fread(f, sep = ",", header = F, dec = ".", stringsAsFactors = F, 
               skip = s_time, nrows = time_frame*3600/period,
               col.names = c("ts_string", header[sel]), 
               select = c(1, which(sel)))
  colnames(out) <- gsub("^.+Group\\s(.+)\\((\\d+):(.+)\\)\\\\(.+$)", 
                        "\\3.\\1.\\4", colnames(out))
  
  # convert types
  out[, names(out) := lapply(.SD, type.convert, as.is = TRUE)] %>% 
    select(1, matches(paste(match_str, collapse="|")))
}


esxtop_read2 <- function(f, w_list) {
  match_str_cpu <- c("% Used", "Ready", "CoStop")
  match_str_vdisk <- c("\\Commands/sec","MBytes Read/sec", "MBytes Written/sec")
  
  period <- 3 # seconds
  time_frame <- 13 # hours
  s_time <- "06/01/2017 21:"
  
  header <- colnames(fread(f, nrows = 0, header = T))
  
  sel <- grepl(paste(w_list, collapse="|"), header) & 
    ((grepl(paste(match_str_vdisk, collapse="|"), header) &
       !grepl("scsi", header)) | 
       ((grepl(paste(match_str_cpu, collapse="|"), header) &
           grepl("Group Cpu", header))))
    
  
  out <- fread(f, sep = ",", header = F, dec = ".", stringsAsFactors = F, 
               skip = s_time, nrows = time_frame*3600/period, 
               col.names = c("ts_string", header[sel]), 
               select = c(1, which(sel)))
  
  colnames(out) <- gsub("-", "_", gsub("\\s|%|:|\\(|\\)", ".", 
                                       gsub("^\\\\{2}.+\\\\(.+)\\\\(.+)$", 
                                            "\\1\\2", colnames(out))))
  out[, names(out) := lapply(.SD, type.convert, as.is = TRUE)]
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
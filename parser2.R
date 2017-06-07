library(data.table)
library(dplyr)

sxtop_convert <- function() {
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


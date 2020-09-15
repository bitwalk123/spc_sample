library(ggplot2)
library(ggQC)
library(ggpubr)
library(openxlsx)
library(qcc)
library(Rspc)
library(tcltk)

# file name definition to read (dialog)
filters <- matrix(c("Excel files", ".xlsx", "All files", "*"), 2, 2, byrow = TRUE)
filename <- tk_choose.files(filter = filters)
if (length(filename) == 0) {
  stop("The value is TRUE, so the script must end here")
}

# read Excel macro sheet
key.master <- "Master"
list.sheet <- getSheetNames(filename)
if (!(key.master %in% list.sheet)) {
  stop("The value is TRUE, so the script must end here")
}

# get master sheet
index.master <- match(key.master, list.sheet)
sheet.master <- read.xlsx(filename, sheet = index.master)

# eliminate NA rows
sheet.master <- sheet.master[!is.na(sheet.master[, 1]), ]

# get data sheet
col.part <- 'Part.Number'
list.part <- sort(unique(sheet.master[, col.part]))
for (name.sheet in list.part) {
  index.sheet <- match(name.sheet, list.sheet)
  assign(name.sheet, read.xlsx(filename, sheet = index.sheet))
}

# prepare empty table for results
n <- 7
tbl.result <- data.frame(matrix(rep(NA, n), nrow=1))[numeric(0), ]
colnames(tbl.result) <- c('Parameter.Name', 'LSL', 'Target', 'USL', 'LCL', 'UCL', 'Normality')

# plot
col.param <- 'Parameter.Name'
for (name.param in sheet.master[, col.param]) {
  print(paste('###', name.param))
  name.part <- sheet.master[sheet.master[, col.param] == name.param, col.part]
  name.param0 <- gsub(' ', '.', name.param)
  tmp <- get(name.part)
  x <- tmp[, 'Sample']
  y <- tmp[, name.param0]
  
  # other indexes
  lsl <- sheet.master[sheet.master[, col.param] == name.param, 'LSL']
  usl <- sheet.master[sheet.master[, col.param] == name.param, 'USL']
  target <- sheet.master[sheet.master[, col.param] == name.param, 'Target']

  # 'Rspc' package 
  limits = CalculateLimits(y, type = "i")
  lcl <- limits$lcl
  ucl <- limits$ucl

  # trend
  q <- qcc(y, type="xbar.one", nsigmas = 3,0, limits = c(lcl, ucl),
           title = name.param, ylab = 'Value', plot = TRUE)
  q$data.name <- name.param
  process.capability(q, spec.limits = c(lsl, usl), target = target)

  # QQ plot to check normality
  qq <- ggqqplot(y, title = name.param)
  plot(qq)
  
  # to perform the Shapiro-Wilk test of normality for univariate
  normality <- shapiro.test(y)

  # summary
  index.row <- nrow(tbl.result) + 1
  tbl.result[index.row,] <- c(name.param, lsl, target, usl, lcl, ucl, normality$p.value)
}

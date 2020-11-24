## C++ DLL multiplyBy console example
args <- commandArgs(trailingOnly = TRUE)
if (length(args) == 0) {
  require(rstudioapi)
  rPath <- dirname(getActiveDocumentContext()$path)
  fpath <- file.path(rPath, '..', 'dll', 'multiplyBy x64.dll')
  dl <- choose.files(default = fpath, caption = "Select files",
                     multi = TRUE, filters = Filters,
                     index = nrow(Filters))
} else {
  dl <- args[1]
}
cppMultiplyBy <- function(x, y, digit = 0, ...) {
  #require(rstudioapi)
  #rPath <- dirname(getSourceEditorContext()$path)
  #rPath <- dirname(getActiveDocumentContext()$path)
  x <- as.double(gsub(",|'", "", x))
  y <- as.double(gsub(",|'", "", y))
  #dyn.load(file.path(rPath, '..', 'dll', 'multiplyBy x64.dll'))
  dyn.load(dl)
  res <- .C("multiplyBy", x, outdata = y)$outdata
  #dyn.unload(file.path(rPath, '..','dll', 'multiplyBy x64.dll'))
  dyn.unload(dl)
  if (!missing(...)) {
    if (digit != 0) res <- round(res, digit)
    noquote(format(res, ...))
  } else {
    res
  }
}
cppMultiplyBy(110.55, 11.23)
cppMultiplyBy(110.55, 11.23, big.mark = ",", digit = 2)
cppMultiplyBy(format(1010.55, big.mark = ","), format(231.22, big.mark = "'"), big.mark = ",", digit = 2)

gui <- function() {
  require(rstudioapi)
  require(tcltk)
  
  addRow <- function(i) {
    cols <- NULL
    for (j in seq(numcol)) {
      ent <- NULL
      if (j < numcol) {
        ent <- tkentry(calc, relief="ridge", justify="right")
        tkgrid(ent, row=i+1, column=j, sticky="nsew")
        tkinsert(ent, "end", "")
        tkbind(ent,'<Key-Return>', function(x, ...) onChange(i))
        tkbind(ent,'<Key-Tab>', function(x, ...) onChange(i))
        tkbind(ent, '<Control-c>', function(x) toClipboard())
      } else {
        ent <- tklabel(calc, text="0", relief="ridge", anchor="e")
        tkgrid(ent, row=i+1, column=j, sticky="nsew")
      }
      if (j == 1) {
        cols <- ent
      } else {
        cols <- c(cols, ent)
      }
    }
    names(cols) <- nms
    rows <<- rbind(rows, t(as.matrix(cols)))
  }
  
  centimetersToPoints <- function(centimeter) {
    res <- centimeter * 28.3464567
  }
  
  frmtChange <- function() {
    for (i in seq(numrow)) {
      num <- restyle(1, normVal(tclvalue(tkget(rows$Number1[[i]]))))
      tkdelete(rows$Number1[[i]], 0, "end")
      tkinsert(rows$Number1[[i]], 0, num)
      num <- restyle(1, normVal(tclvalue(tkget(rows$Number2[[i]]))))
      tkdelete(rows$Number2[[i]], 0, "end")
      tkinsert(rows$Number2[[i]], 0, num)
      tkconfigure(rows$`Multiplied Result`[[i]], 
                  text = restyle(1, normVal(tclvalue(tkcget(rows$`Multiplied Result`[[i]], "-text")))))
    }
    tkconfigure(sums$Number1[[1]], text = restyle(2, normVal(tclvalue(tkcget(sums$Number1[[1]], "-text")))))
    tkconfigure(sums$Number2[[1]], text = restyle(2, normVal(tclvalue(tkcget(sums$Number2[[1]], "-text")))))
    tkconfigure(sums$`Multiplied Result`[[1]], text = 
                  restyle(2, normVal(tclvalue(tkcget(sums$`Multiplied Result`[[1]], "-text")))))
  }
  
  makeWidgets <- function() {
    hdr <<- as.data.frame(t(matrix(rep(NA, 6), dimnames = list(nms))))[0,]
    colnames(hdr) <- nms
    cols <- NULL
    for (i in seq(numcol)) {
      lab <- tklabel(calc, text=hdrs[i], relief="ridge", width=30)
      tkgrid(lab, row=0, column=i, sticky="nsew")
      tkconfigure(lab, bg=bgColour$Hex, fg=fgColour$Hex)
      if (i == 1){
        cols <- lab
      } else {
        cols <- c(cols, lab)
      }
    }
    names(cols) <- nms
    hdr <<- rbind(hdr, t(as.matrix(cols)))
    
    rows <<- as.data.frame(t(matrix(rep(NA, 6), dimnames = list(nms))))[0,]
    colnames(rows) <- nms
    for (i in seq(numrow)) {
      addRow(i)
    }
    
    sums <<- as.data.frame(t(matrix(rep(NA, 6), dimnames = list(nms))))[0,]
    colnames(sums) <- nms
    cols <- NULL
    for (i in seq(numcol)) {
      ent <- tklabel(calc, text="0", relief="sunken", anchor="e")
      tkgrid(ent, row=40, column=i, sticky="nsew")
      if (i == 1) {
        cols <- ent
      } else {
        cols <- c(cols, ent)
      }
    }
    names(cols) <- nms
    sums <<- rbind(sums, t(as.matrix(cols)))
    
    txtBg <- tklabel(toolBar, text="Header Background Colour ") 
    tkgrid(txtBg, row=0, column=0)
    bgRed <<- ttkcombobox(toolBar, textvariable=bgred, values=gettext(0:255), width=5)
    tkbind(bgRed, "<<ComboboxSelected>>", function(x) updColours())
    tkbind(bgRed, "<Key-Tab>", function(x) updColours())
    tkbind(bgRed, "<Key-Return>", function(x) updColours())
    tkgrid(bgRed, row=0, column=1)
    bgGreen <<- ttkcombobox(toolBar, textvariable=bggreen, values=gettext(0:255), width=5)
    tkbind(bgGreen, "<<ComboboxSelected>>", function(x) updColours())
    tkbind(bgGreen, "<Key-Tab>", function(x) updColours())
    tkbind(bgGreen, "<Key-Return>", function(x) updColours())
    tkgrid(bgGreen, row=0, column=2)
    bgBlue <<- ttkcombobox(toolBar, textvariable=bgblue, values=gettext(0:255), width=5)
    tkbind(bgBlue, "<<ComboboxSelected>>", function(x) updColours())
    tkbind(bgBlue, "<Key-Tab>", function(x) updColours())
    tkbind(bgBlue, "<Key-Return>", function(x) updColours())
    tkgrid(bgBlue, row=0, column=3)
    
    txtfg <- tklabel(toolBar, text="    Fore Colour ") 
    tkgrid(txtfg, row=0, column=4)
    fgRed <<- ttkcombobox(toolBar, textvariable=fgred, values=gettext(0:255), width=5)
    tkbind(fgRed, "<<ComboboxSelected>>", function(x) updColours())
    tkbind(fgRed, "<Key-Tab>", function(x) updColours())
    tkbind(fgRed, "<Key-Return>", function(x) updColours())
    tkgrid(fgRed, row=0, column=5)
    fgGreen <<- ttkcombobox(toolBar, textvariable=fggreen, values=gettext(0:255), width=5)
    tkbind(fgGreen, "<<ComboboxSelected>>", function(x) updColours())
    tkbind(fgGreen, "<Key-Tab>", function(x) updColours())
    tkbind(fgGreen, "<Key-Return>", function(x) updColours())
    tkgrid(fgGreen, row=0, column=6)
    fgBlue <<- ttkcombobox(toolBar, textvariable=fgblue, values=gettext(0:255), width=5)
    tkbind(fgBlue, "<<ComboboxSelected>>", function(x) updColours())
    tkbind(fgBlue, "<Key-Tab>", function(x) updColours())
    tkbind(fgBlue, "<Key-Return>", function(x) updColours())
    tkgrid(fgBlue, row=0, column=7)
    
    formats.list <<- ttkcombobox(self, textvariable=var1, values=unlist(formats[, 1]))
    tkbind(formats.list, "<<ComboboxSelected>>", function(x) onChange(-1))
    tkgrid(formats.list, row=2, column=1)
    
    choices = c("", "Add Row", 
               "Send to New Email", "Send to Open Email", 
               "Send to New Word Document", "Send to Open Word Document" 
    )
    functions <<- ttkcombobox(self, textvariable=var2, values=choices)
    tkbind(functions, "<<ComboboxSelected>>", function(x) runFunction())
    tkgrid(functions, row=2, column=2, sticky="nsew")
  }

  multiplyBy <- function(x, y) {
    res <- .C("multiplyBy", normVal(x), outdata = normVal(y))$outdata
    res
  }
  
  normVal <- function(x) {
    x <- gsub(",|'", "", x)
    if (is.na(x) || is.null(x) || x == "") {
      x <- as.double(0)
    } else {
      x <- as.double(x)
    }
    x
  }
  
  onChange <- function(row) {
    if (row >= 1) {
      tkconfigure(rows$`Multiplied Result`[[row]],
        text = multiplyBy(tclvalue(tkget(rows$Number1[[row]])),
        tclvalue(tkget(rows$Number2[[row]]))))
    }
    tkconfigure(sums$Number1[[1]], text =  
                  sum(sapply(seq(nrow(rows)), function(x) normVal(tclvalue(tkget(rows$Number1[[x]]))))))
    tkconfigure(sums$Number2[[1]], text = 
                  sum(sapply(seq(nrow(rows)), function(x) normVal(tclvalue(tkget(rows$Number2[[x]]))))))
    tkconfigure(sums$`Multiplied Result`[[1]], text = 
                  sum(sapply(seq(nrow(rows)), function(x) 
                    normVal(tclvalue(tkcget(rows$`Multiplied Result`[[x]], "-text"))))))
    
    frmtChange()
  }
  
  restyle <- function(typ, num) {
    if (typ == 1) {
      big <- format(as.numeric(strsplit(sprintf(formats[1,]$Format, num), ".", fixed = T)[[1]][1]), big.mark = ",")
      small <- noquote(strsplit(sprintf(formats[1,]$Format, num), ".", fixed = T)[[1]][2])
    } else {
      big <- format(as.numeric(strsplit(sprintf(
        formats[formats$Type==tclvalue(var1),]$Format, num), ".", fixed = T)[[1]][1]), big.mark = ",")
      small <- noquote(strsplit(sprintf(
        formats[formats$Type==tclvalue(var1),]$Format, num), ".", fixed = T)[[1]][2])
    }
    if (is.na(small)) {
      res <- big
    } else if (nchar(as.character(small)) > 0) {
      res <- paste(big, small, sep = ".")
    } else {
      res <- big
    }
    res
  }
  
  RGB <- function(col) {
    res <- apply(as.matrix(col * c(1, 256, 65536), nrow=1), 2, sum)
    res
  }
  
  runFunction <- function() {
    if (tclvalue(var2) != "") {
      if (tclvalue(var2) == "Add Row") {
        numrow <<- numrow + 1
        addRow(numrow)
      }
      if (tclvalue(var2) == "Send to New Email") {
        tbl = tblHtml()
        sendToEmail("New", tbl[[1]])
      }
      if (tclvalue(var2) == "Send to Open Email") {
        tbl = tblHtml()
        sendToEmail("Open", tbl[[1]])
      }
      if (tclvalue(var2) == "Send to New Word Document") {
        tbl = "" #tblHtml()
        sendToWord("New", tbl)
      }
      if (tclvalue(var2) == "Send to Open Word Document") {
        tbl = "" #tblHtml()
        sendToWord("Open", tbl)
      }
      
      tkset(functions, "")
    }
  }
  
  sendToEmail <- function(typ, htm) {
    # devtools::install_github("dkyleward/RDCOMCLIENT")
    require(RDCOMClient)
    olApp <- COMCreate("Outlook.Application")
    if (typ == "New") {
      olApp <- COMCreate("Outlook.Application")
      oMail <- olApp$CreateItem(0)
      oMail$Display()
      oMail[["subject"]] <- "C++ multiplyBy R Studio Test"
      oMail[["htmlbody"]] <- htm
    } else {
      oMail <- olApp$ActiveInspector()[["CurrentItem"]]
      word <- olApp$ActiveInspector()[["WordEditor"]]
      wrdApp <- word[["Application"]]
      wrdApp[["Selection"]] <- "placeHere"
      oMail[["htmlbody"]] = gsub("placeHere", htm, oMail[["htmlbody"]])
    }
  }
  
  sendToWord <- function(typ, htm){
    # devtools::install_github("dkyleward/RDCOMCLIENT")
    require(RDCOMClient)
    wdOrientLandscape = 1; 
    word = COMCreate("Word.Application")
    if (typ == "New"){
      word[["Visible"]] <- TRUE
      doc <<- word$Documents()$Add()
      doc[["PageSetup"]][["Orientation"]] <- wdOrientLandscape;
      cent <- centimetersToPoints(1.75)
      for(i in c("TopMargin", "LeftMargin", "BottomMargin", "RightMargin")) {
        doc[["PageSetup"]][[i]] <- cent
      }
    } else {
      doc <<- word[["ActiveDocument"]]
    }
    tblWord()
  }
  
  tblHtml <- function() {
    stl <- "<style>\n"
    stl <- paste0(stl, "table {\n")
    stl <- paste0(stl, "\tborder-collapse: collapse;\n")
    stl <- paste0(stl, "}\n")
    stl <- paste0(stl, "table, th, td {\n")
    stl <- paste0(stl, "\tpadding: 0px 5px 0px 5px;\n")
    stl <- paste0(stl, "\tborder: 1 solid black;\n")
    stl <- paste0(stl, "\tfont-size: 1em;\n")
    stl <- paste0(stl, "}\n")
    stl <- paste0(stl, "td {\n")
    stl <- paste0(stl, "\tcolor: black;\n")
    stl <- paste0(stl, "\ttext-align: right;\n")
    stl <- paste0(stl, "}\n")
    stl <- paste0(stl, "th {\n")
    stl <- paste0(stl, "\tcolor: rgb(", fgColour$RGB$Red, ", ",  fgColour$RGB$Green, ", ", fgColour$RGB$Blue, ");\n")
    stl <- paste0(stl, "\tbackground-color: rgb(", bgColour$RGB$Red, ", ",  bgColour$RGB$Green, ", ", bgColour$RGB$Blue, ");\n")
    stl <- paste0(stl, "}\n")
    stl <- paste0(stl, "</style>\n")
    tbl <- "<table id=\"mulitplyBy\">\n"
    trd <- "\t<tr>\n"
    tre <- "\n\t</tr>\n"
    thd <- "\t\t<th>\n\t\t\t"
    the <- "\n\t\t</th>\n"
    tdd <- "\t\t<td>\n\t\t\t"
    tde <- "\n\t\t</td>\n"
    th <- ""
    tr <- ""
    tt <- ""
    for (col in seq(numcol)) {
      th <- paste0(th, thd, hdrs[col], the)
    }
    tbl <- paste0(tbl, th)
    for (row in seq(numrow)){
      td <- ""
      td <- paste0(td, tdd, restyle(2, normVal(tclvalue(tkget(rows$Number1[[row]])))), tde)
      td <- paste0(td, tdd, restyle(2, normVal(tclvalue(tkget(rows$Number2[[row]])))), tde)
      td <- paste0(td, tdd, restyle(2, normVal(tclvalue(tkcget(rows$`Multiplied Result`[[row]], "-text")))), tde)
      tr <- paste0(tr, trd, td, tre)
    }
    
    tr <- paste0(tr, trd, "<td colspan=\"3\" style=\"border-left: none; border-right: none;\">  &nbsp; </td>", tre)         
    tt <- paste0(tt, tdd, tclvalue(tkcget(sums$Number1[[1]], "-text")), tde)
    tt <- paste0(tt, tdd, tclvalue(tkcget(sums$Number2[[1]], "-text")), tde)
    tt <- paste0(tt, tdd, tclvalue(tkcget(sums$`Multiplied Result`[[1]], "-text")), tde)
    tr <- paste0(tr, trd, tt, tre)
    tbl <- paste0(tbl, tr, "</table>")
    stl <- paste0(stl, tbl)
    stl
  }
  
  tblWord <- function() {
    wdActiveEndPageNumber = 3; wdBorderLeft = -2; wdBorderRight = -4; wdBorderVertical = -6; wdLineStyleNone = 0;
    wdLineStyleSingle = 1; wdAlignParagraphRight = 2; wdAlignParagraphLeft = 0;
    tbl <- doc[["Application"]][["Selection"]][["Tables"]]$Add(
      doc[["Application"]][["Selection"]][["Range"]], numrow + 3, numcol
    )
    for (i in c("InsideLineStyle", "OutsideLineStyle")) {
      tbl[["Borders"]][[i]] <- wdLineStyleSingle
    }
    for (i in c("InsideColor", "OutsideColor")) {
      tbl[["Borders"]][[i]] <- RGB(unlist(bgColour$RGB))
    }
    pg <- tbl[["Rows"]][[1]][["Range"]]$Information(wdActiveEndPageNumber)
    i <- 1
    for (j in seq(numcol)){
      rng <- tbl[["Rows"]][[i]][["Cells"]][[j]][["range"]]
      rng[["text"]] <- hdrs[[j]]
      shd <- rng[["Shading"]]
      shd[["BackgroundPatternColor"]] <- RGB(unlist(bgColour$RGB))
      rng[["Font"]][["Color"]] <- RGB(unlist(fgColour$RGB))
      rng[["Font"]][["bold"]] <- TRUE
      next
    }
  i <- i + 1
  for (row in seq(numrow)) {
    if (pg < tbl[["Rows"]][[i]][["Range"]]$Information(wdActiveEndPageNumber)){
      pg = tbl[["Rows"]][[i]][["Range"]]$Information(wdActiveEndPageNumber)
      tbl[["Rows"]]$Add(tbl[["Rows"]][[i]])
      for (j in seq(numcol)) {
        rng <- tbl[["Rows"]][[i]][["Cells"]][[j]][["range"]]
        rng[["text"]] <- hdrs[[j]]
        shd <- rng[["Shading"]]
        shd[["BackgroundPatternColor"]] <- RGB(unlist(bgColour$RGB))
        rng[["Font"]][["Color"]] <- RGB(unlist(fgColour$RGB))
        rng[["Font"]][["bold"]] <- TRUE
        next
      }
      i <- i + 1
    }
    for (j in seq(numcol)) { 
      rng <- tbl[["Rows"]][[i]][["Cells"]][[j]][["range"]]
      if (j == 1) {
        num1 <- normVal(tclvalue(tkget(rows$Number1[[row]])))
      } else if (j == 2) {
        num1 <- normVal(tclvalue(tkget(rows$Number2[[row]])))
      } else {
        num1 <- normVal(tclvalue(tkcget(rows$`Multiplied Result`[[row]], "-text")))
      }
      rng[["text"]] <- restyle(2, num1)
      rng[["ParagraphFormat"]][["Alignment"]] <- wdAlignParagraphRight
    }
    i <- i + 1
  }
  if (pg < tbl[["Rows"]][[i]][["Range"]]$Information(wdActiveEndPageNumber)){
    pg = tbl[["Rows"]][[i]][["Range"]]$Information(wdActiveEndPageNumber)
    tbl[["Rows"]]$Add(tbl[["Rows"]][[i]])
    for (j in seq(numcol)) {
      rng <- tbl[["Rows"]][[i]][["Cells"]][[j]][["range"]]
      rng[["text"]] <- hdrs[[j]]
      shd <- rng[["Shading"]]
      shd[["BackgroundPatternColor"]] <- RGB(unlist(bgColour$RGB))
      rng[["Font"]][["Color"]] <- RGB(unlist(fgColour$RGB))
      rng[["Font"]][["bold"]] <- TRUE
      next
    }
    i <- i + 1
  }
  rng <- tbl[["Rows"]][[i]][["Cells"]]
  bdr <- rng$Borders(wdBorderLeft)
  bdr[["LineStyle"]] <- wdLineStyleNone
  bdr <- rng$Borders(wdBorderRight)
  bdr[["LineStyle"]] <- wdLineStyleNone 
  bdr <- rng$Borders(wdBorderVertical)
  bdr[["LineStyle"]] <- wdLineStyleNone
  i <- i + 1
  if (pg < tbl[["Rows"]][[i]][["Range"]]$Information(wdActiveEndPageNumber)){
    pg = tbl[["Rows"]][[i]][["Range"]]$Information(wdActiveEndPageNumber)
    tbl[["Rows"]]$Add(tbl[["Rows"]][[i]])
    for (j in seq(numcol)) {
      rng <- tbl[["Rows"]][[i]][["Cells"]][[j]][["range"]]
      rng[["text"]] <- hdrs[[j]]
      shd <- rng[["Shading"]]
      shd[["BackgroundPatternColor"]] <- RGB(unlist(bgColour$RGB))
      rng[["Font"]][["Color"]] <- RGB(unlist(fgColour$RGB))
      rng[["Font"]][["bold"]] <- TRUE
      next
    }
    i <- i + 1
  }
  for (j in seq(numcol)){
    rng <- tbl[["Rows"]][[i]][["Cells"]][[j]][["range"]]
    if (j == 1) {
      num1 <- tclvalue(tkcget(sums$Number1[[1]], "-text"))
    } else if (j == 2) {
      num1 <- tclvalue(tkcget(sums$Number2[[1]], "-text"))
    } else {
      num1 <- tclvalue(tkcget(sums$`Multiplied Result`[[1]], "-text"))
    }
    rng[["text"]] <- num1
    rng[["ParagraphFormat"]][["Alignment"]] <- wdAlignParagraphRight
  }
}
  
  toClipboard <- function() {
    clip <- paste0(c(paste0(hdrs[1:2], "\t"), paste0(hdrs[3], "\n")), collapse = "")
    for (row in seq(numrow)) {
      clip <- paste0(c(clip, restyle(2, normVal(tclvalue(tkget(rows$Number1[[row]])))), "\t"), collapse = "")
      clip <- paste0(c(clip, restyle(2, normVal(tclvalue(tkget(rows$Number2[[row]])))), "\t"), collapse = "")
      clip <- paste0(c(clip, restyle(2, normVal(tclvalue(tkcget(rows$`Multiplied Result`[[row]], "-text")))), "\n"), 
                     collapse = "")
    }
    clip <- paste0(c(clip, "\n"), collapse = "")
    clip <- paste0(c(clip, restyle(2, normVal(tclvalue(tkcget(sums$Number1[[1]], "-text")))), "\t"), collapse = "")
    clip <- paste0(c(clip, restyle(2, normVal(tclvalue(tkcget(sums$Number2[[1]], "-text")))), "\t"), collapse = "")
    clip <- paste0(c(clip, restyle(2, normVal(tclvalue(tkcget(sums$`Multiplied Result`[[1]], "-text")))), "\n"), 
                   collapse = "")
    utils::writeClipboard(clip)
  }
  
  updColours <- function() {
    bgColour$RGB$Red <<- as.numeric(tclvalue(bgred)); bgColour$RGB$Green <<- as.numeric(tclvalue(bggreen)); 
    bgColour$RGB$Blue <<- as.numeric(tclvalue(bgblue)); 
    bgColour$Hex <<- paste0("#", paste(as.hexmode(unlist(bgColour$RGB)), collapse = ""))
    fgColour$RGB$Red <<- as.numeric(tclvalue(fgred)); fgColour$RGB$Green <<- as.numeric(tclvalue(fggreen)); 
    fgColour$RGB$Blue <<- as.numeric(tclvalue(fgblue))
    fgColour$Hex <<- paste0("#", paste(as.hexmode(unlist(fgColour$RGB)), collapse = ""))
    tkconfigure(hdr$Number1[[1]], bg=bgColour$Hex, fg=fgColour$Hex)
    tkconfigure(hdr$Number2[[1]], bg=bgColour$Hex, fg=fgColour$Hex)
    tkconfigure(hdr$`Multiplied Result`[[1]], bg=bgColour$Hex, fg=fgColour$Hex)
  }

    self <- tktoplevel()
    tktitle(self) <- 'IPI Paul- C++ multiplyBy'
    numrow <- 1
    numcol <- 3
    formats <- as.data.frame(rbind(
      cbind("General", "%f"), 
      cbind("#,###,##0.00", "%.2f"), 
      cbind("#,###,##0", "%.0f")
      ), 
      stringsAsFactors = FALSE
      )
    names(formats) <- c("Type", "Format")
    hdrs <- c("Number1", "Number2", "Multiplied Result")  
    nms <- paste(sapply(seq(hdrs), function(x) cbind(hdrs[x], paste0("Env", x))))
    var1 <- tclVar("General")
    var2 <- tclVar("")
    colnames(formats) <- c("Type", "Format")
    bgColour <- list(RGB = list(Red=0, Green=0, Blue=108), Hex = paste0("#", paste(as.hexmode(c(0, 0, 108)), collapse = "")))
    fgColour <- list(RGB = list(Red=255, Green=255, Blue=255), Hex = paste0("#", paste(as.hexmode(c(255, 255, 255)), collapse = "")))
    bgred <- tclVar(bgColour$RGB$Red)
    bggreen <- tclVar(bgColour$RGB$Green)
    bgblue <- tclVar(bgColour$RGB$Blue)
    fgred <- tclVar(fgColour$RGB$Red)
    fggreen <- tclVar(fgColour$RGB$Green)
    fgblue <- tclVar(fgColour$RGB$Blue)
    main <- tkframe(self)
    frame_canvas <- tkframe(main)
    toolBar <- tkframe(self)
    canvas <- tkcanvas(frame_canvas, height=500, width=650)
    sbar <- tkscrollbar(frame_canvas, command = function(...) tkyview(canvas, ...))
    tkconfigure(canvas, yscrollcommand = function(...) tkset(sbar, ...))
    tkgrid(sbar, row=0, column = 1, sticky = "ns")
    tkgrid(canvas, row=0, column = 0, sticky = "nw")
    calc = tkframe(canvas, width=650, height=1000)
    tkcreate(canvas, "window", 0, 0, window=calc, anchor="nw")
    makeWidgets()
    tkpack(frame_canvas, expand = TRUE, fill = "both")
    tkgrid(main, row=0, columnspan=4)
    tkgrid(toolBar, row=1, columnspan=4)
    
    
    #rPath <- dirname(getActiveDocumentContext()$path)
    #dl <- file.path(rPath, '..', 'dll', 'multiplyBy x64.dll')
    dyn.load(dl)

    tkwait.window(self)
    dyn.unload(dl)
}
gui()
getLoadedDLLs()["multiplyBy x64"]

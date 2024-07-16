library(Rcompression)
library(XML)

if(FALSE) {
    f = "CoE - Masters Program Topics - Rankings by Student Demand_11.09.2023-1.pptx"
    z2 = pptxTables(f, 4:13)
    table(sapply(z2, class))
    table(sapply(z2, nrow))
    
}


DrawNS = "http://schemas.openxmlformats.org/drawingml/2006/main"

pptxTables =
function(file, page = integer(), ar = zipArchive(file))    
{
    if(length(page) == 0)
        page = grep("^ppt/slides/.*\\.xml$", names(ar), value = TRUE)
    else if(is.integer(page)) {
        rx = sprintf("^ppt/slides/slide(%s)\\.xml$", paste(page, collapse = "|"))
        page = grep(rx, names(ar), value = TRUE)
    } 

    ans = lapply(page, readSlideTable, ar)
    names(ans) = page
    ans
}

readSlideTable =
function(docName, ar)
{
    doc = xmlParse(ar[[docName]])
    tbls = xpathApply(doc, "//a:tbl", readPPTXTable, namespaces = c(a = DrawNS))

    if(length(tbls) == 1)
        tbls[[1]]
    else
        tbls
}



readPPTXTable =
function(tbl)    
{
    rows = xpathApply(tbl, ".//a:tr", readRow, namespaces = c(a = DrawNS))
    ans = as.data.frame(do.call(rbind, rows[-1]))
    names(ans) = rows[[1]]
    ans[] = lapply(ans, cvtCol)

    w = names(ans) == "" & (sapply(ans, function(x) all(is.na(x))) | sapply(ans, function(x) all(x == "")))

    as.data.frame( ans[!w] )
}

readRow =
function(tr)
{
    ans = xmlSApply(tr, xmlValue, trim = TRUE)
    ans[ ans == " "] = ""
    ans
}


# type.convert
cvtCol =
function(x)
{
    isPct = is.na(x) | x == "" | grepl("%$", x)
    if(all(isPct))
        x = gsub("%$", "", x)

    type.convert(x, as.is = TRUE)
}


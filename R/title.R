slideTitle =
    #
    # Simple-minded way to get text of title.  Finds the elements with the largest (font) size.
    # Doesn't handle styles, i.e., where the sz attribute is not explicit but inherited.
    #
function(doc)
{
    els = getNodeSet(doc, "//a:r", namespaces = c(a = DrawNS))
    if(length(els) == 0)
        return(NA)
    
    sz = sapply(els, function(x) xmlGetAttr(x[[ "rPr"]], "sz"))
    sz = as.integer(sz)

    w = (sz == max(sz))
    paste(sapply(els[w], xmlValue), collapse = " ")
}

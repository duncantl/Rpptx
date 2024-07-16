ptxSlides =
    #
    # Get the names of the slide documents within the zip archive
    # Optionally pare the documents.
    #
function(file, page = integer(), ar = zipArchive(file), parse = FALSE)    
{
    if(length(page) == 0)
        page = grep("^ppt/slides/.*\\.xml$", names(ar), value = TRUE)
    else if(is.integer(page)) {
        rx = sprintf("^ppt/slides/slide(%s)\\.xml$", paste(page, collapse = "|"))
        page = grep(rx, names(ar), value = TRUE)
    }

    if(parse)
        structure(lapply(page, function(x) xmlParse( ar [[ x ]] )),
                  names = page)
    else
        page
}

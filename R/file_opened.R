#' @title file_opened
#' @description Checks if a file is opened (helper function for plot2ppt)
#'
#' @param path A path to the file that has to be checked
#'
#' @return boolean value, TRUE if file is open, FALSE if closed
#'
#'
#' @export

file_opened <- function(path) {
  suppressWarnings(
    "try-error" %in% class(
      try(file(path,
               open = "w"),
          silent = TRUE
      )
    )
  )
}

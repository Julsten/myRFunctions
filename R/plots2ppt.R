#' @title plots2ppt
#' @description takes a ggplot object and writes it into a power point presentation
#'
#' @param plot a ggplot object
#' @param doc a rpptx object from the package officer, doc <- officer::read_pptx()
#' @param path A path to write the file
#'
#' @return a pptx file
#'
#'
#' @export
#' @importFrom officer add_slide
#' @importFrom officer ph_with
#' @importFrom rvg dml
#'
#' @examples
#' \dontrun{
#' library(ggplot2)
#' data(mpg)
#' carplot <- ggplot(data=mpg) + geom_boxplot(aes(x=manufacturer, y=displ))
#' doc <- officer::read_pptx()
#' plots2ppt(carplot, doc, path="my_presentation.pptx")
#' }




plots2ppt <- function(plot, doc, path){
  #doc needs to be specified with officer::read_pptx()
  if (file_opened(path)) stop("Error: PowerPoint Presentation needs to be closed!")
  if (!("rpptx" %in% class(doc))) stop("Error: doc needs to be of class rpptx (requires officer package)")
  if (!"ggplot" %in% class(plot)) stop("Error: plot needs to be of class ggplot")
  require(officer)
  require(rvg)
  doc <- add_slide(doc, 'Title and Content', 'Office Theme')
  vec_plt <- dml(ggobj=plot) # vector graphic
  doc <- ph_with(doc, vec_plt, ph_location_type(type="body"))
  print(doc, target = path)
}

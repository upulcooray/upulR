
#' Convert tableby object to a flextable
#'
#' @param tblby arsenal::tableby object
#' @param header_row
#' @param colwidths
#'
#' @return none

tableby_to_flextable <- function(tblby,
                                 header_row= "add span row title",
                                 colwidths= c(1,4),
                                 size= 11,
                                 family= "Arial"
                                 ){

  flextable::set_flextable_defaults(
    font.family = family ,
    font.size = size,
    table.layout = "autofit",
    line_spacing= 0.9)

  df<- tblby %>% summary(text=T) %>% as.data.frame()

  colnames(df)[1] <-  "Characteristics"


  indent_cols <- which(startsWith(df$Characteristics, "- "))


  df <- df %>% dplyr::mutate(Characteristics = stringr::str_remove_all(.data$Characteristics, "- "))



  df %>% flextable::flextable() %>%

    flextable::padding(i = indent_cols, j = 1, padding.left = 15, part = "body") %>%
    flextable::align(j = 1, align = "left", part = "all") %>%
    flextable::bold(~ !startsWith(Characteristics, " "), ~Characteristics) %>%
    flextable::add_header_row(values = c("", header_row),
                              colwidths = colwidths) %>%
    flextable::align(i = 1, align = "center", part = "header") %>%
    flextable::hline_top(border = officer::fp_border(width = 2), part = "header") %>%
    flextable::hline_bottom(border = officer::fp_border(width = 2), part = "body") %>%
    flextable::hline(i=1,j=1, border = officer::fp_border(width = 0), part = "header") %>%
    flextable::merge_v(j=1,part = "header") %>%
    flextable::set_header_labels(Characteristics = "Characteristics") %>%
    flextable::footnote(i = 2, j = 1,
                        value = flextable::as_paragraph("Mean (SD) for continuous variables; Frequency (%) for categorical variables"),
                        ref_symbols = c("1"),
                        part = "header")
}


#' Create descriptive tables
#'
#' @param df a dataframe
#' @param headvar The variable name that you want your descriptive statistics stratified by (outcome or exposure variable).
#' @param headvar_levels labels to be used for each level of head variable
#' (eg: c("Yes"= "1","No"= "0", "Missing"= NA))
#' @param rowvars a vector explanatory variables (in the order you want them to appear in the table)
#' example: c("age", "sex", "income")
#' @param levels_order order of columns from left to right (eg: c("No", "Yes"))
#' @param rowvar_labels Labels for explanatory variables (in correct order)
#' example: c("Age (yr)", "Gender", "Monthly income")
#' @param headvar_na_level How to name NA level of head variable (Eg: "Missing")
#' @param file_name Name of the .docx file contain the table
#' @param header Spanning header for the head variable columns
#' @param digits_num Number of digits for numerical values
#' @param do_test TRUE if you want to perform statistical analyses (eg: "chi", "anova", "kwt", ..etc)
#' can specify tets with additional arguments (?arsenal::tableby.control)
#' @param col_total Total for columns if TRUE
#' @param numvar_stats Stats to display for numerical variables
#' @param catvar_stats Stats to display for categorical variables
#' @param stats_labels List of labels for respective stats
#' @param ... Takes any argument for the arsenal::tableby.control() function
#'
#' @return A .docx document with the table
#' @export
#'
#' @examples  df <- mtcars
#'
#' create_table1(df,headvar= "am",
#'                  headvar_levels= c("Yes"= "1","No"= "0") ,
#'                  rowvars= c("wt","gear","carb"),
#'                  levels_order= c("Yes","No"),
#'                  rowvar_labels= c("Weight","Gear","Number of carbs"))
create_table1 <- function(df,
                          headvar,
                          rowvars,
                          headvar_levels=NULL,
                          levels_order= NULL,
                          rowvar_labels=NULL,
                          headvar_na_level= "Missing",
                          file_name= "table",
                          header= "add span row title",
                          digits_num = 1L,
                          do_test= FALSE,
                          col_total=FALSE,
                          numvar_stats= c("meansd","Nmiss"),
                          catvar_stats= c("countrowpct"),
                          stats_labels=list(Nmiss= "(Missing)"),
                          font_size=11,
                          font_family= "Arial",
                          # numeric.test="kwt",
                          # cat.test="chisq",
                          # digits.count = 0L,
                          # digits.pct = 1L,
                          ... ){

  library(magrittr)


  mycontrols  <- arsenal::tableby.control(test=do_test,
                                          total=col_total,
                                          numeric.stats= numvar_stats,
                                          cat.stats= catvar_stats,
                                          stats.labels=stats_labels,
                                          digits = digits_num,
                                          ... )

  num_vars <- df %>% stats::na.omit() %>%
    dplyr::select_if(function(x) length(unique(x))> 6) %>% colnames()

  formula <- glueformula::gf({headvar}~ {rowvars})

  var <- rlang::enquo(headvar)

  df <- dplyr::mutate(df,dplyr::across(-dplyr::all_of(num_vars),
                                       ~as.factor(.)))
  if (!is.null(headvar_levels)){

    df[[headvar]] <- forcats::fct_recode(df[[headvar]], !!!headvar_levels)
  }

  if (!is.null(levels_order)){

    df[[headvar]] <- forcats::fct_relevel(df[[headvar]],  levels_order)
  }

  df[[headvar]] <- forcats::fct_explicit_na(df[[headvar]], headvar_na_level)


  n_levels <- length(unique(df[[headvar]]))

  if (!is.null(rowvar_labels)){
  arsenal::labels(df)<- setNames(rowvar_labels, rowvars)
  }

  tab1_df<- df%>%

    arsenal::tableby(formula = formula,
                     data=.,
                     control = mycontrols)

  tab <- tableby_to_flextable(tab1_df ,
                              header_row= header,
                              colwidths = c(1,n_levels),
                              size = font_size,
                              family= font_family
                              )

  doc <- officer::read_docx()
  doc <- flextable::body_add_flextable(doc, value = tab)
  fileout <- tempfile(fileext = ".docx")
  fileout <- glue::glue({file_name},".docx") # write in your working directory

  print(doc, target = fileout)


}



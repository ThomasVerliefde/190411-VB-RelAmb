library(magrittr)

combinat::permn(c(0,1,2,3)) %>%
  unique %T>%
  {set.seed(20181001)} %>%
  sample(
    size = 500,
    replace = T
  ) %>% unlist %>% paste(collapse = " ") %>%
  write(
    "BlockRandomisation.txt"
    )
  
  

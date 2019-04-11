library(magrittr)

list(0:1) %>%
  rep(3) %>%
  expand.grid %>%
  tidyr::unite(sep="") %>%
  unlist %>%
  combinat::permn(.) %T>%
  {set.seed(20190411)} %>%
  sample(
    size = 500,
    replace = T
  ) %>% unlist %>% paste(collapse = " ") %>%
  write(
    "BlockRandomisation.txt"
  )

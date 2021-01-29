CIMR 001 Import: find\_obvious\_dups
====================================

Custom Citavi macros that specifically imports the results created by
the `find_obvious_dups()` function of the [`CitaviR`
package](https://github.com/SchmidtPaul/CitaviR#citavir-).

![](CIMR%20001.gif)

The excel file shown above can be created as:

    library(CitaviR)

    ## 
    ## Attaching package: 'CitaviR'

    ## The following object is masked _by_ '.GlobalEnv':
    ## 
    ##     %not_in%

    path   <- example_xlsx('3dupsin5refs.xlsx') # replace with path to your xlsx file

    read_Citavi_xlsx(path) %>% 
      find_obvious_dups() %>% 
      write_Citavi_xlsx(path)

require(xml2) #e.g. url escape
require(httr) #eg get
require(openxlsx)
ui <- fluidPage(
    titlePanel("PermuSearch v2"),
    mainPanel(
      strong("To permute your searches in Scopus:"),
      br(),
      br(),
      strong("1. Specify your input in "),
      strong(tags$a(href="https://github.com/spzwanenburg/PermuSearch/raw/main/v2/template v2.xlsx", 
             "the template")),
      strong("and save it."),
      br(),
      br(),
      fileInput("file1", "2. Upload it:",
                accept = c("text/xlsx",
                           "text/xls,
                       .xlsx,
                       .xls")),
      textInput("key","3. Enter your Scopus API key:", ""),
      HTML("<i>Request a Scopus API key 
           <a href=\"https://dev.elsevier.com/\">here</a>.<br><br>The ECIS workshop key is: </i>0f04e8f69bd2e20fecf037a26d3172e8"),
      br(),
      br(),
      strong("4. Start the automated search: "),
      downloadButton("dl", "Start"),
      br(),
      br(),
      strong("5. A download should start when the search is complete."),
      hr(),
      HTML("<small><p>PermuSearch is an open source project licensed under the MIT license, documented on 
      <a href=\"https://github.com/spzwanenburg/PermuSearch/tree/main/\">GitHub</a>.<br>
           Cite the tool as: <i>Zwanenburg, S.P. (2023) PermuSearch V2.0 [Online software]. permusearch.com</i><br>
           Contact us at <i>info / at / permusearch.com</i></p></small>")
  )
)


server <- function(input, output) {
  whichurl <- function(here,ftopic,xterm,yterm,article,when,api,key){
    if(api){
      return(paste0("https://api.elsevier.com/content/search/scopus?query=",
                    here,ftopic,xterm,yterm,article,when,key))
    }
    else{
      return(paste0("https://www.scopus.com/results/results.uri?sort=plf-f&src=s&sot=a&s=",
                    here,ftopic,xterm,yterm,article,when))
    }
  }
  howmany <- function(searchurl) {
    geturl<-GET(searchurl, accept("application/json"))
    return(content(geturl)[["search-results"]][["opensearch:totalResults"]])
  }
  time <- reactive({
    inFile <- input$file1
    if (is.null(inFile)) {
      return(NULL)
    }
    wb<-loadWorkbook(inFile$datapath, isUnzipped = FALSE)
    
  })
  data <- reactive({
    
    inFile <- input$file1
    if (is.null(inFile)) {
      return(NULL)
    }
    key<-paste0("&apiKey=",input$key)
    wb<-loadWorkbook(inFile$datapath, isUnzipped = FALSE)
    searchdf<-read.xlsx(wb,"search",colNames=TRUE,cols=2:6)
    colnames(searchdf)<-c("x","y","c1","c2","c3")
    if(searchdf[1,2]=="not defined"){searchdf[2:3,2]<-" "}
    doctype<-""
    after<-""
    before<-""
    cvenue<-FALSE
    ctopic<-FALSE
    for(i in 3:5){
      if(searchdf[1,i]=="document type"){doctype<-gsub(" ","%20",searchdf[3,i])}
      if(searchdf[1,i]=="after year"){after<-gsub(" ","%20",searchdf[3,i])}
      if(searchdf[1,i]=="before year"){before<-gsub(" ","%20",searchdf[3,i])}
      if(searchdf[1,i]=="topic"){
        ctopic=TRUE
        topic<-searchdf[3:sum(!is.na(searchdf[,i])),i]
      }
      if(searchdf[1,i]=="venue"){
        cvenue=TRUE
        venues<-searchdf[3:sum(!is.na(searchdf[,i])),i]
      }
    }
    
    #set 'then' if constrained (and not x or y)
    then<-""
    if(searchdf[1,1]!="year"&&searchdf[1,2]!="year"){
      if(nchar(before)>0&&nchar(after)>0){
        then<-paste0(before,"%20AND%20",after)
      }
      else{then<-paste0(before,after)}
    } #then is set (empty if either dim x or dim y are years, in which case we specify it during an iteration)
    # set 'here' if constrained
    here<-""
    if(searchdf[1,1]!="venue"&&searchdf[1,2]!="venue"&&cvenue==TRUE){
      here<-url_escape(paste(venues,collapse=" OR "))
    } #here is set to list of venues if constrained.
    ftopic<-""
    if(searchdf[1,1]!="topic"&&searchdf[1,2]!="topic"&&ctopic==TRUE){
      ftopic<-url_escape(paste(topic,collapse=" OR "))
    }
    #execute set up; results is df in r. 
    results<-data.frame(matrix(ncol = sum(!is.na(searchdf$x))-1,nrow = sum(!is.na(searchdf$y))-1))
    colnames(results)<-c("unconstrained",searchdf[3:sum(!is.na(searchdf$x)),"x"])
    rownames(results)<-c("unconstrained",searchdf[3:sum(!is.na(searchdf$y)),"y"])
    writeData(wb, sheet = "results", x = results, startCol = 1, startRow = 1, colNames=TRUE, rowNames=TRUE)
    resultvalue<-"url to insert here"
    names(resultvalue)<-999
    class(resultvalue)<-"hyperlink"
    k<-1
    totalqueries<-ncol(results)*nrow(results)
    withProgress(message = 'Running query', value = 0, {
      for(i in 0:(ncol(results)-1)){
        for(j in 0:(nrow(results)-1)){
          ifelse(i==0,xvalue<-"",xvalue<-url_escape(searchdf[i+2,"x"])) 
          ifelse(j==0,yvalue<-"",yvalue<-url_escape(searchdf[j+2,"y"]))
          if(i==0&&j==0&&sum(searchdf[1,3:5]=="not defined")==3){next}
          #if x or y is year, we must change the order of terms because scopus is picky on order.
          if(searchdf[1,1]=="year"){
            searchurl<-whichurl(here,ftopic,yvalue,doctype,xvalue,then,TRUE,key)
            resultvalue<-whichurl(here,ftopic,yvalue,doctype,xvalue,then,FALSE,key)
          } else if (searchdf[1,2]=="year"){
            searchurl<-whichurl(here,ftopic,xvalue,doctype,yvalue,then,TRUE,key)
            resultvalue<-whichurl(here,ftopic,xvalue,doctype,yvalue,then,FALSE,key)
          } else {
            searchurl<-whichurl(here,ftopic,xvalue,yvalue,doctype,then,TRUE,key)
            resultvalue<-whichurl(here,ftopic,xvalue,yvalue,doctype,then,FALSE,key) 
          }
          numResults<-as.numeric(howmany(searchurl))
          names(resultvalue)<-as.numeric(numResults)
          class(resultvalue)<-"hyperlink"
          results[j+1,i+1]<-numResults
          writeData(wb, sheet = "results", x = resultvalue, startCol = 2+i, startRow = 2+j)
          print(paste0("x: ",xvalue,"; y: ",yvalue,"; Number of results: ",numResults))
          incProgress(1/(ncol(results)*nrow(results)), detail = paste(k, "/", totalqueries, "(cell ", i+1, " x ", j+1,")"))
          k<-k+1
        }
      }
    })
    return(wb)
  })
  myname <- reactive({
    # Test if file is selected
    if (!is.null(input$file1$datapath)) {
      # Extract file name (additionally remove file extension using sub)
      return(sub(".xlsx$", "", basename(input$file1$name)))
    } else {
      return(NULL)
    }
  })
  output$dl <- downloadHandler(filename = function() {
      myfilename <- myname()
      paste0(myfilename,"-results.xlsx")
    },
    content = function(file) {
      wb1 <- data()
      saveWorkbook(wb1, file = file)
    }
  )
}
shinyApp(ui = ui, server = server)

require(xml2) #e.g. url escape
require(httr) #eg get
require(openxlsx)
ui <- fluidPage(
    titlePanel("PermuSearch"),
    
    
    mainPanel(
      strong("To permutate your searches in Scopus:"),
      br(),
      br(),
      strong("1. Specify your input in "),
      strong(tags$a(href="https://github.com/spzwanenburg/PermuSearch/raw/main/new%20template.xlsx", 
             "the template")),
      strong("and save it."),
      br(),
      br(),
      fileInput("file1", "2. Upload your XLSX input:",
                accept = c("text/xlsx",
                           "text/xls,
                       .xlsx,
                       .xls")),
      textInput("key","3. Enter your scopus API key:", "650ab1cd6598f118ea92f90996b1ac88"),
      em("(The above API key is provided for the ACIS workshop only, permitting fair use.)"),
      br(),
      br(),
      strong("4. Start the automated search: "),
      downloadButton("dl", "Start"),
      br(),
      br(),
      strong("5. A download should start when the search is complete."),
      hr(),
      HTML("<small><p>PermuSearch is an open source project licensed under the MIT license, documented on 
      <a href=\"https://github.com/spzwanenburg/PermuSearch/\">GitHub</a>.<br>
           Cite the tool as: <i>Zwanenburg, S.P. (2023) PermuSearch V1.0 [Online software]. permusearch.com</i><br>
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
    topic.list1<-read.xlsx(wb,"topic.list1")
    topic.list2<-read.xlsx(wb,"topic.list2")
    years<-read.xlsx(wb,"years")
    journals<-read.xlsx(wb,"journals")
    journals<-journals[journals$`included?`=="yes",]
    parameters<-read.xlsx(wb,"parameters")
    #set values
    dimx<-parameters[which(parameters$setting=="x dimension"),"value"]
    dimy<-parameters[which(parameters$setting=="y dimension"),"value"]
    documents<-parameters[which(parameters$setting=="documents"),"value"]
    #setting constraints
    if (documents=="articles"){
      article<-"DOCTYPE(AR)"
    } else if (documents=="articles or conference papers"){
      article<-"(DOCTYPE(AR)%20OR%20DOCTYPE(CP))"
    } else if (documents=="all document types"){
      article<-""
    } else {
      abort(message = "document setting value is not recognized")
    }
    after<-parameters[which(parameters$setting=="after year"),"value"]
    if(dimx=="years"|dimy=="years"){
      then<-""
    } else if (!is.na(after)){
      then<-paste0("PUBYEAR%20AFT%20",as.character(after))
    } else{ then<-"PUBYEAR%20AFT%201992" }
    if(dimx!="journals"&&dimy!="journals"){
      if(parameters[which(parameters$setting=="journals"),"value"]=="unconstrained"){
        here<-""
      } else {here<-url_escape(paste(journals[1:nrow(journals),"search.term"],collapse=" OR "))}
    } else{ here<-""}
    if(dimx!="topic.list1"&&dimy!="topic.list1"&&dimx!="topic.list2"&&dimy!="topic.list2"){
      ftopic<-url_escape(parameters[which(parameters$setting=="fixed topic"),"value"])
    } else{ ftopic<-"" }
    #setting constraints is done.
    #execute set up; results is df in r. 
    results<-data.frame(matrix(ncol = 1+nrow(get(as.name(dimx))),nrow = 1+nrow(get(as.name(dimy)))))
    setNames(data.frame(matrix(ncol = 3, nrow = 0)), c("name", "age", "gender"))
    colnames(results)<-c("unconstrained",get(as.name(dimx))$label)
    rownames(results)<-c("unconstrained",get(as.name(dimy))$label)
    writeData(wb, sheet = "results", x = results, startCol = 1, startRow = 1, colNames=TRUE, rowNames=TRUE)
    resultvalue<-"url to insert here"
    names(resultvalue)<-999
    class(resultvalue)<-"hyperlink"
    begintime<-Sys.time()
    i<-1
    totalqueries<-ncol(results)*nrow(results)
    withProgress(message = 'Running query', value = 0, {
    for(x in 0:(ncol(results)-1)){
      for(y in 0:(nrow(results)-1)){
        if(x==0){
          xvalue<-""
        } else { 
          xvalue<-url_escape(get(as.name(dimx))[x,"search.term"]) 
        }
        if(y==0){
          yvalue<-""
        } else { 
          yvalue<-url_escape(get(as.name(dimy))[y,"search.term"]) 
        }
        #if dimx or dimy is years, we must change the order of terms because scopus is picky on order.
        if(dimx=="years"){
          searchurl<-whichurl(here,ftopic,yvalue,article,xvalue,then,TRUE,key)
          resultvalue<-whichurl(here,ftopic,yvalue,article,xvalue,then,FALSE,key)
        } else if (dimy=="years"){
          searchurl<-whichurl(here,ftopic,xvalue,article,yvalue,then,TRUE,key)
          resultvalue<-whichurl(here,ftopic,xvalue,article,yvalue,then,FALSE,key)
        } else {
          searchurl<-whichurl(here,ftopic,xvalue,yvalue,article,then,TRUE,key)
          resultvalue<-whichurl(here,ftopic,xvalue,yvalue,article,then,FALSE,key) 
        }
        numResults<-as.numeric(howmany(searchurl))
        names(resultvalue)<-as.numeric(numResults)
        class(resultvalue)<-"hyperlink"
        results[y+1,x+1]<-numResults
        writeData(wb, sheet = "results", x = resultvalue, startCol = 2+x, startRow = 2+y)
        print(paste0("x: ",xvalue,"; y: ",yvalue,"; Number of results: ",numResults))
        i<i+1
        incProgress(1/(ncol(results)*nrow(results)), detail = paste(i, "/", totalqueries, "(cell ", x, " x ", y,")"))
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

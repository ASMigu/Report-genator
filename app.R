
library(dplyr)
library(readxl)
library(magrittr)
library(rhandsontable)
library(shiny)
library(shinydashboard)
options(shiny.maxRequestSize = 20*1024^2) 

# 1.header
header <- dashboardHeader(title = "Tool Development Report Generator",
                          titleWidth = 280)


# 2.sidebar
sidebar <- dashboardSidebar(width = 280, 
                            sidebarUserPanel("吳三吉 (Ali Wu) - PDST",
                                             subtitle = "ali.wu@shl-medical.com",
                                             image = "https://img.sportsv.net/img/photo/image/6/117636/aspect-lHUN9NETyu-227xauto.jpg"),
                            sidebarMenu(
                              id = "sidebar",
                              menuItem(text = "Benchmaking", tabName = "bm", icon = icon("database"), badgeLabel = "NEW", badgeColor = "red"
                                       #menuSubItem(text = "Upload Data", tabName = "upload_bm", icon = icon("file", lib = "glyphicon")),
                                       #menuSubItem(text = "Fill Data Information", tabName = "fill_bm")
                              ),
                              menuItem(text = "Verification", tabName = "vr", icon = icon("coins")
                                       #menuSubItem(text = "Upload Data", tabName = "upload_vr", icon = icon("file", lib = "glyphicon")),
                                       #menuSubItem(text = "Fill Data Information", tabName = "fill_vr")
                              ),
                              menuItem(text = "Little Game", tabName = "ot", icon = icon("binoculars"))
                            )
)

#success, info, warning, danger
# 3.body
body <- dashboardBody(
  tabItems(
    tabItem("bm",
            #
            fluidRow(
              box(
                title = "Upload data", width = 6, height = 300, solidHeader = TRUE, status = "primary", #background = "olive",
                fileInput(inputId = "bm_inp", label = h1("Raw dataset"), 
                          accept = ".xlsx", width = "300px",
                          buttonLabel = list(icon("cloud-upload", lib = "glyphicon")),
                          placeholder = "Please choose a .xlsx file"),
              ),
              box(
                title = "Data analysis", width = 6, height = 300, solidHeader = TRUE, status = "primary", #background = "olive",
                br(),
                
                actionButton(inputId = "update_server_bm", label = "Upload Data to Server", 
                             icon = icon("arrow-alt-circle-up"), class = "btn-primary btn-lg"),
                hr(),
                
                uiOutput("report_ui_bm")
              )
            ),
            fluidRow(
              box(width = 12, status = "primary",
                  rHandsontableOutput("table_bm")
              )
            ), # Test block
            
            fluidRow(
              column(12,
                     hr(),
                     hr(),
                     hr(),
                     tableOutput("tbl1"),
                     tableOutput("tbl2"),
                     tableOutput("tbl3"),
                     textOutput("tex1"),
                     textOutput("tex2"),
              )
            )#
    ),
    
    #
    tabItem("vr",
            
            #
            fluidRow(
              box(title = "Upload data", width = 6, height = 400, solidHeader = TRUE, status = "warning",
                  fileInput(inputId = "test_vr_inp", label = h3("Raw (testing group) dataset"), 
                            accept = ".xlsx", width = "300px",
                            buttonLabel = list(icon("cloud-upload", lib = "glyphicon")),
                            placeholder = "Please choose a .xlsx file"),
                  hr(),
                  fileInput(inputId = "ref_vr_inp", label = h3("Reference group dataset"), 
                            accept = ".xlsx", width = "300px",
                            buttonLabel = list(icon("cloud-upload", lib = "glyphicon")),
                            placeholder = "Please choose a .xlsx file")
              ),
              box(title = "Data analysis", width = 6, height = 400, solidHeader = TRUE, status = "warning", 
                  br(),
                  
                  actionButton(inputId = "update_server_vr", label = "Upload Data to Server", 
                               icon = icon("arrow-alt-circle-up"), class = "btn-warning btn-lg"),
                  hr(),
                  
                  uiOutput("report_ui_vr")
                  
              )
            ),
            
            fluidRow(
              box(width = 12, status = "warning",
                  checkboxInput(inputId = "eqi_check", label = h5("Equivariance Study Request (ticking the box stands for yes)"),
                                value = FALSE, width = "650px"),
                  rHandsontableOutput("table")
              )
            )
            
    ),
    
    tabItem("ot",
            fluidRow(
              column(1, offset = 1, 
                     actionButton(inputId = "play_vr", label = h4("PLAY"), 
                                  icon = icon("exclamation"))
              ),
              column(4,
                     valueBoxOutput("count_vr"))
            ),
            fluidRow(
              column(3, offset = 1,
                     imageOutput("left_vr")
                     
              ),
              column(3, offset = 1,
                     imageOutput("center_vr")
                     
              ),
              column(3, offset = 1,
                     imageOutput("right_vr")
                     
              )
            ),
            fluidRow(
              column(
                12, offset = 1,
                imageOutput("upper_vr")
              )
            )#
    )
    
  )
)



# 4.server
server <- function(input, output, session){
  
  gloab <- reactiveValues()
  gloab$stop_words <- c("Cover", "NG List", "Dim List", "Note", "Note ", "Note_", " Note", " Note ",
                        paste("Note", 1:20, sep = ""), paste("Note ", 1:20, sep = ""),
                        paste("Note_", 1:20, sep = ""), paste("Note", "(",1:20,")", sep = ""),
                        paste("Note", " (",1:20,")", sep = ""), paste("Note", "  (",1:20,")", sep = ""),
                        paste("Note", "(",1:20,") ", sep = ""), paste("Note(", 1:20, ")",sep = ""))
  
  ###############################################
  ##################### B M #####################
  ###############################################
  
  # Input files
  raw_file_bm <- reactive({
    input$bm_inp
    req(input$bm_inp)
  })
  
  
  # obtain excel sheet names
  raw_file_shtname_bm <- reactive({
    req(raw_file_bm())
    
    readxl::excel_sheets(raw_file_bm()$datapath) %>% 
      setdiff(gloab$stop_words)
    
  })
  
  # Clean data
  
  ## extract data from input files (that are still messy)
  raw_data_bm <- reactive({
    req(raw_file_bm())
    
    temp_data <- data.frame()
    temp_1 <- c("")
    temp_2 <- c("")
    raw_file_shtname_bm <- raw_file_shtname_bm() %>% as.character
    for(i in raw_file_shtname_bm){
      temp_1 <- i
      temp_2 <- read_excel(raw_file_bm()$datapath, sheet = i, skip = 16)
      temp_2 <- merge(temp_2, temp_1)
      temp_data <- rbind(temp_data, temp_2)
    }
    
    temp_data <- temp_data[-which(temp_data["Sample"] == "Sample"), 
                           -c(4, 6, 9, 11, 12, 13, 14)]
    
    #temp_data <- temp_data[!is.na(temp_data["Value"]), ]
    
    names(temp_data) <- c("Dimension", "Sample", "Cavity", "Specification", 
                          "Upper_tolerance", "Lower_tolerance", "Value")
    
    temp_data[, c(1:3)] <- temp_data[, c(1:3)] %>%
      dplyr::mutate_if(is.character, as.factor)
    temp_data[, c(4:7)] <- temp_data[, c(4:7)] %>%
      dplyr::mutate_if(is.character, as.numeric)
    
    temp_data <- temp_data[!is.na(temp_data["Value"]), ]
    temp_data
    
  })
  
  # dimension information table
  
  dim_infomation_bm <- reactive({
    dim_info <- unique(raw_data_bm()[, c(1, 4, 5, 6)])
    dim_info[, c(2:4)] <- dim_info[, c(2:4)] %>% 
      dplyr::mutate_if(is.character, as.numeric)
    
    info <- readxl::read_excel(raw_file_bm()$datapath, sheet = raw_file_shtname_bm()[2])
    dim_info[,c(5:9)] <- data.frame(LSL = dim_info[, 2] - dim_info[, 4],
                                    USL = dim_info[, 2] + dim_info[, 3],
                                    Compensation = 0,
                                    Type = factor("n", levels = c("c", "p", "n"), labels = c("Critical", "Process", "Non-critical")),
                                    Toolset = as.character(substring(info[8, 4], 1, 4))
    )
    
    dim_info <- dim_info[c(9, 1, 7, 8, 3, 4, 5, 2, 6)]
  })
  
  output$table_bm <- renderRHandsontable({
    
    rhandsontable::rhandsontable(dim_infomation_bm(), digits = 6, rowHeaders = NULL)
  })
  
  
  # Upload data to server and generate download button
  observeEvent(input$update_server_bm, {
    
    swd1 <- c("/home/pdsaadmin/dataset/BM") 
    setwd(swd1)
    path_ <- paste0(swd1, "/", substring(raw_file_bm()$name, 1, 9))
    if(!file.exists(path_)) dir.create(path_)
    f <- list.files(swd1, include.dirs = F, full.names = T, recursive = T)
    #d <- list.dirs(swd, full.names = T, recursive = T)
    file.remove(f)
    #unlink(d[2:length(d)], recursive = T, force = F)
    
    if(TRUE){
      openxlsx::write.xlsx(readxl::read_excel(raw_file_bm()$datapath,
                                              sheet = raw_file_shtname_bm()[2])[1:16, ],
                           paste0(path_, "/", "Do not open.xlsx"))
      
      openxlsx::write.xlsx(hot_to_r(input$table_bm),
                           paste0(path_, "/", "Dim information.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(raw_data_bm(),
                           paste0(path_, "/", "Cleaned data copy tool.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(raw_file_bm()$datapath,
                           paste0(swd1, "/", raw_file_bm()$name),
                           row.names = FALSE)
    } else {
      return(999)
    }
    
    
    output$report_ui_bm <- renderUI({
      
      downloadButton("gen_bm", "Generate BM Report  ",
                     icon = icon("arrow-alt-circle-down"), 
                     class = "btn-primary btn-lg")
    })
  })
  
  # Analysis BM report and download the report  
  output$gen_bm <- downloadHandler(
    
    filename = function() {
      "_Benchmarking_ST_Rev.1.0.docx"
    },
    content = function(file) {
      
      src <- normalizePath(c('/home/pdsaadmin/dataset/BM.Rmd', '/home/pdsaadmin/dataset/template.docx')) # SEE HERE
      
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      file.copy(src, c('BM.Rmd', 'template.docx'), overwrite = TRUE) # SEE HERE
      
      #Sys.setenv(RSTUDIO_PANDOC="/usr/lib/rstudio-server/bin/pandoc")
      out <- rmarkdown::render('BM.Rmd',
                               envir = new.env(parent = globalenv())
                               
      )
      
      file.rename(out, file)
    }
    
  )
  
  # ### TEST BLOCK ↓#####
  # output$tex1 <- renderText({
  #   Sys.getenv("RSTUDIO_PANDOC")
  #})
  # output$tex2 <- renderText({
  #   raw_data_bm() %>% str
  # })
  # 
  # # 
  # output$tbl1 <- renderTable({
  #   dim_infomation_bm()
  # })
  # 
  # 
  # output$tbl2 <- renderTable({
  #   dim_infomation_vr()
  # })
  # ### TEST BLOCK ↑#####
  
  
  ###############################################
  ##################### V R #####################
  ###############################################
  
  # Input files
  raw_file_vr <- reactive({
    input$test_vr_inp
    req(input$test_vr_inp)
  })
  
  ref_file_vr <- reactive({
    input$ref_vr_inp
    req(input$ref_vr_inp)
  })
  
  
  # obtain excel sheet names
  raw_file_shtname_vr <- reactive({
    req(raw_file_vr())
    
    readxl::excel_sheets(raw_file_vr()$datapath) %>% 
      setdiff(gloab$stop_words)
    
  })
  
  
  ref_file_shtname_vr <- reactive({
    req(ref_file_vr())
    
    readxl::excel_sheets(ref_file_vr()$datapath) %>%
      setdiff(gloab$stop_words)
    
  })
  
  
  
  # Clean data
  
  ## extract data from input files (that are still messy)
  raw_data_vr <- reactive({
    req(raw_file_vr())
    
    temp_data <- data.frame()
    temp_1 <- c("")
    temp_2 <- c("")
    raw_file_shtname_vr <- raw_file_shtname_vr() %>% as.character
    for(i in raw_file_shtname_vr){
      temp_1 <- i
      temp_2 <- read_excel(raw_file_vr()$datapath, sheet = i, skip = 16)
      temp_2 <- merge(temp_2, temp_1)
      temp_data <- rbind(temp_data, temp_2)
    }
    
    temp_data[-which(temp_data["Sample"] == "Sample"), -c(4, 6, 9, 11, 12, 13)]
    
  })
  
  
  ref_data_vr <- reactive({
    req(ref_file_vr())
    
    temp_data <- data.frame()
    temp_1 <- c("")
    temp_2 <- c("")
    #ref_file_shtname_vr <- ref_file_shtname_vr()
    for(i in ref_file_shtname_vr()){
      temp_1 <- i
      temp_2 <- read_excel(ref_file_vr()$datapath, sheet = i, skip = 16)
      temp_2 <- merge(temp_2, temp_1)
      temp_data <- rbind(temp_data, temp_2)
    }
    
    temp_data[-which(temp_data["Sample"] == "Sample"), -c(4, 6, 9, 11, 12, 13)]
    
    
  })
  
  
  
  ## find difference dimensions b/t two files
  remove <- reactive({
    
    req(ref_data_vr())
    (unique(ref_data_vr()["Dim"])[[1]])%>%
      setdiff(unique(raw_data_vr()["Dim"])[[1]])
    
  })
  
  
  ref_data_vr_1 <- reactive({
    req(ref_file_vr())
    
    if(remove() %>% length == 0){
      ref_data_vr()
    } else{
      ref_data_vr()[-which(ref_data_vr()["Dim"] == remove()), ]
    }
    
    # ref_data_vr()[-which(ref_data_vr()["Dim"] == remove()), ]
  })
  
  
  # Without reference group dataser (1 input file)
  raw_data_vr_1 <- reactive({
    
    ## rename column and define data type
    cdt <- function(x){
      if(x == "P"){
        "High"
      }  else {
        if(x == "M") {
          "Low"
        } else {
          "CPT"
        }
      }
    }
    
    req(raw_file_vr())
    
    temp_data <- raw_data_vr()
    judge <- stringr::str_extract(temp_data[, 8], "[PNM]")
    temp_data[,8] <- sapply(judge, cdt)
    names(temp_data) <- c("Dimension", "Sample", "Cavity", "Specification", "Upper_tolerance",
                          "Lower_tolerance", "Value", "Condition")
    temp_data[, c(1:3, 8)] <- lapply(temp_data[, c(1:3, 8)], as.factor) 
    
    temp_data[, c(4:7)] <- lapply(temp_data[, c(4:7)], as.numeric) 
    temp_data[!is.na(temp_data$Value), ]
    
  })
  
  ref_data_vr_2 <- reactive({
    
    ## rename column and define data type
    cdt <- function(x){
      if(x == "P"){
        "High"
      }  else {
        if(x == "M") {
          "Low"
        } else {
          "CPT"
        }
      }
    }
    
    req(ref_file_vr())
    
    temp_data <- ref_data_vr_1()
    judge <- stringr::str_extract(temp_data[, 8], "[PNM]")
    temp_data[,8] <- sapply(judge, cdt)
    names(temp_data) <- c("Dimension", "Sample", "Cavity", "Specification", "Upper_tolerance",
                          "Lower_tolerance", "Value", "Condition")
    temp_data[, c(1:3, 8)] <- lapply(temp_data[, c(1:3, 8)], as.factor) 
    temp_data[, c(4:7)] <- lapply(temp_data[, c(4:7)], as.numeric)
    temp_data[!is.na(temp_data$Value), ]
    
  })
  
  
  
  # Equivalence data set (2 input files)
  equi_data <- reactive({
    
    req(ref_file_vr())
    temp1 <- raw_data_vr_1()
    temp2 <- ref_data_vr_2()
    
    temp1["Toolset"] <- "Copy tool"
    temp2["Toolset"] <- "Ref. tool"
    rbind(temp1, temp2)
    
    
  })
  
  
  # dimension information table
  
  dim_infomation_vr <- reactive({
    dim_info <- unique(raw_data_vr_1()[,c(1, 4, 5, 6)])
    dim_info[, c(2:4)] <- lapply(dim_info[,c(2:4)], as.numeric) 
    dim_info[, c(5:9)] <- data.frame(LSL = dim_info[, 2] - dim_info[, 4],
                                     USL = dim_info[, 2] + dim_info[, 3],
                                     Compensation = 0,
                                     Type =factor("n", levels = c("c", "p", "n"), labels = c("Critical", "Process", "Non-critical")),
                                     Toolset = readxl::read_excel(raw_file_vr()$datapath, 
                                                                  sheet = raw_file_shtname_vr()[2])[8, 4] %>% 
                                       substring(1, 4) %>% as.character
    )
    dim_info[c(9, 1, 7, 8, 3, 4, 5, 2, 6)]
    
  })
  
  output$table <- renderRHandsontable({
    
    #temp_data <- dim_infomation_vr()
    rhandsontable::rhandsontable(dim_infomation_vr(), digits = 6, rowHeaders = NULL) # converts the R dataframe to rhandsontable object
  })
  
  
  
  # Upload data to server and generate download button
  observeEvent(input$update_server_vr, {
    
    swd <- c("/home/pdsaadmin/dataset/VR") # 修改成server路徑: "/home/pdsaadmin/test/VR"
    setwd(swd)
    path_ <- paste0(swd, "/", substring(raw_file_vr()$name, 1, 9))
    if(!file.exists(path_)) dir.create(path_)
    f <- list.files(swd, include.dirs = F, full.names = T, recursive = T)
    #d <- list.dirs(swd, full.names = T, recursive = T)
    file.remove(f)
    #unlink(d[2:length(d)], recursive = T, force = F)
    
    
    
    
    if(input$eqi_check == TRUE){
      
      openxlsx::write.xlsx(read_excel(raw_file_vr()$datapath,
                                      sheet = raw_file_shtname_vr()[2])[1:16, ],
                           paste0(path_, "/", "Do not open.xlsx")
      )
      
      
      
      openxlsx::write.xlsx(hot_to_r(input$table),
                           paste0(path_, "/", "Dim information.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(raw_data_vr_1(),
                           paste0(path_, "/", "Cleaned data copy tool.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(equi_data(),
                           paste0(path_, "/", "Cleaned data all.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(raw_file_vr()$datapath,
                           paste0(swd, "/", "T-", raw_file_vr()$name),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(ref_file_vr()$datapath,
                           paste0(swd, "/", "R-", ref_file_vr()$name),
                           row.names = FALSE)
    } else {
      
      openxlsx::write.xlsx(read_excel(raw_file_vr()$datapath,
                                      sheet = raw_file_shtname_vr()[2])[1:16, ],
                           paste0(path_, "/", "Do not open.xlsx")
      )
      
      openxlsx::write.xlsx(hot_to_r(input$table),
                           paste0(path_, "/", "Dim information.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(raw_data_vr_1(),
                           paste0(path_, "/", "Cleaned data copy tool.xlsx"),
                           row.names = FALSE)
      
      openxlsx::write.xlsx(raw_file_vr()$datapath,
                           paste0(swd, "/", raw_file_vr()$name),
                           row.names = FALSE)
      
      
    }
    
    
    output$report_ui_vr <- renderUI({
      
      downloadButton("gen_vr", "Generate VR Report", 
                     icon = icon("arrow-alt-circle-down"),
                     class = "btn-warning btn-lg")
    })
    
    
  })
  
  # Analysis VR report and download the report
  output$gen_vr <- downloadHandler(
    
    filename = function() {
      "_Verification_ST_Rev.1.0.docx"
    },
    content = function(file) {
      
      src <- normalizePath(c('/home/pdsaadmin/dataset/report1.Rmd', '/home/pdsaadmin/dataset/template.docx')) # SEE HERE
      
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      file.copy(src, c('report1.Rmd', 'template.docx'), overwrite = TRUE) # SEE HERE
      
      #Sys.setenv(RSTUDIO_PANDOC="/usr/lib/rstudio-server/bin/pandoc")
      out <- rmarkdown::render('report1.Rmd',
                               envir = new.env(parent = globalenv())
                               
      )
      
      file.rename(out, file)
    }
    
  )
  
  
  output$left_vr <- renderImage({
    
    return(list(
      src = paste0("/home/pdsaadmin/dataset/pic/", "7",".jpg"),
      contentType = "image/jpeg",
      width = 300,
      height = 300,
      alt = "Face"
    ))
  }, deleteFile = FALSE)
  
  output$center_vr <- renderImage({
    
    return(list(
      src = paste0("/home/pdsaadmin/dataset/pic/", "7",".jpg"),
      contentType = "image/jpeg",
      width = 300,
      height = 300,
      alt = "Face"
    ))
  }, deleteFile = FALSE)
  
  output$right_vr <- renderImage({
    
    return(list(
      src = paste0("/home/pdsaadmin/dataset/pic/", "7",".jpg"),
      contentType = "image/jpeg",
      width = 300,
      height = 300,
      alt = "Face"
    ))
  }, deleteFile = FALSE)
  
  
  output$upper_vr <- renderImage({
    
    return(list(
      src = paste0("/home/pdsaadmin/dataset/pic/8.jpg"),
      contentType = "image/jpeg",
      width = 1000,
      height = 100,
      alt = "Face"
    ))
  }, deleteFile = FALSE)
  
  
  
  observeEvent(input$play_vr, {
    
    
    temp <- c(rep(7, 3208), rep(6, 3915), rep(5, 4309), rep(4, 4642), rep(3, 4932), rep(2, 5313), rep(1, 5646)) %>% 
      sample(3) %>% as.character
    
    
    output$left_vr <- renderImage({
      
      return(list(
        src = paste0("/home/pdsaadmin/dataset/pic/", temp[1],".jpg"),
        contentType = "image/jpeg",
        width = 300,
        height = 300,
        alt = "Face"
      ))
    }, deleteFile = FALSE)
    
    output$center_vr <- renderImage({
      
      return(list(
        src = paste0("/home/pdsaadmin/dataset/pic/", temp[2],".jpg"),
        contentType = "image/jpeg",
        width = 300,
        height = 300,
        alt = "Face"
      ))
    }, deleteFile = FALSE)
    
    output$right_vr <- renderImage({
      
      return(list(
        src = paste0("/home/pdsaadmin/dataset/pic/", temp[3],".jpg"),
        contentType = "image/jpeg",
        width = 300,
        height = 300,
        alt = "Face"
      ))
    }, deleteFile = FALSE)
    
    
  })
  
  
  output$count_vr <- renderValueBox({
    valueBox(h4(input$play_vr), "COUNT", color = "red")
  })
  
  
}

# 5.app
shinyApp(
  ui = dashboardPage(header, sidebar, body, skin = "red"),
  server = server
)

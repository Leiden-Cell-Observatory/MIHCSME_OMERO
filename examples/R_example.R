#Author Maarten Paul, m.w.paul@lacdr.leidenuniv.nl

if(!require("pacman")){install.packages("pacman")}
#Install R-python reticualte package
pacman::p_load(reticulate)
pacman::p_load(rspacer)
use_virtualenv("~/.virtualenvs/r-reticulate")
py_discover_config()

py_require("pandas")
py_require("zeroc_ice @ https://github.com/glencoesoftware/zeroc-ice-py-linux-x86_64/releases/download/20240202/zeroc_ice-3.6.5-cp312-cp312-manylinux_2_28_x86_64.whl")
#for Linux/Mac find the right whl https://www.glencoesoftware.com/blog/2023/12/08/ice-binaries-for-omero.html also match your Python version!
py_require("ezomero[tables]") #make use of pandas tables
#ezomero <- import("pandas",convert = FALSE)
ezomero <- import("ezomero",convert = FALSE)

#setup OMERO connection
conn <- ezomero$connect(user="root",password="omero",host="localhost",port=4064,group="system",secure=TRUE)







#Get an table from OMERO use the ID found in OMERO.weg
py$table <- ezomero$get_table(conn,file_ann_id=951L) 
py$table

#or directly to a data.frame
table <- ezomero$get_table(conn,file_ann_id=951L)
table

#upload table to OMERO and download it again
py$mtcars <- mtcars
table_id <- ezomero$post_table(conn, py$mtcars,object_type="Image", object_id=201L, title="Table mtcars")
omero_table <- ezomero$get_table(conn,file_ann_id=table_id) 

omero_table

#Get an image object and its pixels from OMERO using its ID
py$img <- ezomero$get_image(conn,51L)

#Get the pixels
pixels <- py$img[2]
data <- pixels[[1]][1, 1, , , 1]
#normalize for visualization
data <- data/max(data)
display(data, method="browser")

#Post an key-pair value
mapAnnID <- ezomero$post_map_annotation(conn,object_type="Image", object_id=201L, kv_dict=dict("Organism"="Mouse"), ns="annotation/metadata")
mapAnnID

#Upload an attachment
fileID <- ezomero$post_file_annotation(conn,file_path="MIHCSME.xlsx",object_type="Image", object_id=201L,description="Metadata template 2025_07_21")

library(readr)
library(dplyr)
library(tidyverse)
library(htmlwidgets)
library(stringr)

library(readxl)
library(writexl)
library(openxlsx)

setwd("C:/Users/cgonzalezb/Desktop/11-10-21/BASES/R_BASES/Trabajo")


#1 CARGAMOS LA BASE PRINCIPAL
BASE_TRABAJO<-as.data.frame(read.xlsx("R_CENTRO Base Ene-Abr 2022.xlsx", colNames = TRUE, detectDates = TRUE))
BASE_TRABAJO$Fecha<-paste(str_sub(BASE_TRABAJO$Fecha,9,10), str_sub(BASE_TRABAJO$Fecha,6,7), str_sub(BASE_TRABAJO$Fecha,1,4), sep = "/") #Cambiamos el formato de la fecha
#BASE_TRABAJO$Fecha<-as.Date(BASE_TRABAJO$Fecha, format = "%d/%m/%Y") Por si lo queremos regresar al estandar
#View(BASE_TRABAJO)

#BASE_TRABAJO<-BASE_TRABAJO %>% filter(Dicociclo=="LINEA" | Dicociclo=="PRUEBA") #Nos quedamos con las importantes SOLO SI ES NECESARIO

#2 Limpiamos las variables de la base y obtenemos las diferentes REGIONES de la BASE
BASE_TRABAJO$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Categoria)
BASE_TRABAJO$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Familia)
BASE_TRABAJO$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Articulo)
BASE_TRABAJO$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Descripcion)
BASE_TRABAJO$Clasificacion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Clasificacion)
BASE_TRABAJO$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Nom_Sucursal)

REGIONES<-levels(as.factor(BASE_TRABAJO$Region))
REGIONES


SUCURSALES<-levels(as.factor( (BASE_TRABAJO %>% filter(Region==REGIONES[2]))$Nom_Sucursal ))
SUCURSALES

IDs_SUCURSALES<-levels(as.factor( (BASE_TRABAJO %>% filter(Region==REGIONES[2]))$ID_Sucursal ))
IDs_SUCURSALES

SucursalesSelect<-as.data.frame(read.xlsx("R_SucursalesSeleccionadasDICOCENTRO.xlsx", colNames = TRUE))
SucursalesSelect
class(SucursalesSelect)
SucursalesSelect$ID_Sucursal[3]

SucursalesSelect$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", SucursalesSelect$Nom_Sucursal)
SucursalesSelect
SucursalesSelect$ID_Sucursal %in% IDs_SUCURSALES



#salvamos
BASE_TRABAJO_ORIGINAL<-BASE_TRABAJO
View(BASE_TRABAJO_ORIGINAL)
BASE_TRABAJO<-BASE_TRABAJO_ORIGINAL







#3 Cargamos las Pestañas del Archivo que contiene los precios de lista e inventario
#ListaPrecios<-as.data.frame(read_excel("R_CentroListaPreciosInventarioEne-Dic2019.xlsx", sheet= "ListaPrecios"))
ListaPrecios<-as.data.frame(read_excel("R_CentroListaPreciosInventario2019-2021.xlsx", sheet= "ListaPrecios"))
View(ListaPrecios)
#ListaPrecios<-as.data.frame(read.csv("R_CentroListaPreciosInventario2019-2021.csv", header = TRUE))

#Inventario<-as.data.frame(read_excel("R_CentroListaPreciosInventarioEne-Oct2021.xlsx", sheet= "Inventario"))


#4 Limpiamos las tablas: Quitamos espacios (y espacios especiales) al principio y al fin de los datos
ListaPrecios<-ListaPrecios[which(!is.na(ListaPrecios$NuevoKit)), ]
ListaPrecios$NuevoKit<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$NuevoKit)
ListaPrecios$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$Descripcion)


#Inventario<-Inventario[which(!is.na(Inventario$Articulo)), ]
#Inventario$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Articulo)
#Inventario$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Familia)
#Inventario$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Categoria)
#Inventario$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Descripcion)


#5 Limpiamos las tablas: Quitamos duplicados 
ListaPrecios<-QuitarMultiplicados_Simple1(ListaPrecios,"NuevoKit")
#Inventario<-QuitarMultiplicados_Simple1(Inventario,"Articulo")

#6 A la tabla ListaPrecios la separamos por año y le creamos una nueva columna
ListaPreciosOriginal <- ListaPrecios

ListaPrecios <- ListaPreciosOriginal[,c(1,2,3:14)]
#ListaPrecios <- ListaPreciosOriginal[,c(1,2,15:26)]
#ListaPrecios <- ListaPreciosOriginal[,c(1,2,27:38)]
ListaPrecios$PrecioLista_Prom<-apply(ListaPrecios[,c(3:ncol(ListaPrecios))], MARGIN = 1, FUN = mean, na.rm= TRUE)














BASE_TRABAJO<-BASE_TRABAJO %>% filter( Region=="DICO CENTRO", Nom_Sucursal=="5 de Mayo")
View(BASE_TRABAJO)





NombresArch=c("Series_5deMayo2019")
UltiMes="11/2019"
preciolista=1
disp=0
ListaOBJETOS<-list()




Resumen<-TablasResumenSimple(BASE_TRABAJO, "DICO CENTRO", "Utilidad", UltiMes)
CATEGORIAS<-Resumen[["CategoriasOrdenadas"]] 
Resumen[["CategoriasOrdenadas"]]<-gsub("[[:punct:]|\\+|¿|\\|¡]", " ", Resumen[["CategoriasOrdenadas"]])

ListaCategoriasTotales<-lapply(CATEGORIAS, TablasPorCategoria, BASE_TRABAJO, "DICO CENTRO",UltiMes)
names(ListaCategoriasTotales)<-CATEGORIAS
ListaOBJETOS[[1]]<-list(RESUMEN=Resumen, 
                        LISTACATEGORIAS=ListaCategoriasTotales)

ListaCategoriasTotales<-lapply(ListaCategoriasTotales, Agregar_PrecioLista_Dispoibilidad, ListaPrecios)

ListaOBJETOS[[1]]<-list(RESUMEN=Resumen, 
                        LISTACATEGORIAS=ListaCategoriasTotales)


AnalisisGral<-TablasGral(ListaCategoriasTotales,Resumen[["MesActual"]], Disponibilidad = disp)
ListaOBJETOS[[1]]<-list(RESUMEN=Resumen, 
                        LISTACATEGORIAS=ListaCategoriasTotales, 
                        ANALISISGRAL=AnalisisGral, 
                        PRECIOLISTA=preciolista,
                        DISPONIBILIDAD=disp,
                        NOMBREARCHIVO=NombresArch[1])

ExportarTablasExcelCompleto(Resumen= ListaOBJETOS[[1]][["RESUMEN"]] , ListaTablasCategoria= ListaOBJETOS[[1]][["LISTACATEGORIAS"]], ListaTablasGral= ListaOBJETOS[[1]][["ANALISISGRAL"]], ListaOBJETOS[[1]][["PRECIOLISTA"]], ListaOBJETOS[[1]][["DISPONIBILIDAD"]], NombreArchivo= ListaOBJETOS[[1]][["NOMBREARCHIVO"]])






####################### INICIAMOS PARA AÑOS COMPLETOS POR SUCURSAL ###############################
SucursalesSelect<-as.data.frame(read.xlsx("R_SucursalesSeleccionadasDICOCENTRO.xlsx", colNames = TRUE))
SucursalesSelect$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", SucursalesSelect$Nom_Sucursal)
SucursalesSelect



NombresBases<-c("R_CentroBaseEne-Dic2019.xlsx",
                "R_CentroBaseEne-Dic2020.xlsx",
                "R_CentroBaseEne-Dic2021.xlsx",
                "R_CentroBaseEne-Dic2018.xlsx")

ListaPrecios<-as.data.frame(read_excel("R_CentroListaPreciosInventario2019-2021.xlsx", sheet= "ListaPrecios"))
#Limpiamos la Lista de Precios: Quitamos NA, quitamos espacios (y espacios especiales) al principio y al fin de los datos
#Y quitamos duplicados
ListaPrecios<-ListaPrecios[which(!is.na(ListaPrecios$NuevoKit)), ]
ListaPrecios$NuevoKit<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$NuevoKit)
ListaPrecios$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$Descripcion)
ListaPrecios<-QuitarMultiplicados_Simple1(ListaPrecios,"NuevoKit")
#Salvamos PORQUE después la vamos a estar partiendo y reutilizando
ListaPreciosOriginal<-ListaPrecios

#Separamos por año
ListaPreciosAnios <- list(Precios2019= ListaPreciosOriginal[,c(1,2,3:14)], 
                          Precios2020= ListaPreciosOriginal[,c(1,2,15:26)], 
                          Precios2021= ListaPreciosOriginal[,c(1,2,27:38)])

VectUltimes<-c("11/2019","11/2020","11/2021","11/2018")
VectpreciolistaBooleano<-c(1,1,1,0)
VectdispBooleano<-c(0,0,0,0)


#ANTES HAY QUE CORROBORAR QUE LOS IDs DE LAS SUCURSALES SÍ ESTÉN EN LAS BASES DE EXCEL
paste("Series", SucursalesSelect$Nom_Sucursal, "2019", sep = "_")
paste("Series", SucursalesSelect$Nom_Sucursal, "2020", sep = "_")
paste("Series", SucursalesSelect$Nom_Sucursal, "2021", sep = "_")

NombresArchAnios<-list(NombresArch2019= paste("Series", SucursalesSelect$Nom_Sucursal, "2019", sep = "_"), 
                       NombresArch2020= paste("Series", SucursalesSelect$Nom_Sucursal, "2020", sep = "_"), 
                       NombresArch2021= paste("Series", SucursalesSelect$Nom_Sucursal, "2021", sep = "_"),
                       NombresArch2018= paste("Series", SucursalesSelect$Nom_Sucursal, "2018", sep = "_"))

ListaOBJETOS<-list()
Contador= 78
REGIONES<-c("DICO BAJÍO","DICO CENTRO","EXPO","MÖBLUM" )
for(n in c(4:4)) #para años del 2019 al 2021
{
  
  #1 CARGAMOS LA BASE PRINCIPAL
  BASE_TRABAJO<-as.data.frame(read.xlsx(NombresBases[n], colNames = TRUE, detectDates = TRUE))
  BASE_TRABAJO$Fecha<-paste(str_sub(BASE_TRABAJO$Fecha,9,10), str_sub(BASE_TRABAJO$Fecha,6,7), str_sub(BASE_TRABAJO$Fecha,1,4), sep = "/") #Cambiamos el formato de la fecha
  #BASE_TRABAJO$Fecha<-as.Date(BASE_TRABAJO$Fecha, format = "%d/%m/%Y") Por si lo queremos regresar al estandar
  
  
  
  UltiMes <- VectUltimes[n]
  preciolista <- VectpreciolistaBooleano[n]
  disp <- VectdispBooleano[n]
  
  NombresArch<-NombresArchAnios[[n]]
  
  
  #2 Limpiamos las variables de la base y obtenemos las diferentes REGIONES de la BASE
  BASE_TRABAJO$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Categoria)
  BASE_TRABAJO$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Familia)
  BASE_TRABAJO$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Articulo)
  BASE_TRABAJO$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Descripcion)
  BASE_TRABAJO$Clasificacion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Clasificacion)
  BASE_TRABAJO$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Nom_Sucursal)
  #salvamos
  BASE_TRABAJO_ORIGINAL<-BASE_TRABAJO
  
  
  #3 Creamos una columna al final con el promedio de los precios de lista
  if(preciolista==1)
  {
    ListaPrecios<-ListaPreciosAnios[[n]]
    ListaPrecios$PrecioLista_Prom<-apply(ListaPrecios[,c(3:ncol(ListaPrecios))], MARGIN = 1, FUN = mean, na.rm= TRUE)
  }
  
  
  #4 COMENZAMOS: armamos un archivo por cada sucursal 
  for(s in c(1:(nrow(SucursalesSelect))))
  {
    Contador<-Contador+1
    BASE_TRABAJO<-BASE_TRABAJO_ORIGINAL %>% filter( Region==REGIONES[2], ID_Sucursal== SucursalesSelect$ID_Sucursal[s])
    
    
    Resumen<-TablasResumenSimple(BASE_TRABAJO, REGIONES[2], "Utilidad", UltiMes)
    CATEGORIAS<-Resumen[["CategoriasOrdenadas"]] 
    Resumen[["CategoriasOrdenadas"]]<-gsub("[[:punct:]|\\+|¿|\\|¡]", " ", Resumen[["CategoriasOrdenadas"]])
    
    ListaCategoriasTotales<-lapply(CATEGORIAS, TablasPorCategoria, BASE_TRABAJO, REGIONES[2], UltiMes)
    names(ListaCategoriasTotales)<-CATEGORIAS
    ListaOBJETOS[[Contador]]<-list(RESUMEN=Resumen, 
                                   LISTACATEGORIAS=ListaCategoriasTotales)
    
    if(preciolista==1)
    {
      ListaCategoriasTotales<-lapply(ListaCategoriasTotales, Agregar_PrecioLista_Dispoibilidad, ListaPrecios)
      ListaOBJETOS[[Contador]]<-list(RESUMEN=Resumen, 
                                     LISTACATEGORIAS=ListaCategoriasTotales)
    }
    
    AnalisisGral<-TablasGral(ListaCategoriasTotales,Resumen[["MesActual"]], Disponibilidad = disp)
    ListaOBJETOS[[Contador]]<-list(RESUMEN=Resumen, 
                                   LISTACATEGORIAS=ListaCategoriasTotales, 
                                   ANALISISGRAL=AnalisisGral, 
                                   PRECIOLISTA=preciolista, 
                                   DISPONIBILIDAD=disp, 
                                   NOMBREARCHIVO=NombresArch[s])
    
    ExportarTablasExcelCompleto(Resumen= ListaOBJETOS[[Contador]][["RESUMEN"]], 
                                ListaTablasCategoria= ListaOBJETOS[[Contador]][["LISTACATEGORIAS"]], 
                                ListaTablasGral= ListaOBJETOS[[Contador]][["ANALISISGRAL"]], 
                                ListaOBJETOS[[Contador]][["PRECIOLISTA"]], 
                                ListaOBJETOS[[Contador]][["DISPONIBILIDAD"]], 
                                NombreArchivo= ListaOBJETOS[[Contador]][["NOMBREARCHIVO"]])
    
  }
  
  
  
}


names(ListaOBJETOS)[c(1:26)]<-paste(SucursalesSelect$Nom_Sucursal,"2019",sep = "_")
names(ListaOBJETOS)[c(27:52)]<-paste(SucursalesSelect$Nom_Sucursal,"2020",sep = "_")
names(ListaOBJETOS)[c(53:78)]<-paste(SucursalesSelect$Nom_Sucursal,"2021",sep = "_")
names(ListaOBJETOS)[c(79:104)]<-paste(SucursalesSelect$Nom_Sucursal,"2018",sep = "_")


names(ListaOBJETOS)
save(ListaOBJETOS, file = "Series26SucursalesCentro2018-2021.RData")

###################################################################



###################### INICIAMOS PARA AÑO INCOMPLETO POR SUCURSAL ##################################
setwd("C:/Users/cgonzalezb/Desktop/11-10-21")
load("C:/Users/cgonzalezb/Desktop/11-10-21/CategoriasFamiliasProductosClasificacion3.RData")
setwd("C:/Users/cgonzalezb/Desktop/11-10-21/BASES/R_BASES/Trabajo")





SucursalesSelect<-as.data.frame(read.xlsx("R_SucursalesSeleccionadas_DICOCENTRO.xlsx", colNames = TRUE))
SucursalesSelect$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", SucursalesSelect$Nom_Sucursal)
SucursalesSelect
#SucursalesSelect<-SucursalesSelect[c(13),]


#DEPENDIENDO DE LAS BASES:
#NombresBases<-c("R_CENTRO Base Ene-Abr 2022.xlsx","R_CENTRO Base Ene-Dic 2021.xlsx","R_CENTRO Base Ene-Dic 2020.xlsx","R_CENTRO Base Ene-Dic 2019.xlsx","R_CENTRO Base Ene-Dic 2018.xlsx")
#NombresBases<-c("R_CENTRO Base Ene-Dic 2018.xlsx","R_CENTRO Base Ene-Dic 2019.xlsx","R_CENTRO Base Ene-Dic 2020.xlsx","R_CENTRO Base Ene-Dic 2021.xlsx","R_CENTRO Base Ene-Abr 2022.xlsx")
NombresBases<-c("R_CENTRO Base Ene-Dic 2018.xlsx","R_CENTRO Base Ene-Dic 2019.xlsx","R_CENTRO Base Ene-Dic 2020.xlsx","R_CENTRO Base Ene-Dic 2021.xlsx","R_CENTRO Base Ene-Ago 2022.xlsx")
NombresBases<-c("R_CENTRO Base Ene-Dic 2021.xlsx","R_CENTRO Base Ene-Dic 2022.xlsx")
NombresBases<-c("R_CENTRO Base Ene-May 2023.xlsx")
NombresBases<-c("R_BAJIO Base Ene01-Ene29 2023.xlsx")
#NombresBases<-c("R_BAJIO Base Ene-Dic 2018.xlsx","R_BAJIO Base Ene-Dic 2019.xlsx","R_BAJIO Base Ene-Dic 2020.xlsx","R_BAJIO Base Ene-Dic 2021.xlsx","R_BAJIO Base Ene-Jul 2022.xlsx")
#NombresBases<-c("R_5 de Mayo-Angelópolis Base Ene2021-Jun2022.xlsx")


#ListaPrecios<-as.data.frame(read_excel("R_CentroListaPreciosInventario2019-2021.xlsx", sheet= "ListaPrecios"))
ListaPrecios<-as.data.frame(read.xlsx("R_CentroListaPreciosInventario2019-2021.xlsx", sheet= "ListaPrecios", colNames = TRUE))


#Limpiamos la Lista de Precios: Quitamos NA, quitamos espacios (y espacios especiales) al principio y al fin de los datos
#Y quitamos duplicados
ListaPrecios<-ListaPrecios[which(!is.na(ListaPrecios$NuevoKit)), ]
ListaPrecios$NuevoKit<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$NuevoKit)
ListaPrecios$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$Descripcion)
ListaPrecios<-QuitarMultiplicados_Simple1(ListaPrecios,"NuevoKit")
#Salvamos PORQUE después la vamos a estar partiendo y reutilizando
ListaPreciosOriginal<-ListaPrecios

#Separamos por año
ListaPreciosAnios <- list(Precios2018= NA,
                          Precios2019= ListaPreciosOriginal[,c(1,2,3:14)], 
                          Precios2020= ListaPreciosOriginal[,c(1,2,15:26)], 
                          Precios2021= ListaPreciosOriginal[,c(1,2,27:38)],
                          Precios2022= NA)

#Separamos por año
ListaPreciosAnios <- list(Precios2021= ListaPreciosOriginal[,c(1,2,27:38)],
                          Precios2022= NA)


#Multiples años
VectUltimes<-c("11/2018","11/2019","11/2020","11/2021","09/2022")
VectpreciolistaBooleano<-c(0,1,1,1,0)
VectdispBooleano<-c(0,0,0,0,0)


#Multiples años
VectUltimes<-c("11/2021","11/2022")
VectpreciolistaBooleano<-c(0,0)
VectdispBooleano<-c(0,0)

#Solo 2023
VectUltimes<-c("05/2023")
VectpreciolistaBooleano<-c(0)
VectdispBooleano<-c(0)



#ANTES HAY QUE CORROBORAR QUE LOS IDs DE LAS SUCURSALES SÍ ESTÉN EN LAS BASES DE EXCEL
#paste("Series", SucursalesSelect$Nom_Sucursal, "Ene-Mar_2022", sep = "_")



NombresBases
#Loc_Seleccionados<-sapply(gregexpr("^(R_)", NombresBases), '[', 1)
#Loc_Seleccionados
#NombresBases[which(Loc_Seleccionados==1)]
Loc_Inicio<-sapply(gregexpr("Base ", NombresBases), '[', 1)
Loc_Inicio
Loc_Inicio <- Loc_Inicio + 5
Loc_Fin<-sapply(gregexpr("\\.", NombresBases), '[', 1)
Loc_Fin
Loc_Fin <- Loc_Fin - 1

str_sub(NombresBases, Loc_Inicio, Loc_Fin)
str_sub(NombresBases, Loc_Inicio, Loc_Fin)[1]






#En caso de que sean los años 2018-2022
NombresArchAnios<-list(NombresArch2018= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[1], sep = "_"),
                       NombresArch2019= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[2], sep = "_"),
                       NombresArch2020= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[3], sep = "_"),
                       NombresArch2021= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[4], sep = "_"),
                       NombresArch2022= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[5], sep = "_"))

#En caso de que sean los años 2021-2022
NombresArchAnios<-list(NombresArch2021= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[1], sep = "_"),
                       NombresArch2022= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[2], sep = "_"))


#SOLO EN CASO DE que no todas las sucursales estén en todos los años a analizar (archivos), MODIFICAMOS (hay que quitar las sucursales que no aparecen en los años correspondientes)
NombresArchAnios
#NombresArchAnios[["NombresArch2018"]]<-NombresArchAnios[["NombresArch2018"]][-c(27,28)]
#NombresArchAnios[["NombresArch2019"]]<-NombresArchAnios[["NombresArch2019"]][-c(27,28)]
#NombresArchAnios[["NombresArch2020"]]<-NombresArchAnios[["NombresArch2020"]][-c(27,28)]
#NombresArchAnios[["NombresArch2021"]]<-NombresArchAnios[["NombresArch2021"]][-c(28)]


#EN CASO DE QUE SOLO SEA 2023
NombresArchAnios<-list(NombresArch2023= paste("Series", SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin)[1], sep = "_"))
NombresArchAnios



FechaIntervalo <- c() #en caso de que se requiera toda la base
FechaIntervalo <- c("2022-11-01","2022-11-30") #En caso de que se requiera el intervalo cerrado


ListaOBJETOS<-list()
Contador= 0

#REGIONES<-levels(as.factor(BASE_TRABAJO$Region))
#REGIONES
REGIONES<-c("DICO CENTRO","EXPO","MÖBLUM" ) #siempre la región correcta siempre tiene que ser la primera (se modificará para mejorar eso)
#REGIONES<-c("DICO BAJÍO","DICO BAJÍO QRO" )
REGIONES<-c("DICO BAJÍO PUEBLA")

for(n in c(1:length(NombresBases))) #n es para los años
{
  
  #1 CARGAMOS LA BASE PRINCIPAL
  cat("Cargando el archivo: ", NombresBases[n], "\n")
  BASE_TRABAJO<-as.data.frame(read.xlsx(NombresBases[n], colNames = TRUE, detectDates = TRUE))
  
  #VEMOS SI UTILIZAREMOS SOLO UN INTERVALO DE LA FECHA
  if(!is.null(FechaIntervalo))
  {
    #BASE_TRABAJO <- BASE_TRABAJO[which( BASE_TRABAJO$Fecha >= FechaIntervalo[1] & BASE_TRABAJO$Fecha <= FechaIntervalo[2]),]
    BASE_TRABAJO <- BASE_TRABAJO %>% filter(Fecha >= FechaIntervalo[1] & Fecha <= FechaIntervalo[2])
  }
  
  #BASE_TRABAJO$Fecha<-paste(str_sub(BASE_TRABAJO$Fecha,9,10), str_sub(BASE_TRABAJO$Fecha,6,7), str_sub(BASE_TRABAJO$Fecha,1,4), sep = "/") #Cambiamos el formato de la fecha
  BASE_TRABAJO$Fecha<-paste(format(BASE_TRABAJO$Fecha,"%d"), format(BASE_TRABAJO$Fecha,"%m"), format(BASE_TRABAJO$Fecha,"%Y"), sep = "/") #Cambiamos el formato de la fecha
  #BASE_TRABAJO$Fecha<-as.Date(BASE_TRABAJO$Fecha, format = "%d/%m/%Y") #Por si lo queremos regresar al estandar
  
  
  
  UltiMes <- VectUltimes[n]
  preciolista <- VectpreciolistaBooleano[n]
  disp <- VectdispBooleano[n]
  
  NombresArch<-NombresArchAnios[[n]]
  
  
  #2 Limpiamos las variables de la base y obtenemos las diferentes REGIONES de la BASE
  cat("Limpiando\n")
  BASE_TRABAJO$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Categoria)
  BASE_TRABAJO$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Familia)
  BASE_TRABAJO$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Articulo)
  BASE_TRABAJO$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Descripcion)
  BASE_TRABAJO$Clasificacion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Clasificacion)
  BASE_TRABAJO$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Nom_Sucursal)
  #salvamos
  BASE_TRABAJO_ORIGINAL<-BASE_TRABAJO
  
  
  #3 Creamos una columna al final con el promedio de los precios de lista
  cat("Valores de preciolista y disp: ", preciolista[n], ", ", disp[n], "\n")
  if(preciolista==1)
  {
    ListaPrecios<-ListaPreciosAnios[[n]]
    ListaPrecios$PrecioLista_Prom<-apply(ListaPrecios[,c(3:ncol(ListaPrecios))], MARGIN = 1, FUN = mean, na.rm= TRUE)
  }
  
  
  #4 COMENZAMOS: armamos un archivo por cada sucursal 
  # pero antes de eso tenemos que extraer los Nom de las Suc que se analizarán en su respectivo año
  Loc_Inicio<-sapply(gregexpr("Series_", NombresArchAnios[[n]]), '[', 1)
  #Loc_Inicio+6
  Loc_Fin<-sapply(gregexpr("_Ene", NombresArchAnios[[n]]), '[', 1)
  #Loc_Fin
  str_sub(NombresArchAnios[[n]], Loc_Inicio+7, Loc_Fin-1)
  SucursalesSelect<-str_sub(NombresArchAnios[[n]], Loc_Inicio+7, Loc_Fin-1)
  
  for( s in c(1:length(SucursalesSelect)) )
  {
    Contador<-Contador+1
    BASE_TRABAJO<-BASE_TRABAJO_ORIGINAL %>% filter( Region==REGIONES[1], Nom_Sucursal== SucursalesSelect[s])
    
    
    Resumen<-TablasResumenSimple(BASE_TRABAJO, "Utilidad", UltiMes)
    Resumen[["BaseOrigen"]]$Fecha <- as.Date(Resumen[["BaseOrigen"]]$Fecha, format = "%d/%m/%Y") #regresamos al formato de fecha estandar
    CATEGORIAS<-Resumen[["CategoriasOrdenadas"]] 
    Resumen[["CategoriasOrdenadas"]]<-gsub("[[:punct:]|\\+|¿|\\|¡]", " ", Resumen[["CategoriasOrdenadas"]])
    
    print("INICIÓ LA FUNCIÓN DE LAS LISTAS POR CATEGORÍA")
    ListaCategoriasTotales<-lapply(CATEGORIAS, TablasPorCategoria, BASE_TRABAJO, UltiMes)
    names(ListaCategoriasTotales)<-CATEGORIAS
    ListaOBJETOS[[Contador]]<-list(RESUMEN=Resumen, 
                                   LISTACATEGORIAS=ListaCategoriasTotales)
    
    if(preciolista==1)
    {
      print("INICIÓ LA FUNCIÓN DE AGREGAR PRECIOS DE LISTA")
      ListaCategoriasTotales<-lapply(ListaCategoriasTotales, Agregar_PrecioLista_Dispoibilidad, ListaPrecios)
      ListaOBJETOS[[Contador]]<-list(RESUMEN=Resumen, 
                                     LISTACATEGORIAS=ListaCategoriasTotales)
    }
    
    print("INICIÓ LA FUNCIÓN DEL ANÁLISIS DE LAS CATEGORÍAS JUNTAS")
    AnalisisGral<-TablasGral(ListaCategoriasTotales,Resumen[["MesActual"]], Disponibilidad = disp)
    ListaOBJETOS[[Contador]]<-list(RESUMEN=Resumen, 
                                   LISTACATEGORIAS=ListaCategoriasTotales, 
                                   ANALISISGRAL=AnalisisGral, 
                                   PRECIOLISTA=preciolista, 
                                   DISPONIBILIDAD=disp, 
                                   NOMBREARCHIVO=NombresArch[s])
    names(ListaOBJETOS)[Contador]<-NombresArchAnios[[n]][s]
    
    print("INICIÓ LA FUNCIÓN DE EXPORTACIÓN A EXCEL")
    ExportarTablasExcelCompleto(Resumen= ListaOBJETOS[[Contador]][["RESUMEN"]], 
                                ListaTablasCategoria= ListaOBJETOS[[Contador]][["LISTACATEGORIAS"]], 
                                ListaTablasGral= ListaOBJETOS[[Contador]][["ANALISISGRAL"]], 
                                ListaOBJETOS[[Contador]][["PRECIOLISTA"]], 
                                ListaOBJETOS[[Contador]][["DISPONIBILIDAD"]], 
                                NombreArchivo= ListaOBJETOS[[Contador]][["NOMBREARCHIVO"]])
    
  }
  
  
  
}

NombresArchAnios

#CUANDO SOLO ES PARA JR 
save(ListaOBJETOS, file = "SeriesSucursalesCENTRO_OTROS_2022.RData")



#2018-2021
names(ListaOBJETOS)[c(1:110)]<-c(NombresArchAnios[[1]],NombresArchAnios[[2]],NombresArchAnios[[3]],NombresArchAnios[[4]])
save(ListaOBJETOS[c(1:110)], file = "SeriesSucursalesCENTRO_2018-2021.RData")
#2022
names(ListaOBJETOS)[c(111:139)]<-c(NombresArchAnios[[5]])
save(ListaOBJETOS[c(111:139)], file = "SeriesSucursalesCENTRO_2022.RData")




#CUANDO ES PARA BAJIO 2018-2022
ListaOBJETOS_2_P1 <- ListaOBJETOS[c(1:8)]
ListaOBJETOS_2_P2 <- ListaOBJETOS[c(9,10)]

save(ListaOBJETOS_2_P1, file = "SeriesSucursalesBAJIO_2018-2021.RData")
save(ListaOBJETOS_2_P2, file = "SeriesSucursalesBAJIO_2022.RData")


#CUANDO SOLO ES PARA BAJIO 2022
save(ListaOBJETOS, file = "SeriesSucursalesBAJIO_2022.RData")

#CUANDO SOLO ES PARA CENTRO 2023
save(ListaOBJETOS, file = "SeriesSucursalesCENTRO_2023.RData")

#CUANDO SOLO ES PARA JR 
save(ListaOBJETOS, file = "SeriesSucursalesJR_2022.RData")

#CUANDO SOLO ES PARA OTROS 
save(ListaOBJETOS, file = "SeriesTuxtla_2018-2022.RData")




NombresArchAnios
names(ListaOBJETOS)<-as_vector(NombresArchAnios)
names(ListaOBJETOS)
save(ListaOBJETOS, file = "SeriesParquePuebla2018-2022.RData")



names(ListaOBJETOS)[c(1:26)]<-paste(SucursalesSelect$Nom_Sucursal,"2019",sep = "_")
names(ListaOBJETOS)[c(27:52)]<-paste(SucursalesSelect$Nom_Sucursal,"2020",sep = "_")
names(ListaOBJETOS)[c(53:78)]<-paste(SucursalesSelect$Nom_Sucursal,"2021",sep = "_")
names(ListaOBJETOS)[c(79:104)]<-paste(SucursalesSelect$Nom_Sucursal,"2018",sep = "_")


names(ListaOBJETOS)
save(ListaOBJETOS, file = "Series26SucursalesCentro2018-2021.RData")


names(ListaOBJETOS)[c(1:28)]<-paste(SucursalesSelect$Nom_Sucursal, str_sub(NombresBases, Loc_Inicio, Loc_Fin-1), sep = "_")
save(ListaOBJETOS, file = paste0("Series28SucursalesCentro_" , str_sub(NombresBases, Loc_Inicio, Loc_Fin-1), ".RData"))






#######################################################################################




























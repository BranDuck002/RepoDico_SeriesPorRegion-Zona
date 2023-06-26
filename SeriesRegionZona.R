
library(readr)
library(dplyr)
library(tidyverse)
library(htmlwidgets)
library(stringr)

library(readxl)
library(writexl)
library(openxlsx)




#0 VAMOS A LA DIRECCIÓN DEL WORKSPACE, CARGAMOS EL WORKSPACE Y NOS MOVEMOS A LA CARPETA DE TRABAJO

setwd("C:/Users/cgonzalezb/Desktop/11-10-21/BASES/R_BASES/Trabajo/ReposDico/RepoDico_SeriesPorRegion-Zona") #Directorio en donde se encuentra el script y sus dependencias
load("C:/Users/cgonzalezb/Desktop/11-10-21/BASES/R_BASES/Trabajo/ReposDico/RepoDico_SeriesPorRegion-Zona/CategoriasFamiliasProductosClasificacion3.RData") #Dependencias
setwd("C:/Users/cgonzalezb/Desktop/11-10-21/BASES/R_BASES/Trabajo") #Directorio en donde se encuentran las bases que serán trabajadas




# EL BUENO!!!!!!!!!!!!!!!!!!!! (PRODUCCIÓN EN MASA)
######################################

#1 CARGAMOS LOS NOMBRES DE LOS ARCHIVOS EN LA CARPETA DE TRABAJO
NombresArchivos<-list.files()
NombresArchivos

#2 IDENTIFICAMOS LOS NOMBRES DE LAS BASES (datos descargados de la región)
NombresBases<- NombresArchivos[which(str_detect(NombresArchivos,"^(R_)"))]
NombresBases<- NombresBases[which(str_detect(NombresBases,"(B|b)(A|a)(S|s)(E|e)"))]
NombresBases<- NombresBases[which(str_detect(NombresBases,"(2023)"))]
NombresBases
#NombresBases<-NombresBases[-c(2,4)]
#NombresBases<-NombresBases[c(4)]

#Ignorar sección según corresponda
#3 IDENTIFICAMOS LOS NOMBRES DE LOS ARCHIVOS CON LOS PRECIOS DE LISTA
NombresPrecios<- NombresArchivos[which(str_detect(NombresArchivos,"^(R_)"))]
NombresPrecios<- NombresPrecios[which(str_detect(NombresPrecios,"(P|p)(R|r)(E|e)(C|c)(I|i)(O|o)(S|s)"))]
NombresPrecios


#4 Revisamos los nombres e identificamos en dónde comienzan las fechas y en dónde terminan
NombresBases

Loc_Inicio<-sapply(gregexpr("Base ", NombresBases), '[', 1)
Loc_Inicio
Loc_Inicio<-Loc_Inicio+5
Loc_Fin<-sapply(gregexpr("\\.", NombresBases), '[', 1)
Loc_Fin
Loc_Fin<-Loc_Fin-1

str_sub(NombresBases, Loc_Inicio, Loc_Fin)
#paste("Series", REGIONES[n], str_sub(NombresBases[m], Loc_Inicio[m], Loc_Fin[m]-1), sep = " ")





#5 AFIRMAMOS SI CONTAMOS CON LOS PRECIOS DE LISTA PARA CADA BASE CORRESPONDIENTEMENTE CON BOOLEANOS
#HAY QUE MODIFICAR LA FUNCIÓN PRINCIPAL EN CASO DE QUE HAYA MÁS DE UN ARCHIVO DE PRECIOS DE LISTA
#(tiene que coincidir la cardinalidad de NombresPrecios y preciolista)
NombresPrecios<-c(NA,NombresPrecios)
NombresPrecios

NombresPrecios<-c(NA,NA,NA,NA,NA)
NombresPrecios

NombresPrecios<-c(NA)
NombresPrecios

#BOOLEANOS OBLIGATORIOS (por base)
preciolista=c(0,0,0,0,0)
disp=c(0,0,0,0,0)

preciolista=c(0)
disp=c(0)


fechacorte<-NULL #agregar las fechas de corte coorespondientes a las bases (SOLO EN CASO DE HABER) con formato (ACTUALMENTE SIEMPRE ES NULL, YA QUE LOS ARCHIVOS DE DISPONIBILIDAD YA NO SE USAN)

#6 LOS DEMÁS PARÁMETROS OBLIGATORIOS
UltiMes=c("06/2023","06/2023","06/2023","06/2023","06/2023")
UltiMes=c("05/2023")

FechaIntervalo <- c() #en caso de que se requiera toda la base
FechaIntervalo <- c("2022-07-01","2023-05-31") #En caso de que se requiera el intervalo cerrado

ParcheFiltro=0 #booleano; 0= NO, 1= SÍ, en caso de que se requiera un filtro especial como modificar zonas

ListaOBJETOS<-list()
NombrArch<-c()
contador=0 #cuenta las REGIONES que va trabajando


#EN CASO DE QUE SE NECESITE EXPERIMENTAR PARA MODIFICAR O QUITAR ZONAS (PARCHE)
########

ParcheFiltro=1 #ES NECESARIO ACTIVAR EL PARCHE

cat("Cargando el archivo: ", NombresBases[1], "\n")
BASE_TRABAJO<-as.data.frame(read.xlsx(NombresBases[1], colNames = TRUE, detectDates = TRUE))



View(BASE_TRABAJO)
View(BASE_TRABAJO %>% filter(Region=="DICO CENTRO"))
View(BASE_TRABAJO %>% filter(Nom_Sucursal=="Teziutlan" | (Zona!="Zona Puebla" & Zona!="Zona Oaxaca")) )


BASE_TRABAJO <- BASE_TRABAJO %>% filter(Region=="DICO CENTRO")
BASE_TRABAJO <- BASE_TRABAJO %>% filter(Nom_Sucursal=="Teziutlan" | (Zona!="Zona Puebla" & Zona!="Zona Oaxaca")) 


m <- 1





########


# m para NombresBases
# n para Regiones
# z para zonas

#contador=0
#m

#7 CORREMOS EL PROGRAMA PRINCIPAL PARA CADA BASE (si comenzamos desde el principio, entonces comenzamos desde m = 1)
for(m in c(1:length(NombresBases)))#for(m in c(1,3,4,5))
{
  #1 CARGAMOS LA BASE PRINCIPAL Lomas Norte Sureste
  cat("Cargando el archivo: ", NombresBases[m], "\n")
  BASE_TRABAJO<-as.data.frame(read.xlsx(NombresBases[m], colNames = TRUE, detectDates = TRUE))
  
  #VEMOS SI UTILIZAREMOS SOLO UN INTERVALO DE LA FECHA
  if(!is.null(FechaIntervalo))
  {
    #BASE_TRABAJO <- BASE_TRABAJO[which( BASE_TRABAJO$Fecha >= FechaIntervalo[1] & BASE_TRABAJO$Fecha <= FechaIntervalo[2]),]
    BASE_TRABAJO <- BASE_TRABAJO %>% filter(Fecha >= FechaIntervalo[1] & Fecha <= FechaIntervalo[2])
  }
  
  
  if(ParcheFiltro==1)
  {
    
    #BASE_TRABAJO <- BASE_TRABAJO %>% filter(Region=="DICO CENTRO")
    #BASE_TRABAJO <- BASE_TRABAJO %>% filter(Nom_Sucursal=="Teziutlan" | (Zona!="Zona Puebla" & Zona!="Zona Oaxaca")) 
    
    #BASE_TRABAJO <- BASE_TRABAJO %>% filter(Region=="DICO CENTRO")
    #BASE_TRABAJO <- BASE_TRABAJO %>% filter( (Zona=="Zona Puebla" | Zona=="Zona Oaxaca") & Nom_Sucursal!="Teziutlan" ) 
    
    BASE_TRABAJO <- BASE_TRABAJO %>% filter(Region=="DICO BAJÍO PUEBLA")
    BASE_TRABAJO <- BASE_TRABAJO %>% filter(Nom_Sucursal=="5 de Mayo" | Nom_Sucursal=="Angelópolis") 
    
    
  }
  
  
  
  
  #BASE_TRABAJO$Fecha<-paste(str_sub(BASE_TRABAJO$Fecha,9,10), str_sub(BASE_TRABAJO$Fecha,6,7), str_sub(BASE_TRABAJO$Fecha,1,4), sep = "/") #Cambiamos el formato de la fecha
  BASE_TRABAJO$Fecha<-paste(format(BASE_TRABAJO$Fecha,"%d"), format(BASE_TRABAJO$Fecha,"%m"), format(BASE_TRABAJO$Fecha,"%Y"), sep = "/") #Cambiamos el formato de la fecha
  #BASE_TRABAJO$Fecha<-as.Date(BASE_TRABAJO$Fecha, format = "%d/%m/%Y") #Por si lo queremos regresar al estandar
  #View(BASE_TRABAJO)
  
  #BASE_TRABAJO<-BASE_TRABAJO %>% filter(Dicociclo=="LINEA" | Dicociclo=="PRUEBA") #Nos quedamos con las importantes SOLO SI ES NECESARIO
  
  #2 Limpiamos las variables de la base y obtenemos las diferentes REGIONES de la BASE
  cat("Limpiando\n")
  BASE_TRABAJO$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Categoria)
  BASE_TRABAJO$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Familia)
  BASE_TRABAJO$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Articulo)
  BASE_TRABAJO$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Descripcion)
  BASE_TRABAJO$Clasificacion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Clasificacion)
  BASE_TRABAJO$Nom_Sucursal<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", BASE_TRABAJO$Nom_Sucursal)
  BASE_TRABAJO_ORIGINAL<-BASE_TRABAJO #salvamos la original para reutilizarla después
  
  
  
  
  
  
  #Para web quedarse con: Clasificacion==NO COMPARATIVA
  #y quitar: Nom_Sucursal==Möblum Tienda Empleados y Tiendita Empleados 
  #View(BASE_TRABAJO %>% filter(Clasificacion=="NO COMPARATIVA", Nom_Sucursal!="Tiendita Empleados" & Nom_Sucursal!="Möblum Tienda Empleados"))
  #BASE_TRABAJO<-BASE_TRABAJO %>% filter(Clasificacion=="NO COMPARATIVA", Nom_Sucursal!="Tiendita Empleados" & Nom_Sucursal!="Möblum Tienda Empleados")
  #REGIONES<-levels(as.factor(BASE_TRABAJO$Region))
  
  
  #3 Cargamos las Pestañas del Archivo que contiene los precios de lista e inventario
  #para limpiarlos
  #y quitar duplicados
  cat("Valores de preciolista y disp: ", preciolista[m], ", ", disp[m], "\n")
  if(preciolista[m]==1)
  {
    ListaPrecios<-as.data.frame(read_excel(NombresPrecios[m], sheet= "ListaPrecios"))#HAY QUE MODIFICAR EN CASO DE QUE HAYA MÁS DE UN ARCHIVO DE PRECIOS DE LISTA
    
    ListaPrecios<-ListaPrecios[which(!is.na(ListaPrecios$NuevoKit)), ]
    ListaPrecios$NuevoKit<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$NuevoKit)
    ListaPrecios$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$Descripcion)
    
    ListaPrecios<-QuitarMultiplicados_Simple1(ListaPrecios,"NuevoKit")
    
    #A la tabla ListaPrecios le creamos una nueva columna
    ListaPrecios$PrecioLista_Prom<-apply(ListaPrecios[,c(6:ncol(ListaPrecios))], MARGIN = 1, FUN = mean, na.rm= TRUE)
  }
  
  if(disp[m]==1)
  {
    Inventario<-as.data.frame(read_excel("R_CentroListaPreciosInventarioEne-Oct2021.xlsx", sheet= "Inventario"))#HAY QUE MODIFICAR
    
    Inventario<-Inventario[which(!is.na(Inventario$Articulo)), ]
    Inventario$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Articulo)
    Inventario$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Familia)
    Inventario$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Categoria)
    Inventario$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Descripcion)
    
    Inventario<-QuitarMultiplicados_Simple1(Inventario,"Articulo")
  }
  
  
  
  
  
  #4 PROGRAMA PRINCIPAL
  
  REGIONES<-levels(as.factor(BASE_TRABAJO$Region))
  cat("Regiones de la base: ", REGIONES, "\n")
  for(n in c(1: length(REGIONES) ))
  {
    
    cat("Corriendo el programa principal para la región: ", REGIONES[n],"\n")
    BASE_TRABAJO <- BASE_TRABAJO_ORIGINAL %>%  filter(Region==REGIONES[n])
    
    
    ZONAS<-levels(as.factor(BASE_TRABAJO$Zona))
    cat("Zonas de la base: ", ZONAS, "\n")
    for(z in c(1:length(ZONAS)))
    {
      contador<-contador+1
      cat("Corriendo el programa principal para
          \nRegión: ", REGIONES[n], "\nZona: ", ZONAS[z], "\nContador= ", contador, "\n")
      

      Resumen<-TablasResumenSimple(BASE_TRABAJO, "Utilidad", UltiMes[m])
      Resumen[["BaseOrigen"]]$Fecha <- as.Date(Resumen[["BaseOrigen"]]$Fecha, format = "%d/%m/%Y") #regresamos al formato de fecha estandar
      CATEGORIAS<-Resumen[["CategoriasOrdenadas"]] 
      Resumen[["CategoriasOrdenadas"]]<-gsub("[[:punct:]|\\+|¿|\\|¡]", " ", Resumen[["CategoriasOrdenadas"]])
      
      ListaCategoriasTotales<-lapply(CATEGORIAS, TablasPorCategoria, BASE_TRABAJO, UltiMes[m])
      names(ListaCategoriasTotales)<-CATEGORIAS
      ListaOBJETOS[[contador]]<-list(RESUMEN=Resumen, 
                                     LISTACATEGORIAS=ListaCategoriasTotales)
      
      if(preciolista[m]==1) #También se le puede agregar el parámetro de Disponibilidad
      {
        if(disp[m]==1 & !is.null(fechacorte[m]))
        {
          ListaCategoriasTotales<-lapply(ListaCategoriasTotales, Agregar_PrecioLista_Dispoibilidad, ListaPrecios, Inventario, fechacorte[m])
        }
        else
        {
          ListaCategoriasTotales<-lapply(ListaCategoriasTotales, Agregar_PrecioLista_Dispoibilidad, ListaPrecios)
        }
        
        ListaOBJETOS[[contador]]<-list(RESUMEN=Resumen, 
                                       LISTACATEGORIAS=ListaCategoriasTotales)
      }
      
      
      AnalisisGral<-TablasGral(ListaCategoriasTotales,Resumen[["MesActual"]], Disponibilidad = disp[m])
      
      
      
      #VEMOS SI SE UTILIZÓ SOLO UN INTERVALO DE LA FECHA PARA PERSONALIZAR MEJOR EL NOMBRE
      if(!is.null(FechaIntervalo))
      {
        NombrArch[contador]<-paste("Series_", REGIONES[n], "-", ZONAS[z], "_", as.character(FechaIntervalo[1]), "a",as.character(FechaIntervalo[2]), sep = " ")
        
      }else
      { #SI NO, PONEMOS UNO POR DEFAULT SEGÚN LA REGIÓN Y EL NOMBRE DE LA BASE
        NombrArch[contador]<-paste("Series_", REGIONES[n], "-", ZONAS[z], "_", str_sub(NombresBases[m], Loc_Inicio[m], Loc_Fin[m]), sep = " ")
      }
      
      
      #NombrArch[contador]<-paste("Series", REGIONES[n], str_sub(NombresBases[m], Loc_Inicio[m], Loc_Fin[m]), sep = " ")
      ListaOBJETOS[[contador]]<-list(RESUMEN=Resumen, 
                                     LISTACATEGORIAS=ListaCategoriasTotales, 
                                     ANALISISGRAL=AnalisisGral, 
                                     PRECIOLISTA=preciolista[m],
                                     DISPONIBILIDAD=disp[m],
                                     NOMBREARCHIVO=NombrArch[contador])
      names(ListaOBJETOS)[contador]<-NombrArch[contador]
      
      ExportarTablasExcelCompleto(Resumen= ListaOBJETOS[[contador]][["RESUMEN"]] , ListaTablasCategoria= ListaOBJETOS[[contador]][["LISTACATEGORIAS"]], ListaTablasGral= ListaOBJETOS[[contador]][["ANALISISGRAL"]], ListaOBJETOS[[contador]][["PRECIOLISTA"]], ListaOBJETOS[[contador]][["DISPONIBILIDAD"]], NombreArchivo= ListaOBJETOS[[contador]][["NOMBREARCHIVO"]])
      
    }
    
    
  }
  
  
}

m
n
z
contador


#APARTADO ESPECIAL
##################
ListaOBJETOS_Copia<-ListaOBJETOS

#luego se carga el archivo anterior
length(ListaOBJETOS)

#Se pegan las listas
ListaOBJETOS[c(4,5,6,7,8)] <- ListaOBJETOS_Copia
names(ListaOBJETOS)[c(4:8)] <- names(ListaOBJETOS_Copia)
##################


NombresBases
NombrArch


save(ListaOBJETOS, file = "Series_REGIONES_TOTALES_Ene-Dic_2019.RData")

save(ListaOBJETOS, file = "Series_REGIONES_TOTALES_Ene-May_2023.RData")

save(ListaOBJETOS, file = "Series_REGIONES_TOTALES_Ene_2023.RData")

save(ListaOBJETOS, file = "Series_REGIONES_CENTRO_LOMAS_2022.RData")

save(ListaOBJETOS, file = "Series_REGIONES_BAJ_NOR_SUR_2022.RData")

save(ListaOBJETOS, file = "Series_REGION_CENTRO_SinPueblaOaxaca_Ene-Dic_2022.RData")

save(ListaOBJETOS, file = "Series_REGION_CENTRO_PueblaOaxaca_Ene-Dic_2022.RData")

save(ListaOBJETOS, file = "Series_RegionCENTRO-ZonasTotales_2022-07-01_a_2023-06-30.RData")

######################################


#ORGANIZANDO EL DESM
#########################
ListaOBJETOS_BajioBajioQro_2021<-ListaOBJETOS_2
ListaOBJETOS_TodoCentro_2021_2022<-ListaOBJETOS

ListaOBJETOS_TodoMenosCentro_2021_2022<-ListaOBJETOS

names(ListaOBJETOS_TodoMenosCentro_2021_2022)
length(ListaOBJETOS_TodoMenosCentro_2021_2022)

for (n in c(1:10)) {
  
  names(ListaOBJETOS_TodoMenosCentro_2021_2022)[n]<-ListaOBJETOS_TodoMenosCentro_2021_2022[[n]][["NOMBREARCHIVO"]]
}




ListaOBJETOS_RegionesTotales_2021<-list()

ListaOBJETOS_RegionesTotales_2021[c(1:3)]<-ListaOBJETOS_TodoCentro_2021_2022[c(4:6)]
names(ListaOBJETOS_RegionesTotales_2021)[c(1:3)]<-names(ListaOBJETOS_TodoCentro_2021_2022[c(4:6)])

ListaOBJETOS_RegionesTotales_2021[c(4:5)]<-ListaOBJETOS_BajioBajioQro_2021
names(ListaOBJETOS_RegionesTotales_2021)[c(4:5)]<-names(ListaOBJETOS_BajioBajioQro_2021)

ListaOBJETOS_RegionesTotales_2021[c(6:8)]<-ListaOBJETOS_TodoMenosCentro_2021_2022[c(6,8,10)]
names(ListaOBJETOS_RegionesTotales_2021)[c(6:8)]<-names(ListaOBJETOS_TodoMenosCentro_2021_2022[c(6,8,10)])

View(ListaOBJETOS_RegionesTotales_2021)
save(ListaOBJETOS_RegionesTotales_2021, file = "Series_RegionesTotales_2021.RData")



ListaOBJETOS_RegionesTotales_2022<-list()

ListaOBJETOS_RegionesTotales_2022[c(1:3)]<-ListaOBJETOS_TodoCentro_2021_2022[c(1:3)]
names(ListaOBJETOS_RegionesTotales_2022)[c(1:3)]<-names(ListaOBJETOS_TodoCentro_2021_2022[c(1:3)])

ListaOBJETOS_RegionesTotales_2022[c(4:8)]<-ListaOBJETOS_TodoMenosCentro_2021_2022[c(1,2,5,7,9)]
names(ListaOBJETOS_RegionesTotales_2022)[c(4:8)]<-names(ListaOBJETOS_TodoMenosCentro_2021_2022[c(1,2,5,7,9)])


View(ListaOBJETOS_RegionesTotales_2022)
save(ListaOBJETOS_RegionesTotales_2022, file = "Series_RegionesTotales_2022.RData")
#########################




















#library(RODBC)
library(readr)
library(dplyr)
library(tidyverse)
library(htmlwidgets)
library(stringr)

library(readxl)
library(writexl)
library(openxlsx)



library("zoo")


#PARA RESUMEN SIMPLE

#1 (ACTIVADA)
ResumenSimpleVariable<-function(TABLA,VARIABLE,MONTHYEAR= NULL)#FORMATO DE MONTHYEAR MM/AAAA
{
  if(!is.null(MONTHYEAR))
  {
    TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  }
  
  
  if(VARIABLE=="Pzas")
  {
    Total<-TABLA %>% summarise(sum(Cantidad))
    x<-as.data.frame(TABLA 
                     %>% group_by(Region, Categoria) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad))
                     %>% arrange(desc(Pzas), desc(Utilidad)) )
    
  }
  if(VARIABLE=="Venta")
  {
    Total<-TABLA %>% summarise(sum(Venta))
    x<-as.data.frame(TABLA 
                     %>% group_by(Region, Categoria) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad))
                     %>% arrange(desc(Venta), desc(Utilidad)) )
  }
  if(VARIABLE=="Utilidad")
  {
    Total<-TABLA %>% summarise(sum(Utilidad))
    x<-as.data.frame(TABLA 
                     %>% group_by(Region, Categoria) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad))
                     %>% arrange(desc(Utilidad), desc(Venta)) )
  }

  return(x)
}
Tabla3<-ResumenSimpleVariable(CENTRO_EneSep_2021,REGIONES_CENTRO[1],"Utilidad")
Tabla3<-ResumenSimpleVariable(BASE_TRABAJO,"Utilidad", "03/2022")
View(Tabla3)
CATEGORIAS<-Tabla3[["Categoria"]]
class(Tabla3)
class(CATEGORIAS)

#2 (ACTIVADA)
TablasResumenSimple<-function(TABLA, VARIABLE, YEARMONTH)#YEARMONTH EN FORMATO MM/AAAA
{
  cat("Trabajando en el Resumen\n")
  Base<-as.data.frame( TABLA )
  
  #Análisis MES seleccionado
  TablaMes<-ResumenSimpleVariable(TABLA,VARIABLE,YEARMONTH)

  #Análisis ACUMULADO
  TablaAcumulado<-ResumenSimpleVariable(TABLA,VARIABLE)
  
  #Categorías en orden desc
  CategOrdenDesc<-as.vector(TablaAcumulado[,2])
  
  
  return(list(BaseOrigen=Base, MesActual=YEARMONTH, ResumenMes=TablaMes, ResumenAcum=TablaAcumulado, CategoriasOrdenadas=unique(CategOrdenDesc)))
}
x<-TablasResumenSimple(BASE_TRABAJO, "Utilidad", "03/2022")
x[[4]]
x[[5]]
unique(x[[5]])



#PARA RESUMEN 
#DESACTIVADA
ResumenPartVariable<-function(TABLA,REGION,VARIABLE,MONTHYEAR= NULL)#FORMATO DE MONTHYEAR MM/AAAA
{
    if(!is.null(MONTHYEAR))
    {
        TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
    }
    
    
    if(VARIABLE=="Pzas")
    {
        Total<-TABLA %>% filter(Region==REGION)%>% summarise(sum(Cantidad))
        x<-as.data.frame(TABLA %>% filter(Region==REGION) 
                         %>% group_by(Region, Categoria) 
                         %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/Total)
                         %>% arrange(desc(Pzas), desc(Utilidad)) )
        
    }
    if(VARIABLE=="Venta")
    {
        Total<-TABLA %>% filter(Region==REGION)%>% summarise(sum(Venta))
        x<-as.data.frame(TABLA %>% filter(Region==REGION) 
                         %>% group_by(Region, Categoria) 
                         %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/Total)
                         %>% arrange(desc(Venta), desc(Utilidad)) )
    }
    if(VARIABLE=="Utilidad")
    {
        Total<-TABLA %>% filter(Region==REGION)%>% summarise(sum(Utilidad))
        x<-as.data.frame(TABLA %>% filter(Region==REGION) 
                         %>% group_by(Region, Categoria) 
                         %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartUtilidad= 100*Utilidad/Total)
                         %>% arrange(desc(Utilidad), desc(Venta)) )
    }
    
    
    PartVARIABLE= paste0("Part",VARIABLE)
    names(x[,7])= PartVARIABLE
    ParticipAcumulada<-x[,PartVARIABLE]
    ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ]
    if(length(ParticipAcumulada)>=2)
    {
        for(n in c(2:length(ParticipAcumulada)))
        {
            ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
        }
        
        Clasif<-"A"
        for(n in c(2:length(ParticipAcumulada)))
        {
            if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.6 <= 50 )
            {
                Clasif<-c(Clasif,"A")
            }else{
                break
            }
        }
        for(m in c(n:length(ParticipAcumulada)))
        {
            if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.6 <= 80 )
            {
                Clasif<-c(Clasif,"B")
            }else{
                break
            }
        }
        Clasif[c(m:length(ParticipAcumulada))]<-"C"
        
    }else{
        Clasif<-"A"
    }
    
    
    if(VARIABLE=="Pzas")
    {
        if(nrow(x)==0)
        {
            x<-rbind(x, c(REGION,NA,NA,NA,NA,NA,NA))
            names(x)<-c("Categoria","Familia","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
            x$PartAcumulada <- NA
            x$ClasifPorPzas <- "_"
        }else{
            x$PartAcumulada <- NA
            x$ClasifPorPzas <- "_"
            
            x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
            x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
            
        }
        x<-x[,c(1:3,7,8,4,5,6,9)]
        
    }
    if(VARIABLE=="Venta")
    {
        if(nrow(x)==0)
        {
            x<-rbind(x, c(REGION,NA,NA,NA,NA,NA,NA))
            names(x)<-c("Categoria","Familia","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
            x$PartAcumulada <- NA
            x$ClasifPorVenta <- "_"
        }else{
            x$PartAcumulada <- NA
            x$ClasifPorVenta <- "_"
            
            x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
            x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
        }
        x<-x[,c(1:4,7,8,5,6,9)]
        
    }
    if(VARIABLE=="Utilidad")
    {
        if(nrow(x)==0)
        {
            x<-rbind(x, c(REGION,NA,NA,NA,NA,NA,NA))
            names(x)<-c("Categoria","Familia","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
            x$PartAcumulada <- NA
            x$ClasifPorUtilidad <- "_"
        }else{
            x$PartAcumulada <- NA
            x$ClasifPorUtilidad <- "_"
            
            x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
            x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
        }
        
        
    }
    
    
    #x[,c(4:7)]<- mapply(round, x[,c(4:7)], 0) 
    return(x)
}
Tabla3<-ResumenPartVariable(CENTRO_EneSep_2021,REGIONES[1],"Utilidad")
View(Tabla3)

#DESACTIVADA
ResumenClasifCategoriasMes<-function(MONTHYEAR,TABLA,REGION)#MONTHYEAR EN Formato MM/AAAA
{
    TABLA <- TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
    
    Tabla1<-ResumenPartVariable(TABLA,REGION,"Pzas")
    Tabla2<-ResumenPartVariable(TABLA,REGION,"Venta")
    Tabla3<-ResumenPartVariable(TABLA,REGION,"Utilidad")
    
    Tabla4<-Tabla3[,c(1,2)]
    Tabla4<-left_join(Tabla4, Tabla1[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4<-left_join(Tabla4, Tabla2[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4<-left_join(Tabla4, Tabla3[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4$ClasifABC<-paste0(Tabla4[,3],Tabla4[,4],Tabla4[,5])
    #Tabla4<-Tabla4[order(Tabla4$ClasifABC, decreasing = FALSE), ]#PARCHE
    
    #ListaMes<-list(Tabla1$ClasifPorPzas,Tabla2$ClasifPorVenta,Tabla3$ClasifPorUtilidad)
    #Tabla4$ClasifABC<-do.call(paste0,ListaMes)
    
    #print(unique(str_sub(TABLA$Fecha,4,10)))
    
    names(Tabla4)[6]= MONTHYEAR
    
    return(Tabla4[,c(2,6)])
}
View(ResumenClasifCategoriasMes("01/2021", CENTRO_EneSep_2021, REGIONES[1]))
rm(TablasResumen)

#DESACTIVADA
TablasResumen<-function(TABLA, REGION, YEARMONTH)#YEARMONTH EN FORMATO MM/AAAA
{
    cat("Trabajando en el Resumen\n")
    
    #Análisis MES seleccionado
    Tabla1Mes<-ResumenPartVariable(TABLA,REGION,"Pzas",YEARMONTH)
    Tabla2Mes<-ResumenPartVariable(TABLA,REGION,"Venta",YEARMONTH)
    Tabla3Mes<-ResumenPartVariable(TABLA,REGION,"Utilidad",YEARMONTH)

    Tabla4Mes<-Tabla3Mes[,c(1,2)]
    Tabla4Mes<-left_join(Tabla4Mes, Tabla1Mes[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla2Mes[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla3Mes[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4Mes$ClasifABC<-paste0(Tabla4Mes[,3],Tabla4Mes[,4],Tabla4Mes[,5])
    Tabla4Mes<-Tabla4Mes[,c(-3,-4,-5)]

    
    
    #Análisis ACUMULADO
    Tabla1Acumulado<-ResumenPartVariable(TABLA,REGION,"Pzas")
    Tabla2Acumulado<-ResumenPartVariable(TABLA,REGION,"Venta")
    Tabla3Acumulado<-ResumenPartVariable(TABLA,REGION,"Utilidad")
    CategOrdenDesc<-as.vector(Tabla3Acumulado[,2])
   
    Tabla4Acumulado<-Tabla3Acumulado[,c(1,2)]
    Tabla4Acumulado<-left_join(Tabla4Acumulado, Tabla1Acumulado[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4Acumulado<-left_join(Tabla4Acumulado, Tabla2Acumulado[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4Acumulado<-left_join(Tabla4Acumulado, Tabla3Acumulado[,c(2,9)], by=c("Categoria"="Categoria"))
    Tabla4Acumulado$ClasifABC<-paste0(Tabla4Acumulado[,3],Tabla4Acumulado[,4],Tabla4Acumulado[,5])
    Tabla4Acumulado<-Tabla4Acumulado[,c(-3,-4,-5)]
    #Tabla4<-Tabla4[order(Tabla4$ClasifABC, decreasing = FALSE), ] #PARCHE
  
    
    
    #Resumen Clasif de TODOS LOS MESES  
    MESES<-unique(str_sub(TABLA$Fecha,4,10))#MesAño
    #ANIOS<-unique(str_sub(TABLA$Fecha,7,10))#Año
   
    ListaMeses<-lapply(MESES, ResumenClasifCategoriasMes, TABLA, REGION)
    Tabla5<-Tabla4Acumulado[,c(1,2)]
    for(n in c(1:length(MESES)))
    {
        Tabla5<-left_join(Tabla5, ListaMeses[[n]])
    }
    Tabla5$Acumulada<-Tabla4Acumulado$ClasifABC
 
    
    
    return(ListaTotal<-list(YEARMONTH,Tabla1Mes,Tabla2Mes,Tabla3Mes,Tabla4Mes,Tabla1Acumulado,Tabla2Acumulado,Tabla3Acumulado,Tabla4Acumulado,Tabla5,CategOrdenDesc))
}
x<-TablasResumen(CENTRO_EneSep_2021,REGIONES[1], "09/2021")
x




















#PARA CATEGORÍAS (funciones antiguas)
#################################
#1 Arreglar posible bug de TotalCat 1
ClasificarFamilias<-function(x, VARIABLE, CATEGORIA, TotalCat)#TotalCat mod
{
  PartVARIABLE= paste0("Part",VARIABLE)
  
  ParticipAcumulada<-x[,PartVARIABLE]
  ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ]
  if(length(ParticipAcumulada)>=2 & TotalCat!=0)#TotalCat mod
  {
    for(n in c(2:length(ParticipAcumulada)))
    {
      ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
    }
    
    Clasif<-"A"
    for(n in c(2:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.6 <= 50 )
      {
        Clasif<-c(Clasif,"A")
      }else{
        break
      }
    }
    for(m in c(n:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.6 <= 80 )
      {
        Clasif<-c(Clasif,"B")
      }else{
        break
      }
    }
    Clasif[c(m:length(ParticipAcumulada))]<-"C"
    
  }else{
    ifelse(TotalCat!=0, Clasif<-"A", Clasif<-"_")#mod
    
    #Clasif<-"A"
  }
  
  
  if(VARIABLE=="Pzas")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Costo"=NA,"Utilidad"=NA,"Clasif"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorPzas <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","PartPzas","PartAcumulada","Venta","Costo","Utilidad","ClasifPorPzas")]
    }else{
      if(TotalCat!=0)
      {
        x$PartAcumulada <- NA
        x$ClasifPorPzas <- "_"
        
        x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
        x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
        x<-x[,c(1:3,7,8,4,5,6,9)]
      }else{#mod
        x$PartAcumulada <- NA
        x$ClasifPorPzas <- "_"
        
        x$PartAcumulada <- ParticipAcumulada
        x$ClasifPorPzas <- Clasif
        x<-x[,c(1:3,7,8,4,5,6,9)]
      }
      
    }
    
    
  }
  if(VARIABLE=="Venta")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Costo"=NA,"Utilidad"=NA,"Clasif"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorVenta <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","PartVenta","PartAcumulada","Costo","Utilidad","ClasifPorPzas")]
    }else{
      if(TotalCat!=0)
      {
        x$PartAcumulada <- NA
        x$ClasifPorVenta <- "_"
        
        x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
        x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
        x<-x[,c(1:4,7,8,5,6,9)]
      }else{#mod
        x$PartAcumulada <- NA
        x$ClasifPorVenta <- "_"
        
        x$PartAcumulada <- ParticipAcumulada
        x$ClasifPorVenta <- Clasif
        x<-x[,c(1:4,7,8,5,6,9)]
      }
      
    }
    
  }
  if(VARIABLE=="Utilidad")
  {
    if(nrow(x)==0)
    {
      
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"PartUtilidad"=NA,"PartAcumulada"=NA,"Clasif"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorUtilidad <- "_"
    }else{
      if(TotalCat!=0)
      {
        x$PartAcumulada <- NA
        x$ClasifPorUtilidad <- "_"
        
        x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
        x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      }else{
        x$PartAcumulada <- NA
        x$ClasifPorUtilidad <- "_"
        
        x$PartAcumulada <- ParticipAcumulada
        x$ClasifPorUtilidad <- Clasif
      }
      
    }
    
    
  }
  
  return(x)
}

#1 Arreglar posible bug de TotalCat 2 (Mejor)
ClasificarFamilias<-function(x, VARIABLE, CATEGORIA, TotalCat)#TotalCat mod
{
  PartVARIABLE= paste0("Part",VARIABLE)
  
  ParticipAcumulada<-x[,PartVARIABLE]
  ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ]
  if(length(ParticipAcumulada)>=2 & TotalCat!=0)#TotalCat mod
  {
    for(n in c(2:length(ParticipAcumulada)))
    {
      ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
    }
    
    Clasif<-"A"
    for(n in c(2:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.6 <= 50 )
      {
        Clasif<-c(Clasif,"A")
      }else{
        break
      }
    }
    for(m in c(n:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.6 <= 80 )
      {
        Clasif<-c(Clasif,"B")
      }else{
        break
      }
    }
    Clasif[c(m:length(ParticipAcumulada))]<-"C"
    
  }else{
    Clasif<-ifelse(TotalCat!=0, "A", "_")#mod
    
    #Clasif<-"A"
  }
  
  
  if(VARIABLE=="Pzas")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"PartPzas"=NA,"PartAcumulada"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"ClasifPorPzas"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorPzas <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","PartPzas","PartAcumulada","Venta","Costo","Utilidad","ClasifPorPzas")]
    }else{
      
        x$PartAcumulada <- NA
        x$ClasifPorPzas <- "_"
        
        x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
        x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
        x<-x[,c(1:3,7,8,4,5,6,9)]
      
    }
    
    
  }
  if(VARIABLE=="Venta")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Costo"=NA,"Utilidad"=NA,"ClasifPorVenta"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorVenta <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","PartVenta","PartAcumulada","Costo","Utilidad","ClasifPorPzas")]
    }else{
      
        x$PartAcumulada <- NA
        x$ClasifPorVenta <- "_"
        
        x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
        x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
        x<-x[,c(1:4,7,8,5,6,9)]
      
    }
    
  }
  if(VARIABLE=="Utilidad")
  {
    if(nrow(x)==0)
    {
      
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"PartUtilidad"=NA,"PartAcumulada"=NA,"ClasifPorUtilidad"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorUtilidad <- "_"
    }else{
      
        x$PartAcumulada <- NA
        x$ClasifPorUtilidad <- "_"
        
        x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
        x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      
    }
    
    
  }
  
  return(x)
}

#2 Arreglar posible bug de if(length)
ClasificarArticulosLocales<-function(x, VARIABLE)#antigua
{
  PartVARIABLE= paste0("Part",VARIABLE)
  
  ParticipAcumulada<-x[,PartVARIABLE]
  ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ]
  if(length(ParticipAcumulada)>=2)
  {
    for(n in c(2:length(ParticipAcumulada)))
    {
      ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
    }
    
    Clasif<-"a"
    for(n in c(2:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.55 <= 50 )
      {
        Clasif<-c(Clasif,"a")
      }else{
        break
      }
    }
    for(m in c(n:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.55 <= 80 )
      {
        Clasif<-c(Clasif,"b")
      }else{
        break
      }
    }
    Clasif[c(m:length(ParticipAcumulada))]<-"c"
    
  }else{
    Clasif<-"a"
  }
  
  
  if(VARIABLE=="Pzas")
  {
    x$PartAcumulada <- NA
    x$ClasifPorPzas <- "_"
    
    if( length(x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ]!=0) )
    {
      #cat("x$PartAcumulada\n", x$PartAcumulada, "\nParticipAcumulada\n", ParticipAcumulada)
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- ParticipAcumulada
      x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- Clasif
    }
    #x<-x[,c(1:3,7,8,4:6,9)]
    x<-x[,c(1:5,9,10,6:8,11)]
  }
  
  if(VARIABLE=="Venta")
  {
    x$PartAcumulada <- NA
    x$ClasifPorVenta <- "_"
    
    if( length(x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ]!=0) )
    {
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- ParticipAcumulada
      x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- Clasif
    }
    #x<-x[,c(1:4,7,8,5,6,9)]
    x<-x[,c(1:6,9,10,7,8,11)]
  }
  
  if(VARIABLE=="Utilidad")
  {
    x$PartAcumulada <- NA
    x$ClasifPorUtilidad <- "_"
    
    if( length(x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ]!=0) )
    {
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- ParticipAcumulada
      x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- Clasif
    }
  }
  
  return(x)
}



#3 Posible bug if(RENGLON_DF[,VARIABLE]==0)
TablaLocalPartVariable<-function(RENGLON_DF, TABLA, VARIABLE, MONTHYEAR= NULL)
{
  if(!is.null(MONTHYEAR))
  {
    TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  }
  
  RENGLON_DF<-data.frame(RENGLON_DF)
  #print(RENGLON_DF)
  AUX=FALSE
  if(VARIABLE=="Pzas" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/RENGLON_DF[,"Pzas"])
                     %>% arrange(desc(Pzas), desc(Utilidad)))
    AUX=TRUE
  }
  if(VARIABLE=="Venta" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/RENGLON_DF[,"Venta"])
                     %>% arrange(desc(Venta), desc(Utilidad)))
    AUX=TRUE
  }
  if(VARIABLE=="Utilidad" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartUtilidad= 100*Utilidad/RENGLON_DF[,"Utilidad"])
                     %>% arrange(desc(Utilidad), desc(Venta)))
    AUX=TRUE
  }
  
  
  
  #if( nrow(x)>0 )
  if( AUX )
  {
    #print("Entr??? al if")
    x$Categoria<-NA
    PartVARIABLE= paste0("Part",VARIABLE)
    names(x[,9])= PartVARIABLE
    
    #x<-x[,c(1:5,9,10,6:8,11)] Para VARIABLE=="Pzas"
    x<-ClasificarArticulosLocales(x,VARIABLE)
    if(RENGLON_DF[,VARIABLE]==0)
    {
      x[,PartVARIABLE]<-NA
    }
    
    RENGLON_DF$Articulo<-NA
    RENGLON_DF$Descripcion<-NA
    RENGLON_DF<-RENGLON_DF[,c(1,2,10,11,3:9)]
    
    return(rbind(RENGLON_DF,x))
    
  }else{
    return(RENGLON_DF)
  }
  
  
}


#4
TablaPartVariable<-function(TABLA,REGION,CATEGORIA,VARIABLE,MONTHYEAR= NULL)#FORMATO DE MONTHYEAR MM/AAAA
{
    if(!is.null(MONTHYEAR))
    {
        TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
    }
    
    
    if(VARIABLE=="Pzas")
    {
        Total<-TABLA %>% filter(Region==REGION, Categoria==CATEGORIA) %>% summarise(sum(Cantidad))
        x<-as.data.frame(TABLA %>% filter(Region==REGION, Categoria==CATEGORIA) 
                         %>% group_by(Categoria, Familia) 
                         %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/Total)
                         %>% arrange(desc(Pzas), desc(Utilidad)))
    }
    if(VARIABLE=="Venta")
    {
        Total<-TABLA %>% filter(Region==REGION, Categoria==CATEGORIA) %>% summarise(sum(Venta))
        x<-as.data.frame(TABLA %>% filter(Region==REGION, Categoria==CATEGORIA) 
                         %>% group_by(Categoria, Familia) 
                         %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/Total)
                         %>% arrange(desc(Venta), desc(Utilidad)))
    }
    if(VARIABLE=="Utilidad")
    {
        Total<-TABLA %>% filter(Region==REGION, Categoria==CATEGORIA) %>% summarise(sum(Utilidad))
        x<-as.data.frame(TABLA %>% filter(Region==REGION, Categoria==CATEGORIA) 
                         %>% group_by(Categoria, Familia) 
                         %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartUtilidad= 100*Utilidad/Total)
                         %>% arrange(desc(Utilidad), desc(Venta)))
    }
    
    PartVARIABLE= paste0("Part",VARIABLE)
    names(x[,7])= PartVARIABLE
    x<-ClasificarFamilias(x,VARIABLE, CATEGORIA, Total)#Total mod
    
    
    #PartVARIABLE= paste0("Part",VARIABLE)
    #prueba(TABLA,x,VARIABLE)
    
    #x[,c(4:7)]<- mapply(round, x[,c(4:7)], 0) 
    #x<-data.frame(do.call(rbind, lapply(split(x, f = x$Familia), prueba, TABLA, VARIABLE) ), row.names = NULL )
    
    
    # y<-list()
    # 
    # for(n in 1:nrow(x))
    # {
    #   y[n]<-prueba(x[n,], TABLA, VARIABLE)
    #   #y<-c(y,prueba(x[n,], TABLA, VARIABLE))
    # }
    # x<-data.frame(do.call(rbind, y), row.names = NULL)
    ListaFamilias<-lapply(transpose(x), TablaLocalPartVariable, BASE_TRABAJO, VARIABLE, MONTHYEAR)
    #View(p3[[2]])
    x<-do.call(rbind, ListaFamilias)
    #View(p3)
    
    # if(!is.null(x))
    # {
    #   return(x)
    # }else{
    #   x<-data.frame()
    # }
    
    
    return(x)
}
Tabla3<-TablaPartUtilidad(CENTRO_EneSep_2021,REGIONES[1],CATEGORIAS_CENTRO[1])
View(Tabla3)
Tabla3<-TablaPartVariable(BASE_TRABAJO,REGIONES[1],CATEGORIAS[1], "Pzas")
View(Tabla3)
Tabla4<-TablaPartVariable(BASE_TRABAJO,REGIONES[1],CATEGORIAS[1], "Utilidad")
View(Tabla4)

Tabla4<-TablaPartVariable(BASE_TRABAJO,REGIONES[2],CATEGORIAS[1], "Utilidad","03/2022")
View(Tabla4)
#################################

#PARA CATEGORÍAS (FIXED)
##############
#1 (ACTIVA) Arreglar posible bug de TotalCat 2 (Mejor)
ClasificarFamilias<-function(x, VARIABLE, CATEGORIA, TotalCat)#TotalCat mod
{
  PartVARIABLE= paste0("Part",VARIABLE)
  
  ParticipAcumulada<-x[,PartVARIABLE]
  ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ]
  if(length(ParticipAcumulada)>=2 & TotalCat!=0)#TotalCat mod
  {
    for(n in c(2:length(ParticipAcumulada)))
    {
      ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
    }
    
    Clasif<-"A"
    for(n in c(2:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.6 <= 50 )
      {
        Clasif<-c(Clasif,"A")
      }else{
        break
      }
    }
    for(m in c(n:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.6 <= 80 )
      {
        Clasif<-c(Clasif,"B")
      }else{
        break
      }
    }
    Clasif[c(m:length(ParticipAcumulada))]<-"C"
    
  }else{
    Clasif<-ifelse(TotalCat!=0, "A", "_")#mod
    
    #Clasif<-"A"
  }
  
  
  if(VARIABLE=="Pzas")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"PartPzas"=NA,"PartAcumulada"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"ClasifPorPzas"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorPzas <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","PartPzas","PartAcumulada","Venta","Costo","Utilidad","ClasifPorPzas")]
    }else{
      
      x$PartAcumulada <- NA
      x$ClasifPorPzas <- "_"
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
      x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      x<-x[,c(1:3,7,8,4,5,6,9)]
      
    }
    
    
  }
  if(VARIABLE=="Venta")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Costo"=NA,"Utilidad"=NA,"ClasifPorVenta"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorVenta <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","PartVenta","PartAcumulada","Costo","Utilidad","ClasifPorPzas")]
    }else{
      
      x$PartAcumulada <- NA
      x$ClasifPorVenta <- "_"
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
      x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      x<-x[,c(1:4,7,8,5,6,9)]
      
    }
    
  }
  if(VARIABLE=="Utilidad")
  {
    if(nrow(x)==0)
    {
      
      X<-data.frame("Categoria"=CATEGORIA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"PartUtilidad"=NA,"PartAcumulada"=NA,"ClasifPorUtilidad"=NA)
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorUtilidad <- "_"
    }else{
      
      x$PartAcumulada <- NA
      x$ClasifPorUtilidad <- "_"
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
      x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      
    }
    
    
  }
  #print("2.- x:")
  #print(x)
  return(x)
}

#2 (ACTIVA)  Arreglar posible bug de if(length) ya se mejoró
ClasificarArticulosLocales<-function(x, VARIABLE)
{
  PartVARIABLE= paste0("Part",VARIABLE)
  
  ParticipAcumulada<-x[,PartVARIABLE]
  ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ]
  if(length(ParticipAcumulada)>=2)
  {
    for(n in c(2:length(ParticipAcumulada)))
    {
      ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
    }
    
    Clasif<-"a"
    for(n in c(2:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.55 <= 50 )
      {
        Clasif<-c(Clasif,"a")
      }else{
        break
      }
    }
    for(m in c(n:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.55 <= 80 )
      {
        Clasif<-c(Clasif,"b")
      }else{
        break
      }
    }
    Clasif[c(m:length(ParticipAcumulada))]<-"c"
    
  }else{
    Clasif<-"a"
  }
  
  
  if(VARIABLE=="Pzas")
  {
    x$PartAcumulada <- NA
    x$ClasifPorPzas <- "_"
    
    if( length(x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ])!=0 )
    {
      #cat("x$PartAcumulada\n", x$PartAcumulada, "\nParticipAcumulada\n", ParticipAcumulada)
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- ParticipAcumulada
      x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- Clasif
    }
    #x<-x[,c(1:3,7,8,4:6,9)]
    x<-x[,c(1:5,9,10,6:8,11)]
  }
  
  if(VARIABLE=="Venta")
  {
    x$PartAcumulada <- NA
    x$ClasifPorVenta <- "_"
    
    if( length(x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ])!=0 )
    {
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- ParticipAcumulada
      x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- Clasif
    }
    #x<-x[,c(1:4,7,8,5,6,9)]
    x<-x[,c(1:6,9,10,7,8,11)]
  }
  
  if(VARIABLE=="Utilidad")
  {
    x$PartAcumulada <- NA
    x$ClasifPorUtilidad <- "_"
    
    if( length(x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ])!=0 )
    {
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- ParticipAcumulada
      x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 & x[,VARIABLE]!=Inf & x[,PartVARIABLE]!=Inf) ] <- Clasif
    }
  }
  
  return(x)
}


#3 (ACTIVA)   Posible bug if(RENGLON_DF[,VARIABLE]==0)
TablaLocalPartVariable<-function(RENGLON_DF, TABLA, VARIABLE, MONTHYEAR= NULL)
{ 
  if(!is.null(MONTHYEAR))
  {
    TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  }
  
  RENGLON_DF<-data.frame(RENGLON_DF)
  
  AUX=FALSE
  if(VARIABLE=="Pzas" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/RENGLON_DF[,"Pzas"])
                     %>% arrange(desc(Pzas), desc(Utilidad)))
    AUX=TRUE
  }
  if(VARIABLE=="Venta" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/RENGLON_DF[,"Venta"])
                     %>% arrange(desc(Venta), desc(Utilidad)))
    AUX=TRUE
  }
  if(VARIABLE=="Utilidad" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartUtilidad= 100*Utilidad/RENGLON_DF[,"Utilidad"])
                     %>% arrange(desc(Utilidad), desc(Venta)))
    AUX=TRUE
  }

  #if( nrow(x)>0 )
  if( AUX )
  {
    #print("Entr??? al if")
    x$Categoria<-NA
    #PartVARIABLE= paste0("Part",VARIABLE)
    #names(x[,9])= PartVARIABLE
    
    #x<-x[,c(1:5,9,10,6:8,11)] Para VARIABLE=="Pzas"
    x<-ClasificarArticulosLocales(x,VARIABLE)
    if(RENGLON_DF[,VARIABLE]==0)
    {
      x[,paste0("Part",VARIABLE)]<-NA #Modificado el 22/05/2022, estaba x[,PartVARIABLE]<-NA
    }
    
    RENGLON_DF$Articulo<-NA
    RENGLON_DF$Descripcion<-NA
    RENGLON_DF<-RENGLON_DF[,c(1,2,10,11,3:9)]
    
    return(rbind(RENGLON_DF,x))
    
  }else{
    return(RENGLON_DF)
  }
  
  
}

#4 (ACTIVA)
TablaPartVariable<-function(TABLA,CATEGORIA,VARIABLE,MONTHYEAR= NULL)#FORMATO DE MONTHYEAR MM/AAAA
{
  if(!is.null(MONTHYEAR))
  {
    TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  }
  
  
  if(VARIABLE=="Pzas")
  {
    Total<-as.numeric( TABLA %>% filter(Categoria==CATEGORIA) %>% summarise(sum(Cantidad)) )
    x<-as.data.frame(TABLA %>% filter(Categoria==CATEGORIA) 
                     %>% group_by(Categoria, Familia) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/Total)
                     %>% arrange(desc(Pzas), desc(Utilidad)))
  }
  if(VARIABLE=="Venta")
  {
    Total<-as.numeric( TABLA %>% filter(Categoria==CATEGORIA) %>% summarise(sum(Venta)) )
    x<-as.data.frame(TABLA %>% filter(Categoria==CATEGORIA) 
                     %>% group_by(Categoria, Familia) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/Total)
                     %>% arrange(desc(Venta), desc(Utilidad)))
  }
  if(VARIABLE=="Utilidad")
  {
    Total<-as.numeric( TABLA %>% filter(Categoria==CATEGORIA) %>% summarise(sum(Utilidad)) )
    x<-as.data.frame(TABLA %>% filter(Categoria==CATEGORIA) 
                     %>% group_by(Categoria, Familia) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartUtilidad= 100*Utilidad/Total)
                     %>% arrange(desc(Utilidad), desc(Venta)))
  }
  
  #PartVARIABLE= paste0("Part",VARIABLE)
  #names(x[,7])= PartVARIABLE
  x<-ClasificarFamilias(x,VARIABLE, CATEGORIA, Total)#Total mod
  
  #x[,c(4:7)]<- mapply(round, x[,c(4:7)], 0) 
  #x<-data.frame(do.call(rbind, lapply(split(x, f = x$Familia), prueba, TABLA, VARIABLE) ), row.names = NULL )
  
  
  # y<-list()
  # 
  # for(n in 1:nrow(x))
  # {
  #   y[n]<-prueba(x[n,], TABLA, VARIABLE)
  #   #y<-c(y,prueba(x[n,], TABLA, VARIABLE))
  # }
  # x<-data.frame(do.call(rbind, y), row.names = NULL)
  
  ListaFamilias<-lapply(transpose(x), TablaLocalPartVariable, TABLA, VARIABLE, MONTHYEAR)
  x<-do.call(rbind, ListaFamilias)
  #View(p3)
  
  # if(!is.null(x))
  # {
  #   return(x)
  # }else{
  #   x<-data.frame()
  # }
  
  
  return(x)
}

#Pruebas fallidas
############################
#3 (NUEVA, prueba menos rápida)   Posible bug if(RENGLON_DF[,VARIABLE]==0)
TablaLocalPartVariable<-function(RENGLON_DF, TABLA, VARIABLE, MONTHYEAR= NULL)
{
  if(!is.null(MONTHYEAR))
  {
    TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  }
  
  RENGLON_DF<-data.frame(RENGLON_DF)
  #print(RENGLON_DF)
  AUX=FALSE
  if(VARIABLE=="Pzas" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/RENGLON_DF[,"Pzas"])
                     %>% arrange(desc(Pzas), desc(Utilidad)))
    AUX=TRUE
  }
  if(VARIABLE=="Venta" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/RENGLON_DF[,"Venta"])
                     %>% arrange(desc(Venta), desc(Utilidad)))
    AUX=TRUE
  }
  if(VARIABLE=="Utilidad" & !is.na(RENGLON_DF[,"Familia"]))
  {
    x<-as.data.frame(TABLA %>% filter(Categoria==RENGLON_DF[,"Categoria"] & Familia==RENGLON_DF[,"Familia"])
                     %>% group_by(Categoria, Familia, Articulo, Descripcion) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad))
                     %>% mutate(PartUtilidad= 100*Utilidad/RENGLON_DF[,"Utilidad"])
                     %>% arrange(desc(Utilidad), desc(Venta)))
    AUX=TRUE
  }
  
  
  
  #if( nrow(x)>0 )
  if( AUX )
  {
    #print("Entr??? al if")
    x$Categoria<-NA
    #PartVARIABLE= paste0("Part",VARIABLE) #quitado el 22/05/2022
    #names(x[,9])= PartVARIABLE #quitado el 22/05/2022
    
    #x<-x[,c(1:5,9,10,6:8,11)] Para VARIABLE=="Pzas"
    x<-ClasificarArticulosLocales(x,VARIABLE)
    if(RENGLON_DF[,VARIABLE]==0)
    {
      x[,paste0("Part",VARIABLE)]<-NA 
    }
    
    RENGLON_DF$Articulo<-NA
    RENGLON_DF$Descripcion<-NA
    RENGLON_DF<-RENGLON_DF[,c(1,2,10,11,3:9)]
    
    return(rbind(RENGLON_DF,x))
    
  }else{
    return(RENGLON_DF)
  }
  
  
}

#4 (NUEVA, prueba menos rápida)
TablaPartVariable<-function(TABLA,CATEGORIA,VARIABLE,MONTHYEAR= NULL)#FORMATO DE MONTHYEAR MM/AAAA
{
  if(!is.null(MONTHYEAR))
  {
    TABLA<-TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  }
  
  
  if(VARIABLE=="Pzas")
  {
    Total<-as.numeric( TABLA %>% filter(Categoria==CATEGORIA) %>% summarise(sum(Cantidad)) )
    x<-as.data.frame(TABLA %>% filter(Categoria==CATEGORIA) 
                     %>% group_by(Categoria, Familia) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartPzas= 100*Pzas/Total)
                     %>% arrange(desc(Pzas), desc(Utilidad)))
  }
  if(VARIABLE=="Venta")
  {
    Total<-as.numeric( TABLA %>% filter(Categoria==CATEGORIA) %>% summarise(sum(Venta)) )
    x<-as.data.frame(TABLA %>% filter(Categoria==CATEGORIA) 
                     %>% group_by(Categoria, Familia) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad), PartVenta= 100*Venta/Total)
                     %>% arrange(desc(Venta), desc(Utilidad)))
  }
  if(VARIABLE=="Utilidad")
  {
    Total<-as.numeric( TABLA %>% filter(Categoria==CATEGORIA) %>% summarise(sum(Utilidad)) )
    x<-as.data.frame(TABLA %>% filter(Categoria==CATEGORIA) 
                     %>% group_by(Categoria, Familia) 
                     %>% summarise( Pzas= sum(Cantidad), Venta= sum(Venta), Costo= sum(Costo), Utilidad= sum(Utilidad))
                     %>% mutate(PartUtilidad= 100*Utilidad/Total)
                     %>% arrange(desc(Utilidad), desc(Venta)))
  }
  
  #PartVARIABLE= paste0("Part",VARIABLE)
  #names(x[,7])= PartVARIABLE
  x<-ClasificarFamilias(x,VARIABLE, CATEGORIA, Total)#Total mod
  
  
  #PartVARIABLE= paste0("Part",VARIABLE)
  #prueba(TABLA,x,VARIABLE)
  
  #x[,c(4:7)]<- mapply(round, x[,c(4:7)], 0) 
  #x<-data.frame(do.call(rbind, lapply(split(x, f = x$Familia), prueba, TABLA, VARIABLE) ), row.names = NULL )
  
  
  # y<-list()
  # 
  # for(n in 1:nrow(x))
  # {
  #   y[n]<-prueba(x[n,], TABLA, VARIABLE)
  #   #y<-c(y,prueba(x[n,], TABLA, VARIABLE))
  # }
  # x<-data.frame(do.call(rbind, y), row.names = NULL)
  ListaFamilias<-lapply(transpose(x), TablaLocalPartVariable, TABLA, VARIABLE, MONTHYEAR)
  #View(p3[[2]])
  x<-do.call(rbind, ListaFamilias)
  #View(p3)
  
  # if(!is.null(x))
  # {
  #   return(x)
  # }else{
  #   x<-data.frame()
  # }
  
  
  return(x)
}


cronometro <- proc.time() # Inicia el cronómetro
# NUESTRO CODIGO
Tabla4<-TablaPartVariable(BASE_TRABAJO,"SALAS", "Utilidad","03/2022")
proc.time()-cronometro    # Detiene el cronómetro

viejo
user  system elapsed 
3.14    0.06    3.16 
2.86    0.02    2.87
2.85    0.03    2.86
3.11    0.02    3.14
2.94    0.00    2.96
3.03    0.03    3.02
nuevo
user  system elapsed 
3.05    0.02    3.07
3.38    0.02    3.39
3.17    0.02    3.19 
3.49    0.13    3.63
3.16    0.05    3.27
3.22    0.03    3.29
############################

Tabla4<-TablaPartVariable(BASE_TRABAJO,"DECORACION", "Pzas","08/2022")
View(Tabla4)
rm(Tabla4)

BASE_TRABAJO_CENTRO<-BASE_TRABAJO %>% filter(Region=="DICO CENTRO")
Tabla4<-TablaPartVariable(BASE_TRABAJO_CENTRO,CATEGORIAS[1], "Utilidad","03/2022")
View(Tabla4)

BASE_TRABAJO_EXPO<-BASE_TRABAJO %>% filter(Region=="EXPO")
Tabla4<-TablaPartVariable(BASE_TRABAJO_EXPO,CATEGORIAS[1], "Utilidad","03/2022")
View(Tabla4)

BASE_TRABAJO_MOBLUM<-BASE_TRABAJO %>% filter(Region=="MÖBLUM")
Tabla4<-TablaPartVariable(BASE_TRABAJO_MOBLUM,CATEGORIAS[1], "Utilidad","03/2022")
View(Tabla4)

#################


#PRUEBAS
############################
la<-list()
la[1]<-1
la[2]<-"a"
la

View(ListaCategoriasTotales[[1]][[2]])


Tabla3<-TablaPartVenta(CENTRO_EneSep_2021,REGIONES[1],CATEGORIAS_CENTRO[1])
View(Tabla3)
Tabla3<-TablaPartVariable(CENTRO_EneSep_2021,REGIONES[1],CATEGORIAS_CENTRO[1], "Venta")
View(Tabla3)
Tabla3<-TablaPartPzas(CENTRO_EneSep_2021,REGIONES[1],CATEGORIAS_CENTRO[1])
View(Tabla3)
Tabla3<-TablaPartVariable(CENTRO_EneSep_2021,REGIONES[1],CATEGORIAS_CENTRO[1], "Pzas", "01/2021")
View(Tabla3)

Tabla3<-Tabla3[c(1:10),]
Tabla3
as.list(Tabla3)
as.list(Tabla3, )
apply(Tabla3, MARGIN = 1, print)

PegarTodo<-function(renglon)
{
  do.call(paste0, as.list(renglon))
}

renglon<-c("a","b","c","d")
do.call(paste0, as.list(renglon))

PegarTodo(renglon)

renglon<-Tabla3[1,]
renglon
renglon[,"Familia"]
class(renglon)
do.call(paste0, renglon)

Tabla3
apply(Tabla3, MARGIN = 1, PegarTodo)

renglon<-apply(Tabla3, MARGIN = 1, PegarTodo)
renglon[[1]]
class(renglon)

p3<-apply(Tabla3, MARGIN = 1, identidad)
p3[1]

identidad<-function(x)
{
  return(x)
}

#library("plyr") 
#rbind.fill(data_plyr1, data_plyr2) 







p3<-lapply(Tabla3[,c(1:9)], prueba, BASE_TRABAJO)

as.list(Tabla3[,c(1:9)])

p1<-prueba(renglon,BASE_TRABAJO)
p1
View(p1)


Tabla3<-Tabla3[c(1:3),]
Tabla3
p2<-apply(X= Tabla3, MARGIN = 1, FUN= prueba, TABLA=BASE_TRABAJO)

#1
p2<-lapply(split(Tabla3, f = Tabla3$Familia), prueba, BASE_TRABAJO)
p2
View(p2[[2]])

p2<-split(Tabla3, f = Tabla3$Familia)
p2
class(p2)

merge.default

p3<-do.call(rbind,p2)

#2
p3<-data.frame(do.call(rbind,p2), row.names = NULL)
class(p3)
View(p3)

rbind()


p3<-data.frame(do.call(rbind, lapply(split(Tabla3, f = Tabla3$Familia), prueba, BASE_TRABAJO) ), row.names = NULL )
p3<-data.frame(p3,row.names = NULL)
View(p3)




View(Tabla3)



as.list.data.frame(Tabla3[row(Tabla3),])
p3<-apply(transpose(p2), MARGIN = 1, prueba, BASE_TRABAJO,"Pzas")
p2<-Tabla3[c(1:3),]

length(transpose(p2))
class(transpose(p2))

data.frame(transpose(p2)[[1]])


p3<-lapply(transpose(p2), prueba, BASE_TRABAJO,"Pzas")
View(p3[[2]])
p3<-do.call(rbind, p3)
View(p3)






############################



as.character( as.yearmon("01/2021", format="%m/%Y") )
#5 ACTIVADA (ya modificada, incluye MU)
ResumenClasifFamiliasMes<-function(MONTHYEAR,CATEGORIA,TABLA)#MONTHYEAR EN Formato MM/AAAA
{
  cat(CATEGORIA, MONTHYEAR, sep = " ")
  TABLA <- TABLA %>% filter( str_sub(TABLA$Fecha,4,10)==MONTHYEAR )
  
  Tabla1<-TablaPartVariable(TABLA,CATEGORIA,"Pzas", MONTHYEAR)
  Tabla2<-TablaPartVariable(TABLA,CATEGORIA,"Venta", MONTHYEAR)
  Tabla3<-TablaPartVariable(TABLA,CATEGORIA,"Utilidad", MONTHYEAR)
  
  Tabla4<-Tabla3[,c("Categoria","Familia","Articulo","Descripcion")]
  # print("Tabla4")
  # print(Tabla4)
  
  if(!is.null(Tabla4))
  {
    Tabla4<-left_join(Tabla4, Tabla1[,c("Familia","Articulo","Descripcion","Pzas","PartPzas","ClasifPorPzas")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4<-left_join(Tabla4, Tabla2[,c("Familia","Articulo","Descripcion","Venta","PartVenta","ClasifPorVenta")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4<-left_join(Tabla4, Tabla3[,c("Familia","Articulo","Descripcion","Utilidad","PartUtilidad","ClasifPorUtilidad")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4$MU <-100*(Tabla4$Utilidad/Tabla4$Venta)
    Tabla4$ClasifABC<-paste0(Tabla4[,"ClasifPorPzas"],Tabla4[,"ClasifPorVenta"],Tabla4[,"ClasifPorUtilidad"])
    
    Tabla4<-Tabla4[,c("Familia","Articulo","Descripcion","ClasifABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")]
    names(Tabla4)<-c("Familia", "Articulo", "Descripcion", paste("ABC",MONTHYEAR,sep = "_"), paste("Pzas",MONTHYEAR,sep = "_"), paste("Venta",MONTHYEAR,sep = "_"), paste("Utilidad",MONTHYEAR,sep = "_"), paste("MU",MONTHYEAR,sep = "_"), paste("PartPzas", MONTHYEAR,sep = "_"), paste("PartVenta",MONTHYEAR,sep= "_"), paste("PartUtilidad",MONTHYEAR,sep = "_"))
    return(Tabla4)
    
  }else{
    Tabla4<-data.frame("Familia"=NA, "Articulo"=NA, "Descripcion"=NA, "ClasifABC"=NA, "Pzas"=NA, "Venta"=NA, "Utilidad"=NA, "MU"=NA, "PartPzas"=NA, "PartVenta"=NA, "PartUtilidad"=NA)
    names(Tabla4)<-c("Familia", "Articulo", "Descripcion", paste("ABC",MONTHYEAR,sep = "_"), paste("Pzas",MONTHYEAR,sep = "_"), paste("Venta",MONTHYEAR,sep = "_"), paste("Utilidad",MONTHYEAR,sep = "_"), paste("MU",MONTHYEAR,sep = "_"), paste("PartPzas", MONTHYEAR,sep = "_"), paste("PartVenta",MONTHYEAR,sep= "_"), paste("PartUtilidad",MONTHYEAR,sep = "_"))
    return(Tabla4)
    
  }
  
}
yy<-ResumenClasifFamiliasMes("03/2022", CATEGORIAS[1], BASE_TRABAJO_EXPO)
View(yy)
rm(yy)

#6 ACTIVADA (ya modificada, incluye MU)
TablasPorCategoria<-function(CATEGORIA,TABLA,YEARMONTH)#YEARMONTH EN FORMATO MM/AAAA
{
  cat("Trabajando en la Categoría: ", CATEGORIA, "\n")
  
  #Análisis MES seleccionado
  Tabla1Mes<-TablaPartVariable(TABLA,CATEGORIA,"Pzas",YEARMONTH)
  Tabla2Mes<-TablaPartVariable(TABLA,CATEGORIA,"Venta",YEARMONTH)
  Tabla3Mes<-TablaPartVariable(TABLA,CATEGORIA,"Utilidad",YEARMONTH)
  # print("Tabla1Mes= ")
  # print(Tabla1Mes)
  # print("Tabla2Mes= ")
  # print(Tabla2Mes)
  # print("Tabla3Mes= ")
  # print(Tabla3Mes)
  
  Tabla4Mes<-Tabla3Mes[,c("Categoria","Familia","Articulo","Descripcion")]
  # print("Tabla4Mes= ")    
  # print(Tabla4Mes)
  # 
  
  if(!is.null(Tabla4Mes))
  {
    Tabla4Mes<-left_join(Tabla4Mes, Tabla1Mes[,c("Familia","Articulo","Descripcion","Pzas","PartPzas","ClasifPorPzas")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla2Mes[,c("Familia","Articulo","Descripcion","Venta","PartVenta","ClasifPorVenta")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla3Mes[,c("Familia","Articulo","Descripcion","Utilidad","PartUtilidad","ClasifPorUtilidad")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes$MU <-100*(Tabla4Mes$Utilidad/Tabla4Mes$Venta) 
    Tabla4Mes$ClasifABC<-paste0(Tabla4Mes[,"ClasifPorPzas"],Tabla4Mes[,"ClasifPorVenta"],Tabla4Mes[,"ClasifPorUtilidad"])
    #Tabla4Mes<-Tabla4Mes[,c(-3,-4,-5)]
    Tabla4Mes<-Tabla4Mes[,c("Categoria","Familia","Articulo","Descripcion","ClasifABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")]
  }else{
    Tabla1Mes<-data.frame("Categoria"=CATEGORIA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"Pzas"=NA,"PartPzas"=NA,"PartAcumulada"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"ClasifABC"=NA)
    Tabla2Mes<-data.frame("Categoria"=CATEGORIA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Costo"=NA,"Utilidad"=NA,"ClasifABC"=NA)
    Tabla3Mes<-data.frame("Categoria"=CATEGORIA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"Costo"=NA,"Utilidad"=NA,"PartUtilidad"=NA,"PartAcumulada"=NA,"ClasifABC"=NA)
    Tabla4Mes<-data.frame("Categoria"=CATEGORIA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"ClasifABC"=NA,"Pzas"=NA,"Venta"=NA,"Utilidad"=NA,"MU"=NA,"PartPzas"=NA,"PartVenta"=NA,"PartUtilidad"=NA)
  }
  
  
  
  
  
  #Análisis ACUMULADO
  #print("Analisis AcumulAdo")
  
  Tabla1Acumulado<-TablaPartVariable(TABLA,CATEGORIA,"Pzas")
  Tabla2Acumulado<-TablaPartVariable(TABLA,CATEGORIA,"Venta")
  Tabla3Acumulado<-TablaPartVariable(TABLA,CATEGORIA,"Utilidad")
  
  # print("Tabla1Acumulado= ")
  # print(Tabla1Acumulado)
  # print("Tabla2Acumulado= ")
  # print(Tabla2Acumulado)
  # print("Tabla3Acumulado= ")
  # print(Tabla3Acumulado)        
  #     
  Tabla4Acumulado<-Tabla3Acumulado[,c("Categoria","Familia","Articulo","Descripcion")]
  # print("Tabla4Acuulado")
  # print(Tabla4Acumulado)
  
  Tabla4Acumulado<-left_join(Tabla4Acumulado, Tabla1Acumulado[,c("Familia","Articulo","Descripcion","Pzas","PartPzas","ClasifPorPzas")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  Tabla4Acumulado<-left_join(Tabla4Acumulado, Tabla2Acumulado[,c("Familia","Articulo","Descripcion","Venta","PartVenta","ClasifPorVenta")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  Tabla4Acumulado<-left_join(Tabla4Acumulado, Tabla3Acumulado[,c("Familia","Articulo","Descripcion","Utilidad","PartUtilidad","ClasifPorUtilidad")], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  Tabla4Acumulado$MU <-100*(Tabla4Acumulado$Utilidad/Tabla4Acumulado$Venta) 
  Tabla4Acumulado$ClasifABC<-paste0(Tabla4Acumulado[,"ClasifPorPzas"],Tabla4Acumulado[,"ClasifPorVenta"],Tabla4Acumulado[,"ClasifPorUtilidad"])
  #Tabla4Acumulado<-Tabla4Acumulado[,c(-3,-4,-5)]
  #Tabla4<-Tabla4[order(Tabla4$ClasifABC, decreasing = FALSE), ] #PARCHE
  Tabla4Acumulado<-Tabla4Acumulado[,c("Categoria","Familia","Articulo","Descripcion","ClasifABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")]
  
  
  
  #Resumen Clasif de TODOS LOS MESES  
  MESES<-unique(str_sub(TABLA$Fecha,4,10))#MesAño
  #ANIOS<-unique(str_sub(TABLA$Fecha,7,10))#Año
  ListaMeses<-lapply(MESES, ResumenClasifFamiliasMes, CATEGORIA, TABLA)
  #Para optimizar más adelante
  #which(c("01/2022","02/2022","03/2022","04/2022")=="03/2022")
  Tabla5<-Tabla4Acumulado[,c("Categoria","Familia","Articulo","Descripcion")]
  for(n in c(1:length(MESES)))
  {
    Tabla5<-left_join(Tabla5, ListaMeses[[n]], by=c("Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    #names(Tabla5)[c((2+4*n):(4+4*n))]<-c(paste("PartPzas",MESES[n],sep = "_"), paste("PartVenta",MESES[n],sep = "_"), paste("PartUtilidad",MESES[n],sep = "_"))
    #names(Tabla5)[c((2+4*n):(4+4*n))]<-c(paste("PartPzas",MESES[n],sep = "_"), paste("PartVenta",MESES[n],sep = "_"), paste("PartUtilidad",MESES[n],sep = "_"))
  }
  Tabla5$ClasifABC_Acum<-Tabla4Acumulado$ClasifABC
  Tabla5$Pzas_Acum<-Tabla4Acumulado$Pzas
  Tabla5$Venta_Acum<-Tabla4Acumulado$Venta
  Tabla5$Utilidad_Acum<-Tabla4Acumulado$Utilidad
  Tabla5$MU_Acum<-Tabla4Acumulado$MU
  Tabla5$PartPzas_Acumulada<-Tabla4Acumulado$PartPzas
  Tabla5$PartVenta_Acumulada<-Tabla4Acumulado$PartVenta
  Tabla5$PartUtilidad_Acumulada<-Tabla4Acumulado$PartUtilidad
  #Tabla5<-left_join(Tabla4, do.call(full_join,ListaMeses, by=c("Familia"=="Familia")), by=c("Familia"="Familia"))
  
  
  
  return(list(MesActual=YEARMONTH, TablaPzasMes=Tabla1Mes, TablaVentaMes=Tabla2Mes, TablaUtilidadMes=Tabla3Mes, ResultadosMes=Tabla4Mes, TablaPzasAcum=Tabla1Acumulado, TablaVentaAcum=Tabla2Acumulado, TablaUtilidadAcum=Tabla3Acumulado ,ResultadosAcum=Tabla4Acumulado ,SerieResultadosMensuales=Tabla5))
}
y<-TablasPorCategoria(CATEGORIAS[1],BASE_TRABAJO_EXPO, "03/2022")
View(y[["ResultadosMes"]])
View(y[["SerieResultadosMensuales"]])
rm(y)







#PARA LOS CRUCES
#1 ACTIVADA
QuitarMultiplicados_Simple1<-function(tabla, nombrecolIDs)#Para evitar duplicados al cruzar con ListaPrecios o Inventario
{ #De una tabla extrae los registros que comparten un mismo ID,                                            
  #selecciona al PRIMER registro y borra los demás,
  #lo anterior lo hace para cada ID dentro de un campo de interés (nombrecolIDs)
  
  tabladefrecuencias<-table(tabla[,nombrecolIDs]);     
  ConjIDsRepetidos<-names(tabladefrecuencias[which(tabladefrecuencias!=1)]);
  TotalIDsRepetidos= length(ConjIDsRepetidos);
  cat("Total ID's Repetidos: ", TotalIDsRepetidos,"\n");
  
  subtabla2<-data.frame();
  if(TotalIDsRepetidos>=1)
  {
    for(n in c(1:TotalIDsRepetidos))                       
    { 
      subtabla1<-tabla[which(tabla[,nombrecolIDs]==ConjIDsRepetidos[n]),];
      subtabla2<-rbind(subtabla2, subtabla1[1,]);
      tabla<-tabla[which( tabla[,nombrecolIDs]!=ConjIDsRepetidos[n] | is.na(tabla[,nombrecolIDs]) ), ];
      #cat("ID\t",n,"de\t",TotalIDsRepetidos," repetidos\n");
    }
    
    tabla<-rbind(tabla,subtabla2);
  }
  cat("Total Registros Únicos: ", nrow(tabla),"\n");
  
  return(tabla);
}
pru<-ListaPrecios[which(ListaPrecios$NuevoKit=="SAL34643S1"),]
pru<-rbind(pru,ListaPrecios[c(1:10),])
View(pru)
pru<-QuitarMultiplicados_Simple1(pru,"NuevoKit")

#2 ACTIVADA
Orden1<-function(Indice,Inicio,Final,Agregado)#Reordenamos las variables
{
  return(c(Inicio[Indice]:Final[Indice],Agregado[Indice]))
}



names(ListaCategoriasTotales)
rm(x)
x<-ListaCategoriasTotales[1]
x[2]<-ListaCategoriasTotales[2]

x<-list(names(ListaCategoriasTotales)[1]= ListaCategoriasTotales[1])


names(x)[2]<-"COMEDORES"
names(x)
names(x[[2]])


#2.1 Cargamos las Pestañas del Archivo
ListaPrecios<-as.data.frame(read_excel("R_ListaPreciosInventarioEne-Oct2021.xlsx", sheet= "ListaPrecios"))
Inventario<-as.data.frame(read_excel("R_ListaPreciosInventarioEne-Oct2021.xlsx", sheet= "Inventario"))


#2.2 Limpiamos las tablas: Quitamos espacios (y espacios especiales) al principio y al fin de los datos
ListaPrecios<-ListaPrecios[which(!is.na(ListaPrecios$NuevoKit)), ]
ListaPrecios$NuevoKit<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$NuevoKit)
ListaPrecios$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", ListaPrecios$Descripcion)

Inventario<-Inventario[which(!is.na(Inventario$Articulo)), ]
Inventario$Articulo<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Articulo)
Inventario$Familia<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Familia)
Inventario$Categoria<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Categoria)
Inventario$Descripcion<-gsub("^([[:space:]\037]+)|([[:space:]\037]+)$", "", Inventario$Descripcion)


#2.3 Limpiamos las tablas: Quitamos duplicados 
ListaPrecios<-QuitarMultiplicados_Simple1(ListaPrecios,"NuevoKit")
Inventario<-QuitarMultiplicados_Simple1(Inventario,"Articulo")


#2.4 A la tabla ListaPrecios le creamos una nueva columna
ListaPrecios$PrecioLista_Prom<-apply(ListaPrecios[,c(6:ncol(ListaPrecios))], MARGIN = 1, FUN = mean, na.rm= TRUE)

View(ListaPrecios)
View(Inventario)


#3 (Viejito)
Agregar_PrecioLista_Dispoibilidad<-function(xx, TablaPrecioLista, TablaDisponibilidad= NULL, FechaCorteDisp= NULL)
{
  #print(names(xx))
  
  #AGREGAMOS LOS PRECIOS DE LISTA Y ORDENAMOS
  
  MES= xx[["MesActual"]]
  nvar=ncol(xx[["ResultadosMes"]])-4 #Calculamos Cantidad de Variables
  TamOrigSerie<-ncol(xx[["SerieResultadosMensuales"]])
  m=(TamOrigSerie-4)/nvar - 1 #Calculamos la cantidad de meses
  Letras<-seq((4+1),TamOrigSerie,nvar) #Calculamos los Indices de la primera variable
  ParticipUtilidades<-seq((10+1),TamOrigSerie,nvar) #Calculamos los índices de la última variable
 
  
  xx[["ResultadosMes"]]<-left_join(xx[["ResultadosMes"]], TablaPrecioLista[,c("NuevoKit",paste0("PrecioLista_",MES))], by= c("Articulo"="NuevoKit"))
  xx[["ResultadosAcum"]]<-left_join(xx[["ResultadosAcum"]], TablaPrecioLista[,c("NuevoKit","PrecioLista_Prom")], by= c("Articulo"="NuevoKit"))
  xx[["SerieResultadosMensuales"]]<-left_join(xx[["SerieResultadosMensuales"]], TablaPrecioLista[,c(-2,-3,-4,-5)], by= c("Articulo"="NuevoKit"))

  
  Precios<-c((TamOrigSerie+1):ncol(xx[["SerieResultadosMensuales"]])) #índices de la variable PrecioLista
  NuevoOrden<-as_vector(lapply(c(1:(m+1)), Orden1, Letras, ParticipUtilidades, Precios))
  xx[["SerieResultadosMensuales"]]<-xx[["SerieResultadosMensuales"]][,c(1:4,NuevoOrden)]

  
  
  #AGREGAMOS LA DISPONIBILIDAD DEL INVENTARIO
  if( !is.null(TablaDisponibilidad) & !is.null(FechaCorteDisp) )
  {
  xx[["ResultadosMes"]]<-left_join(xx[["ResultadosMes"]], TablaDisponibilidad[,c("Articulo","DispTotal")], by= c("Articulo"="Articulo"))
  xx[["ResultadosAcum"]]<-left_join(xx[["ResultadosAcum"]], TablaDisponibilidad[,c("Articulo","DispTotal")], by= c("Articulo"="Articulo"))
  xx[["SerieResultadosMensuales"]]<-left_join(xx[["SerieResultadosMensuales"]], TablaDisponibilidad[,c("Articulo","DispTotal")], by= c("Articulo"="Articulo"))
  
  names(xx[["ResultadosMes"]])[ncol(xx[["ResultadosMes"]])]<- paste0("DispTot_", FechaCorteDisp)
  names(xx[["ResultadosAcum"]])[ncol(xx[["ResultadosAcum"]])]<- paste0("DispTot_", FechaCorteDisp)
  names(xx[["SerieResultadosMensuales"]])[ncol(xx[["SerieResultadosMensuales"]])]<- paste0("DispTot_", FechaCorteDisp)
  }
  return(xx)
}

xx<-ListaCategoriasTotales[[1]]
View(xx[["SerieResultadosMensuales"]])
xx<-Agregar_PrecioLista_Dispoibilidad(xx, ListaPrecios, Inventario, "05/11/2021")
View(xx[["ResultadosMes"]])


#3 ACTIVADA (Ya incluye el MU)
Agregar_PrecioLista_Dispoibilidad<-function(xx, TablaPrecioLista, TablaDisponibilidad= NULL, FechaCorteDisp= NULL)
{
  #print(names(xx))
  
  #AGREGAMOS LOS PRECIOS DE LISTA Y ORDENAMOS
  
  MES= xx[["MesActual"]]
  nvar=ncol(xx[["ResultadosMes"]])-4 #Calculamos Cantidad de Variables
  TamOrigSerie<-ncol(xx[["SerieResultadosMensuales"]])
  m=(TamOrigSerie-4)/nvar - 1 #Calculamos la cantidad de meses
  Letras<-seq((4+1),TamOrigSerie,nvar) #Calculamos los Indices de la primera variable
  ParticipUtilidades<-seq((4+nvar),TamOrigSerie,nvar) #Calculamos los índices de la última variable
  
  xx[["ResultadosMes"]]<-left_join(xx[["ResultadosMes"]], TablaPrecioLista[,c("NuevoKit",paste0("PrecioLista_",MES))], by= c("Articulo"="NuevoKit"))
  xx[["ResultadosAcum"]]<-left_join(xx[["ResultadosAcum"]], TablaPrecioLista[,c("NuevoKit","PrecioLista_Prom")], by= c("Articulo"="NuevoKit"))
  IgnorarCols<-which( names(TablaPrecioLista)=="Descripcion" | names(TablaPrecioLista)=="Proveedor" | names(TablaPrecioLista)=="Origen" | names(TablaPrecioLista)=="CicloVida" )
  xx[["SerieResultadosMensuales"]]<-left_join(xx[["SerieResultadosMensuales"]], TablaPrecioLista[,-IgnorarCols], by= c("Articulo"="NuevoKit"))
  
  
  Precios<-c((TamOrigSerie+1):ncol(xx[["SerieResultadosMensuales"]])) #índices de la variable PrecioLista
  NuevoOrden<-as_vector(lapply(c(1:(m+1)), Orden1, Letras, ParticipUtilidades, Precios))
  xx[["SerieResultadosMensuales"]]<-xx[["SerieResultadosMensuales"]][,c(1:4,NuevoOrden)]
  
  
  
  #AGREGAMOS LA DISPONIBILIDAD DEL INVENTARIO
  #Hay que generalizar para cuando el INVENTARIO sea una serie temporal
  if( !is.null(TablaDisponibilidad) & !is.null(FechaCorteDisp) )
  {
    xx[["ResultadosMes"]]<-left_join(xx[["ResultadosMes"]], TablaDisponibilidad[,c("Articulo","DispTotal")], by= c("Articulo"="Articulo"))
    xx[["ResultadosAcum"]]<-left_join(xx[["ResultadosAcum"]], TablaDisponibilidad[,c("Articulo","DispTotal")], by= c("Articulo"="Articulo"))
    xx[["SerieResultadosMensuales"]]<-left_join(xx[["SerieResultadosMensuales"]], TablaDisponibilidad[,c("Articulo","DispTotal")], by= c("Articulo"="Articulo"))
    
    names(xx[["ResultadosMes"]])[ncol(xx[["ResultadosMes"]])]<- paste0("DispTot_", FechaCorteDisp)
    names(xx[["ResultadosAcum"]])[ncol(xx[["ResultadosAcum"]])]<- paste0("DispTot_", FechaCorteDisp)
    names(xx[["SerieResultadosMensuales"]])[ncol(xx[["SerieResultadosMensuales"]])]<- paste0("DispTot_", FechaCorteDisp)
  }
  return(xx)
}
y<-Agregar_PrecioLista_Dispoibilidad_PRUEBA(y, ListaPrecios, Inventario, "05/11/2021")
View(y[["ResultadosMes"]])
View(y[["SerieResultadosMensuales"]])
rm(Agregar_PrecioLista_Dispoibilidad_PRUEBA)



ListaCategoriasTotalesCopia<-ListaCategoriasTotales

nrow(ListaCategoriasTotalesCopia[[1]][["TablaPzasMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["TablaVentaMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["TablaUtilidadMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]])
View(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]])

ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]] %>% filter(Articulo=="SAL34643S1")
ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]]<-left_join(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]], ListaPrecios[,c("NuevoKit", "PrecioLista_10/2021")], by= c("Articulo"="NuevoKit"))
left_join()

View(ListaPrecios %>% filter(NuevoKit== "SAL34643S1"))
ListaPrecios[which(ListaPrecios$NuevoKit=="SAL34643S1"),]
#HAY DUPLICADOS EN LA LISTA DE PRECIOS!!!!!!!!!!!!!!!
#merge(x = df_1, y = df_2, by= c(), all.x = TRUE)


ListaCategoriasTotalesCopia<-lapply(ListaCategoriasTotalesCopia, Agregar_PrecioLista_Dispoibilidad, ListaPrecios, Inventario, "05/11/2021")
View(ListaCategoriasTotalesCopia[[3]][["SerieResultadosMensuales"]])

nrow(ListaCategoriasTotalesCopia[[1]][["TablaPzasMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["TablaVentaMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["TablaUtilidadMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]])
View(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]])


a









x<-data.frame(familia=c("aaa","bbb","ccc"), part=c(5,4,3))
x
y<-data.frame(familia=c("bbb","ccc","aaa"), part=c(8,6,10))
y

left_join()
sufijos<-c("01/2021","02/2021")
z<-left_join(x,y,by="familia", suffix=sufijos)
z








#PARA EL ANÁLISIS GRAL (ACTIVADA, pendiente modificar parte de serie resultados mensuales)
####################
ExtraerResultados<-function(ListaTablasCategoria, TipoResultado, disp=NULL) 
{ #Tiporesultado puede ser ResultadosMes, ResultadosAcum, SerieResultadosMensuales
  #pl: precio lista puede ser 1 si hay, y si no hay tiene que ser 0
  #disp: disponibilidad puede ser 1 si hay, y si no hay tiene que ser 0
  
  if(TipoResultado=="ResultadosMes" | TipoResultado=="ResultadosAcum")
  {
    ListaResultados<-list()
    for(n in c(1:length(ListaTablasCategoria)))
    {
      ListaResultados[[n]]<-ListaTablasCategoria[[n]][[TipoResultado]]
    }

    ResultadosCategTot<-do.call(rbind, ListaResultados)
    ResultadosCategTot<-ResultadosCategTot[which(!is.na(ResultadosCategTot$Categoria)),]
    ResultadosCategTot<-ResultadosCategTot[which(!is.na(ResultadosCategTot$Familia)),]
    
    ResultadosCategTot<-ResultadosCategTot[,c("Categoria","Familia","Articulo","Descripcion","Pzas","Venta","Utilidad","MU")]
    return(ResultadosCategTot)
  }
  
  if(TipoResultado=="SerieResultadosMensuales" & !is.null(disp))
  {
    ListaSeries<-list()
    for(n in c(1:length(ListaTablasCategoria)))
    {
      ListaSeries[[n]]<-ListaTablasCategoria[[n]][[TipoResultado]]  
    }
    SeriesCategTot<-do.call(rbind, ListaSeries)
    SeriesCategTot<-SeriesCategTot[which(!is.na(SeriesCategTot$Categoria)),]
    SeriesCategTot<-SeriesCategTot[which(!is.na(SeriesCategTot$Familia)),]
    

    nvar=ncol(ListaTablasCategoria[[1]][["ResultadosMes"]])-(4+disp) #Calculamos la cantidad de variables:-4 si porque no hay disp
    TamSerie<-ncol(SeriesCategTot)-disp #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/(nvar) - 0 #Calculamos la cantidad de meses: -1 porque quitamos el Acumulado, -0 para incluirlo
    Letras<-seq((4+1),TamSerie,(nvar)) #Índices en los que está la variable ClasifABC
    #Piezas<-seq((4+2),TamSerie,(nvar)) #Índices en los que está la variable Pzas
    #Ventas<-seq((4+3),TamSerie,(nvar)) #Índices en los que está la variable Venta
    #Utilidades<-seq((4+4),TamSerie,(nvar)) #Índices en los que está la variable Utilidad
    #MUs<-seq((4+5),TamSerie,(nvar)) #Índices en los que está la variable Utilidad
    #ParticipPiezas<-seq((4+6),TamSerie,(nvar)) #Índices en los que está la variable PartPzas
    #ParticipVentas<-seq((4+7),TamSerie,(nvar)) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+8),TamSerie,(nvar)) #Índices en los que está la variable PartUtilidad
    #Precios<-seq((4+9),TamSerie,(nvar+pl)) #Índices en los que está la variable PrecioLista
    
    ListaResultadosMeses<-list()
    TablaSoporte<-SeriesCategTot[,c("Categoria","Familia","Articulo","Descripcion")]
    for(n in c(1:m))
    {
      ListaResultadosMeses[[n]]<-cbind(TablaSoporte,SeriesCategTot[,c(Letras[n]:ParticipUtilidades[n])])
    }
    
    
    return(ListaResultadosMeses)
  }
  
  
}

ClasificarFamiliasGral<-function(x, VARIABLE, TotalGral)#TotalGral mod
{
  PartVARIABLE= paste0("Part",VARIABLE)
  
  ParticipAcumulada<-x[,PartVARIABLE]
  ParticipAcumulada<- ParticipAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ]
  if(length(ParticipAcumulada)>=2 & TotalGral!=0)#TotalGral mod
  {
    for(n in c(2:length(ParticipAcumulada)))
    {
      ParticipAcumulada[n]<-ParticipAcumulada[n-1] + ParticipAcumulada[n]
    }
    
    Clasif<-"A"
    for(n in c(2:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[n]<=50 | ParticipAcumulada[n-1]+(ParticipAcumulada[n]-ParticipAcumulada[n-1])*0.6 <= 50 )
      {
        Clasif<-c(Clasif,"A")
      }else{
        break
      }
    }
    for(m in c(n:length(ParticipAcumulada)))
    {
      if(ParticipAcumulada[m]<=80 | ParticipAcumulada[m-1]+(ParticipAcumulada[m]-ParticipAcumulada[m-1])*0.6 <= 80 )
      {
        Clasif<-c(Clasif,"B")
      }else{
        break
      }
    }
    Clasif[c(m:length(ParticipAcumulada))]<-"C"
    
  }else{
    Clasif<-ifelse(TotalGral!=0, "A", "_")#mod
    
    #Clasif<-"A"
  }
  
  
  if(VARIABLE=="Pzas")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=NA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"PartPzas"=NA,"PartAcumulada"=NA,"Venta"=NA,"Utilidad"=NA,"MU"=NA,"ClasifPorPzas"=NA)#NO HAY Costo
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorPzas <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","PartPzas","PartAcumulada","Venta","Costo","Utilidad","ClasifPorPzas")]
    }else{
      
      x$PartAcumulada <- NA
      x$ClasifPorPzas <- "_"
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
      x$ClasifPorPzas[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      x<-x[,c("Categoria","Familia","Articulo","Descripcion","Pzas","PartPzas","PartAcumulada","Venta","Utilidad","MU","ClasifPorPzas")]
      
    }
    
    
  }
  if(VARIABLE=="Venta")
  {
    if(nrow(x)==0)
    {
      X<-data.frame("Categoria"=NA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Utilidad"=NA,"MU"=NA,"ClasifPorVenta"=NA)#NO HAY Costo
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorVenta <- "_"
      # x<-x[,c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","PartVenta","PartAcumulada","Costo","Utilidad","ClasifPorPzas")]
    }else{
      
      x$PartAcumulada <- NA
      x$ClasifPorVenta <- "_"
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
      x$ClasifPorVenta[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      x<-x[,c("Categoria","Familia","Articulo","Descripcion","Pzas","Venta","PartVenta","PartAcumulada","Utilidad","MU","ClasifPorVenta")]
      
    }
    
  }
  if(VARIABLE=="Utilidad")
  {
    if(nrow(x)==0)
    {
      
      X<-data.frame("Categoria"=NA,"Familia"=NA, "Articulo"=NA, "Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"Utilidad"=NA,"PartUtilidad"=NA,"PartAcumulada"=NA,"MU"=NA,"ClasifPorUtilidad"=NA)#NO HAY Costo
      # x<-rbind(x, c(CATEGORIA,NA,NA,NA,NA,NA,NA,NA,NA))
      # names(x)<-c("Categoria","Familia", "Articulo", "Descripcion","Pzas","Venta","Costo","Utilidad",PartVARIABLE)
      # x$PartAcumulada <- NA
      # x$ClasifPorUtilidad <- "_"
    }else{
      
      x$PartAcumulada <- NA
      x$ClasifPorUtilidad <- "_"
      
      x$PartAcumulada[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- ParticipAcumulada
      x$ClasifPorUtilidad[ which(x[,VARIABLE]>0 & x[,PartVARIABLE]>0 ) ] <- Clasif
      x<-x[,c("Categoria","Familia","Articulo","Descripcion","Pzas","Venta","Utilidad","PartUtilidad","PartAcumulada","MU","ClasifPorUtilidad")]
      
    }
    
    
  }
  
  return(x)
}

TablaPartVariableGral<-function(TABLA, VARIABLE)#FORMATO DE MONTHYEAR MM/AAAA
{
  
  if(VARIABLE=="Pzas")
  {
    TABLA<-TABLA %>% arrange(desc(Pzas), desc(Utilidad))
    Total<-as.numeric(TABLA %>% summarise(sum(Pzas, na.rm = TRUE)))
    TABLA$PartPzas<-100*TABLA$Pzas/Total
  }
  if(VARIABLE=="Venta")
  {
    TABLA<-TABLA %>% arrange(desc(Venta), desc(Utilidad))
    Total<-as.numeric(TABLA %>% summarise(sum(Venta, na.rm = TRUE)))
    TABLA$PartVenta<-100*TABLA$Venta/Total
  }
  if(VARIABLE=="Utilidad")
  {
    TABLA<-TABLA %>% arrange(desc(Utilidad), desc(Venta))
    Total<-as.numeric(TABLA %>% summarise(sum(Utilidad, na.rm = TRUE)))
    TABLA$PartUtilidad<-100*TABLA$Utilidad/Total
  }
  
  PartVARIABLE= paste0("Part",VARIABLE)
  #names(x[,7])= PartVARIABLE
  TABLA<-ClasificarFamiliasGral(TABLA,VARIABLE, Total)#Total mod
  
  
  return(TABLA)
}

ResumenClasifFamiliasMesGral<-function(TABLA)
{
  NombresOriginales<-names(TABLA)
  cat("Trabajando en: ", str_sub(NombresOriginales[5],-7,-1), "\n")
  
  #Análisis MES seleccionado
  
  NombresOriginales<-names(TABLA)
  print(NombresOriginales)
  
  names(TABLA)<-c("Categoria","Familia","Articulo","Descripcion","ABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")
  
  Tabla1Mes<-TablaPartVariableGral(TABLA,"Pzas")
  Tabla2Mes<-TablaPartVariableGral(TABLA,"Venta")
  Tabla3Mes<-TablaPartVariableGral(TABLA,"Utilidad")
  
  Tabla4Mes<-Tabla3Mes[,c("Categoria","Familia","Articulo","Descripcion")]
  # print("Tabla4Mes= ")    
  # print(Tabla4Mes)
  # 
  
  if(!is.null(Tabla4Mes))
  {
    Tabla4Mes<-left_join(Tabla4Mes, Tabla1Mes[,c("Categoria","Familia","Articulo","Descripcion","Pzas","PartPzas","ClasifPorPzas")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla2Mes[,c("Categoria","Familia","Articulo","Descripcion","Venta","PartVenta","ClasifPorVenta")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla3Mes[,c("Categoria","Familia","Articulo","Descripcion","Utilidad","PartUtilidad","ClasifPorUtilidad","MU")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    #Tabla4Mes$MU <-100*(Tabla4Mes$Utilidad/Tabla4Mes$Venta) 
    Tabla4Mes$ClasifABC<-paste0(Tabla4Mes[,"ClasifPorPzas"],Tabla4Mes[,"ClasifPorVenta"],Tabla4Mes[,"ClasifPorUtilidad"])
    #Tabla4Mes<-Tabla4Mes[,c(-3,-4,-5)]
    Tabla4Mes<-Tabla4Mes[,c("Categoria","Familia","Articulo","Descripcion","ClasifABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")]
  }else{
    Tabla4Mes<-data.frame("Categoria"=NA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"ClasifABC"=NA,"Pzas"=NA,"Venta"=NA,"Utilidad"=NA,"MU"=NA,"PartPzas"=NA,"PartVenta"=NA,"PartUtilidad"=NA)
  }
  
  names(Tabla4Mes)<-NombresOriginales
  return(Tabla4Mes)
}

TablasGral<-function(ListaTablasCategoria, YEARMONTH, Disponibilidad)#YEARMONTH EN FORMATO MM/AAAA
{ #pl: precio lista puede ser 1 si hay, y si no hay tiene que ser 0
  #disp: disponibilidad puede ser 1 si hay, y si no hay tiene que ser 0
  
  cat("Trabajando en: ", YEARMONTH, "\n")
  
  #Análisis MES seleccionado
  ResultadosMesCategTot<-ExtraerResultados(ListaTablasCategoria,"ResultadosMes")
  Tabla1Mes<-TablaPartVariableGral(ResultadosMesCategTot,"Pzas")
  Tabla2Mes<-TablaPartVariableGral(ResultadosMesCategTot,"Venta")
  Tabla3Mes<-TablaPartVariableGral(ResultadosMesCategTot,"Utilidad")
  
  Tabla4Mes<-Tabla3Mes[,c("Categoria","Familia","Articulo","Descripcion")]
  # print("Tabla4Mes= ")    
  # print(Tabla4Mes)
  # 
  #View(ResultadosMesCategTot)
  #View(Tabla1Mes)
  #View(Tabla2Mes)
  #View(Tabla3Mes)
   
  # 
  if(!is.null(Tabla4Mes))
  {
    Tabla4Mes<-left_join(Tabla4Mes, Tabla1Mes[,c("Categoria","Familia","Articulo","Descripcion","Pzas","PartPzas","ClasifPorPzas")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla2Mes[,c("Categoria","Familia","Articulo","Descripcion","Venta","PartVenta","ClasifPorVenta")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    Tabla4Mes<-left_join(Tabla4Mes, Tabla3Mes[,c("Categoria","Familia","Articulo","Descripcion","Utilidad","PartUtilidad","ClasifPorUtilidad","MU")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
    #Tabla4Mes$MU <-100*(Tabla4Mes$Utilidad/Tabla4Mes$Venta) 
    Tabla4Mes$ClasifABC<-paste0(Tabla4Mes[,"ClasifPorPzas"],Tabla4Mes[,"ClasifPorVenta"],Tabla4Mes[,"ClasifPorUtilidad"])
    #Tabla4Mes<-Tabla4Mes[,c(-3,-4,-5)]
    Tabla4Mes<-Tabla4Mes[,c("Categoria","Familia","Articulo","Descripcion","ClasifABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")]
  }else{
    Tabla1Mes<-data.frame("Categoria"=NA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"Pzas"=NA,"PartPzas"=NA,"PartAcumulada"=NA,"Venta"=NA,"Utilidad"=NA,"MU"=NA,"ClasifABC"=NA)
    Tabla2Mes<-data.frame("Categoria"=NA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"PartVenta"=NA,"PartAcumulada"=NA,"Utilidad"=NA,"MU"=NA,"ClasifABC"=NA)
    Tabla3Mes<-data.frame("Categoria"=NA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"Pzas"=NA,"Venta"=NA,"Utilidad"=NA,"PartUtilidad"=NA,"PartAcumulada"=NA,"MU"=NA,"ClasifABC"=NA)
    Tabla4Mes<-data.frame("Categoria"=NA,"Familia"=NA,"Articulo"=NA,"Descripcion"=NA,"ClasifABC"=NA,"Pzas"=NA,"Venta"=NA,"Utilidad"=NA,"MU"=NA,"PartPzas"=NA,"PartVenta"=NA,"PartUtilidad"=NA)
  }
  
  
  
  
  
  #Análisis ACUMULADO
  cat("Trabajando en: ACUMULADO\n")
  ResultadosAcumCategTot<-ExtraerResultados(ListaTablasCategoria,"ResultadosAcum")
  Tabla1Acum<-TablaPartVariableGral(ResultadosAcumCategTot,"Pzas")
  Tabla2Acum<-TablaPartVariableGral(ResultadosAcumCategTot,"Venta")
  Tabla3Acum<-TablaPartVariableGral(ResultadosAcumCategTot,"Utilidad")
  
  Tabla4Acum<-Tabla3Acum[,c("Categoria","Familia","Articulo","Descripcion")]
  # print("Tabla4Acuulado")
  # print(Tabla4Acum)
  
  Tabla4Acum<-left_join(Tabla4Acum, Tabla1Acum[,c("Categoria","Familia","Articulo","Descripcion","Pzas","PartPzas","ClasifPorPzas")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  Tabla4Acum<-left_join(Tabla4Acum, Tabla2Acum[,c("Categoria","Familia","Articulo","Descripcion","Venta","PartVenta","ClasifPorVenta")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  Tabla4Acum<-left_join(Tabla4Acum, Tabla3Acum[,c("Categoria","Familia","Articulo","Descripcion","Utilidad","PartUtilidad","ClasifPorUtilidad","MU")], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  #Tabla4Acum$MU <-100*(Tabla4Mes$Utilidad/Tabla4Mes$Venta) 
  Tabla4Acum$ClasifABC<-paste0(Tabla4Acum[,"ClasifPorPzas"],Tabla4Acum[,"ClasifPorVenta"],Tabla4Acum[,"ClasifPorUtilidad"])
  #Tabla4Acum<-Tabla4Mes[,c(-3,-4,-5)]
  Tabla4Acum<-Tabla4Acum[,c("Categoria","Familia","Articulo","Descripcion","ClasifABC","Pzas","Venta","Utilidad","MU","PartPzas","PartVenta","PartUtilidad")]
  
  
  
  #Analisis SeriesMensuales
  cat("Trabajando en: SERIE MENSUAL\n")
  
  ListaResultadosMensuales<-ExtraerResultados(ListaTablasCategoria,"SerieResultadosMensuales", disp = Disponibilidad)
  ListaResultadosMensuales<-lapply(ListaResultadosMensuales,ResumenClasifFamiliasMesGral)
  
  # 
  # #Resumen Clasif de TODOS LOS MESES  
  # MESES<-unique(str_sub(TABLA$Fecha,4,10))#MesAño
  # #ANIOS<-unique(str_sub(TABLA$Fecha,7,10))#Año
  # ListaMeses<-lapply(MESES, ResumenClasifFamiliasMes, CATEGORIA, TABLA, REGION)
  
  #Nos quedamos con el soporte y el orden de la tabla del acumulado (la última)
  Tabla5<-ListaResultadosMensuales[[length(ListaResultadosMensuales)]][,c("Categoria","Familia","Articulo","Descripcion")]
  for(n in c(1:length(ListaResultadosMensuales)))
  {
    Tabla5<-left_join(Tabla5, ListaResultadosMensuales[[n]], by=c("Categoria"="Categoria","Familia"="Familia","Articulo"="Articulo","Descripcion"="Descripcion"))
  }
  
  Tabla1Mes<-Tabla1Mes[,-c(3,4)]
  Tabla2Mes<-Tabla2Mes[,-c(3,4)]
  Tabla3Mes<-Tabla3Mes[,-c(3,4)]
  Tabla4Mes<-Tabla4Mes[,-c(3,4)]
  Tabla1Acum<-Tabla1Acum[,-c(3,4)]
  Tabla2Acum<-Tabla2Acum[,-c(3,4)]
  Tabla3Acum<-Tabla3Acum[,-c(3,4)]
  Tabla4Acum<-Tabla4Acum[,-c(3,4)]
  Tabla5<-Tabla5[,-c(3,4)]
  
  
  return(list(MesActual=YEARMONTH, TablaPzasMes=Tabla1Mes, TablaVentaMes=Tabla2Mes, TablaUtilidadMes=Tabla3Mes, ResultadosMes=Tabla4Mes, TablaPzasAcum=Tabla1Acum, TablaVentaAcum=Tabla2Acum, TablaUtilidadAcum=Tabla3Acum, ResultadosAcum=Tabla4Acum, SerieResultadosMensuales= Tabla5))
}
####################



#PARA EXPORTAR A EXCEL CON RESUMEN SIMPLE
#https://api.rpubs.com/MOrtiz/319605 Manual openxlsx
#ESTILOS PARA EXCEL
###################
EstiloMesTop<-createStyle(fontSize = 20,fgFill = "#a5c6fa",border = "TopBottom")
EstiloAcumuladoTop<-createStyle(fontSize = 20,fgFill = "#3cc962",border = "TopBottom")
EstiloResultadosMensualesTop<-createStyle(fontSize = 20,fgFill = "#e03442",border = "TopBottom")

EstiloPorPzas<-createStyle(fgFill = "#f5f5f5",border = "TopBottom", halign = "left")
EstiloPorVenta<-createStyle(fgFill = "#ebebeb",border = "TopBottom", halign = "left")
EstiloPorUtilidad<-createStyle(fgFill = "#dedede",border = "TopBottom", halign = "left")
EstiloResultado<-createStyle(fontColour = "#ffffff" ,fgFill = "#737373",border = "TopBottom", halign = "left")
EstiloResultadosMensuales<-createStyle(fontColour = "#ffffff" ,fgFill = "#737373",border = "TopBottom", halign = "left")

EstiloMesBody<-createStyle(fgFill = "#edf3fc",halign = "left", numFmt = "#,##0;-#,##0")
EstiloAcumuladoBody<-createStyle(fgFill = "#d0f7db",halign = "left", numFmt = "#,##0;-#,##0")
EstiloResultadosMensualesBody<-createStyle(fgFill = "#ffe8ea",halign = "left", numFmt = "#,##0.0;-#,##0.0")
###################



#FUNCIONES DE EXPORTACIÓN VIEJAS
###############################

#VIEJITO (DESACTIVADO)
ExportarTablasExcelSimple<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    print(CATEGORIAS[n])
    # 
    # #"MES SELECCIONADO"
    # writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = 1)
    # writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
    #                                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
    #                                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
    #                                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = 1)
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(1:42))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(1:12))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(13:24))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(25:36))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(37:42))
    # addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:300, cols = 1:42, gridExpand = TRUE)
    # 
    # #"ACUMULADO"
    # DesfaseGral= ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]]) + ncol(ListaTablasCategoria[[n]][["TablaVentaMes"]]) + ncol(ListaTablasCategoria[[n]][["TablaUtilidadMes"]]) + ncol(ListaTablasCategoria[[n]][["ResultadosMes"]]) + 4
    # writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = 43)
    # writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][[6]],
    #                                    PorVenta=rep(""), ListaTablasCategoria[[n]][[7]],
    #                                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][[8]],
    #                                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][[9]]), startRow = 2, startCol = 43)
    # addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(43:84))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(43:54))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(55:66))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(67:78))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(79:84))
    # addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:300, cols = 43:84, gridExpand = TRUE)
    # 
    # 
    # #"TODOS LOS MESES"
    # writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = 85)
    # 
    # 
    # #COMPLETA
    # writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]]), startRow = 2, startCol = 85)
    # #SOLO PartVenta
    # #writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]][ , -seq(4, ncol(ListaTablasCategoria[[n]][[10]]), 2) ]), startRow = 2, startCol = 69)
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(85:141))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(85:141))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:300, cols = 85:141, gridExpand = TRUE)
    # 
    # 
    # setColWidths(wb,CATEGORIAS[n], cols = c(1,12,13,24,25,36,37,42,43,54,55,66,67,78,79,84,85), widths = 15)
    # 
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-4 #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    Precios<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    #  writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #            startRow = 1, startCol = SeccionGralFin[3])
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    # solo si hay Disponibilidad
    # writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #           startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    # 
    
    # solo si hay Disponibilidad:
    # writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #si no la hay:
    writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+4*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+4*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5+4*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+4*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5+4*(m+1))))
    
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}
ExportarTablasExcelSimple(Resumen, ListaCategoriasTotalesCopia, NombreArchiv)

nrow(ListaCategoriasTotalesCopia[[1]][["TablaPzasMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["TablaVentaMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["TablaUtilidadMes"]])
nrow(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]])
View(ListaCategoriasTotalesCopia[[1]][["ResultadosMes"]])

names(ListaCategoriasTotalesCopia[[1]])


?getCellRefs()
?getNamedRegions()

#Sin Precios de Lista ni Disponibilidad (Viejito)
ExportarTablasExcelSimple_0<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    print(CATEGORIAS[n])
    # 
    # #"MES SELECCIONADO"
    # writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = 1)
    # writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
    #                                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
    #                                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
    #                                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = 1)
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(1:42))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(1:12))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(13:24))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(25:36))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(37:42))
    # addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:300, cols = 1:42, gridExpand = TRUE)
    # 
    # #"ACUMULADO"
    # DesfaseGral= ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]]) + ncol(ListaTablasCategoria[[n]][["TablaVentaMes"]]) + ncol(ListaTablasCategoria[[n]][["TablaUtilidadMes"]]) + ncol(ListaTablasCategoria[[n]][["ResultadosMes"]]) + 4
    # writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = 43)
    # writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][[6]],
    #                                    PorVenta=rep(""), ListaTablasCategoria[[n]][[7]],
    #                                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][[8]],
    #                                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][[9]]), startRow = 2, startCol = 43)
    # addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(43:84))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(43:54))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(55:66))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(67:78))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(79:84))
    # addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:300, cols = 43:84, gridExpand = TRUE)
    # 
    # 
    # #"TODOS LOS MESES"
    # writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = 85)
    # 
    # 
    # #COMPLETA
    # writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]]), startRow = 2, startCol = 85)
    # #SOLO PartVenta
    # #writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]][ , -seq(4, ncol(ListaTablasCategoria[[n]][[10]]), 2) ]), startRow = 2, startCol = 69)
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(85:141))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(85:141))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:300, cols = 85:141, gridExpand = TRUE)
    # 
    # 
    # setColWidths(wb,CATEGORIAS[n], cols = c(1,12,13,24,25,36,37,42,43,54,55,66,67,78,79,84,85), widths = 15)
    # 
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-4 #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    #Precios<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,ParticipPiezas,ParticipVentas,ParticipUtilidades)]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    # solo si hay Disponibilidad
    # writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #           startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    # 
    
    # solo si hay Disponibilidad:
    # writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #si no la hay:
    writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+4+3*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+4+3*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+4+3*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+4+3*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+4+3*(m+1))))
    
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}

#Sin Precios Lista ni Disponibilidad PRUEBA
ExportarTablasExcelSimple_0<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    cat("Hoja: ", CATEGORIAS[n], "\n")
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-4 #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    MUs<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades)]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    writeData(wb, CATEGORIAS[n], "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    
    
    #ESCRIBIMOS LAS TABLAS
    writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5+6*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+6*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5+6*(m+1))))
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  cat("GUARDANDO\n")
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}

#Con Precios de Lista (Viejito)
ExportarTablasExcelSimple_1<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    print(CATEGORIAS[n])
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-4 #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    Precios<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    #  writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #            startRow = 1, startCol = SeccionGralFin[3])
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    # solo si hay Disponibilidad
    # writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #           startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    # 
    
    # solo si hay Disponibilidad:
    # writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #si no la hay:
    writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+4+4*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+4+4*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+4+4*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+4+4*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+4+4*(m+1))))
    
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}

#Con Precios de Lista PRUEBA
ExportarTablasExcelSimple_1<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    cat("Hoja: ", CATEGORIAS[n], "\n")
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-4 #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    MUs<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    Precios<-seq((4+9),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+8*(m+1))
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    writeData(wb, CATEGORIAS[n], "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    writeData(wb, CATEGORIAS[n], "PRECIOS LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    # solo si hay Disponibilidad
    #writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #         startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    
    
    #ESCRIBIMOS LAS TABLAS
    # solo si hay Disponibilidad:
    #writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #si no la hay:
    writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+0)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5+7*(m+1))))
    
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  cat("GUARDANDO\n")
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}

#Con Precios de Lista y Disponibilidad PRUEBA
ExportarTablasExcelSimple_2<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    cat("Hoja: ", CATEGORIAS[n], "\n")
    # 
    # #"MES SELECCIONADO"
    # writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = 1)
    # writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
    #                                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
    #                                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
    #                                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = 1)
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(1:42))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(1:12))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(13:24))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(25:36))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(37:42))
    # addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:300, cols = 1:42, gridExpand = TRUE)
    # 
    # #"ACUMULADO"
    # DesfaseGral= ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]]) + ncol(ListaTablasCategoria[[n]][["TablaVentaMes"]]) + ncol(ListaTablasCategoria[[n]][["TablaUtilidadMes"]]) + ncol(ListaTablasCategoria[[n]][["ResultadosMes"]]) + 4
    # writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = 43)
    # writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][[6]],
    #                                    PorVenta=rep(""), ListaTablasCategoria[[n]][[7]],
    #                                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][[8]],
    #                                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][[9]]), startRow = 2, startCol = 43)
    # addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(43:84))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(43:54))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(55:66))
    # addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(67:78))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(79:84))
    # addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:300, cols = 43:84, gridExpand = TRUE)
    # 
    # 
    # #"TODOS LOS MESES"
    # writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = 85)
    # 
    # 
    # #COMPLETA
    # writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]]), startRow = 2, startCol = 85)
    # #SOLO PartVenta
    # #writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]][ , -seq(4, ncol(ListaTablasCategoria[[n]][[10]]), 2) ]), startRow = 2, startCol = 69)
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(85:141))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(85:141))
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:300, cols = 85:141, gridExpand = TRUE)
    # 
    # 
    # setColWidths(wb,CATEGORIAS[n], cols = c(1,12,13,24,25,36,37,42,43,54,55,66,67,78,79,84,85), widths = 15)
    # 
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-5 #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-1 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    MUs<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    Precios<-seq((4+9),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+8*(m+1))
    #  writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #            startRow = 1, startCol = SeccionGralFin[3])
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    writeData(wb, CATEGORIAS[n], "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    writeData(wb, CATEGORIAS[n], "PRECIOS LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    # solo si hay Disponibilidad
    writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
              startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    
    
    #ESCRIBIMOS LAS TABLAS
    # solo si hay Disponibilidad:
    writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #si no la hay:
    #writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,ParticipVentas,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+1)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+1)))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+1), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+7*(m+1)+1)), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5+7*(m+1))))
    
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  cat("GUARDANDO\n")
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}



#Sin Precios Lista ni Disponibilidad PRUEBA (DESACTIVADA)
ExportarTablasExcelSimpleGral_0<-function(Resumen,ListaTablasGral, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[["BaseOrigen"]])
  
  
  #Creamos hoja RESUMEN
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[["MesActual"]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[["ResumenMes"]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[["ResumenAcum"]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  #CREAMOS LAS HOJA AnálisisGral
  
  addWorksheet(wb, "AnálisisGral", gridLines = FALSE)
  cat("Hoja: AnálisisGral\n")
  
  #Definimos los límites de los estilos
  ncolTabVar<-ncol(ListaTablasGral[["TablaPzasMes"]])
  ncolTabRes<-ncol(ListaTablasGral[["ResultadosMes"]])
  ncolTabSerie<-ncol(ListaTablasGral[["SerieResultadosMensuales"]])
  SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
  SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
  
  
  #MES ACTUAL
  print("Mes Actual")
  writeData(wb, "AnálisisGral", ListaTablasGral[["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
  writeData(wb, "AnálisisGral", 
            cbind(PorPzas=rep(""), ListaTablasGral[["TablaPzasMes"]],
                  PorVenta=rep(""), ListaTablasGral[["TablaVentaMes"]],
                  PorUtilidad=rep(""), ListaTablasGral[["TablaUtilidadMes"]],
                  ClasifTotal=rep(""), ListaTablasGral[["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
  
  #Para Estilo Mes
  SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
  SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
  
  addStyle(wb, "AnálisisGral", style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
  addStyle(wb, "AnálisisGral", style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
  
  
  
  #ACUMULADO
  print("Acumulado")
  writeData(wb, "AnálisisGral", "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
  writeData(wb, "AnálisisGral", 
            cbind(PorPzas=rep(""), ListaTablasGral[["TablaPzasAcum"]],
                  PorVenta=rep(""), ListaTablasGral[["TablaVentaAcum"]],
                  PorUtilidad=rep(""), ListaTablasGral[["TablaUtilidadAcum"]],
                  ClasifTotal=rep(""), ListaTablasGral[["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
  
  #para Estilo Acumulado
  SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
  SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
  
  addStyle(wb, "AnálisisGral", style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
  addStyle(wb, "AnálisisGral", style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
  
  
  
  #SERIE MENSUAL
  #"TODOS LOS MESES"
  print("Todos Los Meses")
  writeData(wb, "AnálisisGral", "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
  
  #al 4 le restamos 2 debido a que no hay Articulo ni Descripcion
  nvar=ncol(ListaTablasGral[["ResultadosMes"]])-(4-2) #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
  TamSerie<-ncol(ListaTablasGral[["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
  m=(TamSerie-2)/nvar - 1 #Calculamos la cantidad de meses
  Letras<-seq((2+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
  Piezas<-seq((2+2),TamSerie,nvar) #Índices en los que está la variable Pzas
  Ventas<-seq((2+3),TamSerie,nvar) #Índices en los que está la variable Venta
  Utilidades<-seq((2+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
  MUs<-seq((2+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
  ParticipPiezas<-seq((2+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
  ParticipVentas<-seq((2+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
  ParticipUtilidades<-seq((2+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
  cat("m= ", m, "\n")
  
  # #FORMATO COMPLETo
  #  writeData(wb, "AnálisisGral", "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5-2)
  #  writeData(wb, "AnálisisGral", "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+(m+1))
  #  writeData(wb, "AnálisisGral", "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+2*(m+1))
  #  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+3*(m+1))
  #  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+4*(m+1))
  #  writeData(wb, "AnálisisGral", "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+5*(m+1))
  #  writeData(wb, "AnálisisGral", "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+6*(m+1))
  #  writeData(wb, "AnálisisGral", "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+7*(m+1))
  #  
  #  writeData(wb, "AnálisisGral", cbind(ClasifABC=rep(""), ListaTablasGral[["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades)]), startRow = 2, startCol = SeccionGralInicio[3])
  #  
  #  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
  #  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
  #  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
  # 
  #  setColWidths(wb,"AnálisisGral", cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
  #  addFilter(wb, "AnálisisGral", rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
  
  
  #FORMATO IRVING
  #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
  writeData(wb, "AnálisisGral", "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5-2)
  writeData(wb, "AnálisisGral", "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+(m+1))
  writeData(wb, "AnálisisGral", "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+2*(m+1))
  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+3*(m+1))
  writeData(wb, "AnálisisGral", "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5-2+4*(m+1))
  writeData(wb, "AnálisisGral", "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+5*(m+1))
  
  
  #ESCRIBIMOS LAS TABLAS
  writeData(wb, "AnálisisGral", cbind(ClasifTotal=rep(""), ListaTablasGral[["SerieResultadosMensuales"]][,c(1:(4-2),Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
  
  
  
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
  
  setColWidths(wb,"AnálisisGral", cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
  addFilter(wb, "AnálisisGral", rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5-2+6*(m+1))))
  
  
  
  #deleteData()
  #sheetVisibility()
  
  cat("GUARDANDO\n")
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}

#PrcioLista y Disponibilidad son BOOLEANOS (DESACTIVADA)
ExportarTablasExcelCompleto<-function(Resumen, ListaTablasCategoria, ListaTablasGral, PrecioLista, Disponibilidad, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  cat("Hoja: BASE\n")
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  cat("Hoja: RESUMEN\n")
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:50000, cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:50000, cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  
  
  
  #CREAMOS LA HOJA AnálisisGral
  addWorksheet(wb, "AnálisisGral", gridLines = FALSE)
  cat("Hoja: AnálisisGral\n")
  
  #Definimos los límites de los estilos
  ncolTabVar<-ncol(ListaTablasGral[["TablaPzasMes"]])
  ncolTabRes<-ncol(ListaTablasGral[["ResultadosMes"]])
  ncolTabSerie<-ncol(ListaTablasGral[["SerieResultadosMensuales"]])
  SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
  SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
  
  #MES ACTUAL
  print("Mes Actual")
  writeData(wb, "AnálisisGral", ListaTablasGral[["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
  writeData(wb, "AnálisisGral", 
            cbind(PorPzas=rep(""), ListaTablasGral[["TablaPzasMes"]],
                  PorVenta=rep(""), ListaTablasGral[["TablaVentaMes"]],
                  PorUtilidad=rep(""), ListaTablasGral[["TablaUtilidadMes"]],
                  ClasifTotal=rep(""), ListaTablasGral[["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
  
  #Para Estilo Mes
  SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
  SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
  
  addStyle(wb, "AnálisisGral", style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
  addStyle(wb, "AnálisisGral", style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
  
  #ACUMULADO
  print("Acumulado")
  writeData(wb, "AnálisisGral", "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
  writeData(wb, "AnálisisGral", 
            cbind(PorPzas=rep(""), ListaTablasGral[["TablaPzasAcum"]],
                  PorVenta=rep(""), ListaTablasGral[["TablaVentaAcum"]],
                  PorUtilidad=rep(""), ListaTablasGral[["TablaUtilidadAcum"]],
                  ClasifTotal=rep(""), ListaTablasGral[["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
  
  #para Estilo Acumulado
  SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
  SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
  
  addStyle(wb, "AnálisisGral", style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
  addStyle(wb, "AnálisisGral", style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
  
  #SERIE MENSUAL
  print("SeriesMensuales")
  writeData(wb, "AnálisisGral", "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
  
  #al 4 le restamos 2 debido a que no hay Articulo ni Descripcion
  nvar=ncol(ListaTablasGral[["ResultadosMes"]])-(4-2) #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
  TamSerie<-ncol(ListaTablasGral[["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
  m=(TamSerie-2)/nvar - 1 #Calculamos la cantidad de meses
  Letras<-seq((2+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
  Piezas<-seq((2+2),TamSerie,nvar) #Índices en los que está la variable Pzas
  Ventas<-seq((2+3),TamSerie,nvar) #Índices en los que está la variable Venta
  Utilidades<-seq((2+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
  MUs<-seq((2+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
  ParticipPiezas<-seq((2+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
  ParticipVentas<-seq((2+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
  ParticipUtilidades<-seq((2+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
  #cat("m= ", m, "\n")
  
  # #FORMATO COMPLETo
  #  writeData(wb, "AnálisisGral", "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5-2)
  #  writeData(wb, "AnálisisGral", "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+(m+1))
  #  writeData(wb, "AnálisisGral", "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+2*(m+1))
  #  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+3*(m+1))
  #  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+4*(m+1))
  #  writeData(wb, "AnálisisGral", "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+5*(m+1))
  #  writeData(wb, "AnálisisGral", "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+6*(m+1))
  #  writeData(wb, "AnálisisGral", "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+7*(m+1))
  #  
  #  writeData(wb, "AnálisisGral", cbind(ClasifABC=rep(""), ListaTablasGral[["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades)]), startRow = 2, startCol = SeccionGralInicio[3])
  #  
  #  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
  #  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
  #  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
  # 
  #  setColWidths(wb,"AnálisisGral", cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
  #  addFilter(wb, "AnálisisGral", rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
  
  
  #FORMATO IRVING
  #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
  writeData(wb, "AnálisisGral", "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5-2)
  writeData(wb, "AnálisisGral", "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+(m+1))
  writeData(wb, "AnálisisGral", "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+2*(m+1))
  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+3*(m+1))
  writeData(wb, "AnálisisGral", "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5-2+4*(m+1))
  writeData(wb, "AnálisisGral", "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+5*(m+1))
  
  #ESCRIBIMOS LAS TABLAS
  writeData(wb, "AnálisisGral", cbind(ClasifTotal=rep(""), ListaTablasGral[["SerieResultadosMensuales"]][,c(1:(4-2),Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
  
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
  setColWidths(wb,"AnálisisGral", cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
  addFilter(wb, "AnálisisGral", rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5-2+6*(m+1))))
  
  
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  print(CATEGORIAS)
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    cat("Hoja: ", CATEGORIAS[n], "\n")
    
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:2000, cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:2000, cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    #PrecioLista Disponibilidad
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-(4+Disponibilidad) #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-(Disponibilidad) #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    MUs<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    if(PrecioLista==1)
    {
      Precios<-seq((4+9),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    }
    
    
    # #FORMATO COMPLETo
    #  writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    #  writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    #  writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    #  writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PARTICIPACION UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    #  writeData(wb, CATEGORIAS[n], "PRECIO LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+8*(m+1))
    #  writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #            startRow = 1, startCol = SeccionGralFin[3])
    #  
    #  writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    #  
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    #  addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    # 
    #  setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    #  addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    
    #FORMATO IRVING
    #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    writeData(wb, CATEGORIAS[n], "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    if(PrecioLista==1)
    {
      writeData(wb, CATEGORIAS[n], "PRECIOS LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    }
    if(Disponibilidad==1 & PrecioLista==1)
    {
      # solo si hay Disponibilidad
      writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
                startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    }
    
    
    
    #ESCRIBIMOS LAS TABLAS
    if(Disponibilidad==1 & PrecioLista==1)
    {
      writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    }
    if(Disponibilidad==0 & PrecioLista==1)
    {
      writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
      
    }
    if(Disponibilidad==0 & PrecioLista==0)
    {
      writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
    }
    
    
    
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad))))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad))))# +1 si hay disp, +0 si no la hay
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad)), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad))), widths = 17)# +1 si hay disp, +0 si no la hay
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1))))
    
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  cat("GUARDANDO\n")
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}

###############################

#PrcioLista y Disponibilidad son BOOLEANOS (ACTIVADA NUEVA)
ExportarTablasExcelCompleto<-function(Resumen, ListaTablasCategoria, ListaTablasGral, PrecioLista, Disponibilidad, NombreArchivo)
{
  wb<-createWorkbook()
  
  #Creamos hoja de la BASE
  cat("Hoja: BASE\n")
  addWorksheet(wb, "BASE")
  writeData(wb, "BASE", Resumen[[1]])
  
  
  #Creamos hoja RESUMEN
  cat("Hoja: RESUMEN\n")
  addWorksheet(wb, "RESUMEN", gridLines = FALSE)
  
  #"MES SELECCIONADO"
  writeData(wb, "RESUMEN", Resumen[[2]], startRow = 1, startCol = 1)
  writeData(wb, "RESUMEN", Resumen[[3]], startRow = 2, startCol = 1)
  addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:6))
  addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:(nrow(Resumen[["ResumenAcum"]])+2), cols = 1:6, gridExpand = TRUE)
  
  #"ACUMULADO"
  writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 7)
  writeData(wb, "RESUMEN", Resumen[[4]],startRow = 2, startCol = 7)
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(7:12))
  addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:(nrow(Resumen[["ResumenAcum"]])+2), cols = 7:12, gridExpand = TRUE)
  
  setColWidths(wb,"RESUMEN", cols = c(1,2,7,8), widths = 15)
  addFilter(wb, "RESUMEN", rows = 2, cols = c(1:12))
  
  
  
  
  #CREAMOS LA HOJA AnálisisGral
  addWorksheet(wb, "AnálisisGral", gridLines = FALSE)
  cat("Hoja: AnálisisGral\n")
  
  #Definimos los límites de los estilos
  ncolTabVar<-ncol(ListaTablasGral[["TablaPzasMes"]])
  ncolTabRes<-ncol(ListaTablasGral[["ResultadosMes"]])
  ncolTabSerie<-ncol(ListaTablasGral[["SerieResultadosMensuales"]])
  nfilTabSerie<-nrow(ListaTablasGral[["SerieResultadosMensuales"]])
  SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
  SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
  
  #MES ACTUAL
  print("Mes Actual")
  writeData(wb, "AnálisisGral", ListaTablasGral[["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
  writeData(wb, "AnálisisGral", 
            cbind(PorPzas=rep(""), ListaTablasGral[["TablaPzasMes"]],
                  PorVenta=rep(""), ListaTablasGral[["TablaVentaMes"]],
                  PorUtilidad=rep(""), ListaTablasGral[["TablaUtilidadMes"]],
                  ClasifTotal=rep(""), ListaTablasGral[["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
  
  #Para Estilo Mes
  SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
  SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
  
  addStyle(wb, "AnálisisGral", style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
  addStyle(wb, "AnálisisGral", style = EstiloMesBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
  
  #ACUMULADO
  print("Acumulado")
  writeData(wb, "AnálisisGral", "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
  writeData(wb, "AnálisisGral", 
            cbind(PorPzas=rep(""), ListaTablasGral[["TablaPzasAcum"]],
                  PorVenta=rep(""), ListaTablasGral[["TablaVentaAcum"]],
                  PorUtilidad=rep(""), ListaTablasGral[["TablaUtilidadAcum"]],
                  ClasifTotal=rep(""), ListaTablasGral[["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
  
  #para Estilo Acumulado
  SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
  SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
  
  addStyle(wb, "AnálisisGral", style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
  addStyle(wb, "AnálisisGral", style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
  addStyle(wb, "AnálisisGral", style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
  addStyle(wb, "AnálisisGral", style = EstiloAcumuladoBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
  
  #SERIE MENSUAL
  print("SeriesMensuales")
  writeData(wb, "AnálisisGral", "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
  
  #al 4 le restamos 2 debido a que no hay Articulo ni Descripcion
  nvar=ncol(ListaTablasGral[["ResultadosMes"]])-(4-2) #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
  TamSerie<-ncol(ListaTablasGral[["SerieResultadosMensuales"]])-0 #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
  m=(TamSerie-2)/nvar - 1 #Calculamos la cantidad de meses
  Letras<-seq((2+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
  Piezas<-seq((2+2),TamSerie,nvar) #Índices en los que está la variable Pzas
  Ventas<-seq((2+3),TamSerie,nvar) #Índices en los que está la variable Venta
  Utilidades<-seq((2+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
  MUs<-seq((2+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
  ParticipPiezas<-seq((2+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
  ParticipVentas<-seq((2+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
  ParticipUtilidades<-seq((2+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
  #cat("m= ", m, "\n")
  
  #FORMATO COMPLETo
  writeData(wb, "AnálisisGral", "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5-2)
  writeData(wb, "AnálisisGral", "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+(m+1))
  writeData(wb, "AnálisisGral", "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+2*(m+1))
  writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+3*(m+1))
  writeData(wb, "AnálisisGral", "MÁRGENES DE UTILIDAD %", startRow = 1, startCol = SeccionGralInicio[3]+5-2+4*(m+1))
  writeData(wb, "AnálisisGral", "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+5*(m+1))
  writeData(wb, "AnálisisGral", "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+6*(m+1))
  writeData(wb, "AnálisisGral", "PARTICIPACIÓN UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+7*(m+1))
   
  writeData(wb, "AnálisisGral", cbind(ClasifABC=rep(""), ListaTablasGral[["SerieResultadosMensuales"]][,c(1:2,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades)]), startRow = 2, startCol = SeccionGralInicio[3])
   
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
  addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
  
  setColWidths(wb,"AnálisisGral", cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
  addFilter(wb, "AnálisisGral", rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
  
  # 
  # #FORMATO IRVING
  # #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
  # writeData(wb, "AnálisisGral", "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5-2)
  # writeData(wb, "AnálisisGral", "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+(m+1))
  # writeData(wb, "AnálisisGral", "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+2*(m+1))
  # writeData(wb, "AnálisisGral", "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5-2+3*(m+1))
  # writeData(wb, "AnálisisGral", "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5-2+4*(m+1))
  # writeData(wb, "AnálisisGral", "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5-2+5*(m+1))
  # 
  # #ESCRIBIMOS LAS TABLAS
  # writeData(wb, "AnálisisGral", cbind(ClasifTotal=rep(""), ListaTablasGral[["SerieResultadosMensuales"]][,c(1:(4-2),Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
  # 
  # addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
  # addStyle(wb, "AnálisisGral", style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)))# +1 si hay disp, +0 si no la hay
  # addStyle(wb, "AnálisisGral", style = EstiloResultadosMensualesBody, rows = 3:2000, cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
  # setColWidths(wb,"AnálisisGral", cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5-2+6*(m+1)+0)), widths = 17)# +1 si hay disp, +0 si no la hay
  # addFilter(wb, "AnálisisGral", rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5-2+6*(m+1))))
  # 
  
  
  #CREAMOS LAS HOJAS DE CADA CATEGORÍA
  CATEGORIAS<-Resumen[[5]]
  print(CATEGORIAS)
  for(n in c(1:length(CATEGORIAS)))
  {
    addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
    cat("Hoja: ", CATEGORIAS[n], "\n")
    
    
    #Definimos los límites de los estilos
    ncolTabVar<-ncol(ListaTablasCategoria[[n]][["TablaPzasMes"]])
    ncolTabRes<-ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])
    ncolTabSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    nfilTabSerie<-nrow(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])
    SeccionGralInicio<-c(1, ( ncolTabVar*3 + ncolTabRes + 4 )+1, 2*( ncolTabVar*3 + ncolTabRes + 4 )+1)
    SeccionGralFin<-c(( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 ), 2*( ncolTabVar*3 + ncolTabRes + 4 )+(ncolTabSerie+1))
    
    
    #MES ACTUAL
    print("Mes Actual")
    writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][["MesActual"]], startRow = 1, startCol = SeccionGralInicio[1])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasMes"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaMes"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadMes"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosMes"]]), startRow = 2, startCol = SeccionGralInicio[1])
    
    #Para Estilo Mes
    SeccionLocalInicio<-c(1, (ncolTabVar+1)+1, 2*(ncolTabVar+1)+1, 3*(ncolTabVar+1)+1)
    SeccionLocalFin<-c(ncolTabVar+1, 2*(ncolTabVar+1), 3*(ncolTabVar+1), 3*(ncolTabVar+1)+(ncolTabRes+1))
    
    addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(SeccionGralInicio[1]:SeccionGralFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[1]:SeccionGralFin[1], gridExpand = TRUE)
    
    
    
    #ACUMULADO
    print("Acumulado")
    writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = SeccionGralInicio[2])
    writeData(wb, CATEGORIAS[n], 
              cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][["TablaPzasAcum"]],
                    PorVenta=rep(""), ListaTablasCategoria[[n]][["TablaVentaAcum"]],
                    PorUtilidad=rep(""), ListaTablasCategoria[[n]][["TablaUtilidadAcum"]],
                    ClasifTotal=rep(""), ListaTablasCategoria[[n]][["ResultadosAcum"]]), startRow = 2, startCol = SeccionGralInicio[2])
    
    #para Estilo Acumulado
    SeccionLocalInicio<-SeccionGralFin[1]+SeccionLocalInicio
    SeccionLocalFin<-SeccionGralFin[1]+SeccionLocalFin
    
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(SeccionGralInicio[2]:SeccionGralFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(SeccionLocalInicio[1]:SeccionLocalFin[1]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(SeccionLocalInicio[2]:SeccionLocalFin[2]))
    addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(SeccionLocalInicio[3]:SeccionLocalFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(SeccionLocalInicio[4]:SeccionLocalFin[4]))
    addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[2]:SeccionGralFin[2], gridExpand = TRUE)
    
    
    #PrecioLista Disponibilidad
    #SERIE MENSUAL
    #"TODOS LOS MESES"
    print("Todos Los Meses")
    writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = SeccionGralInicio[3])
    
    nvar=ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])-(4+Disponibilidad) #Calculamos la cantidad de variables: -5 si hay Disponibilidad y -4 si no la hay
    TamSerie<-ncol(ListaTablasCategoria[[n]][["SerieResultadosMensuales"]])-(Disponibilidad) #no consideramos la disponibilidad: -1 si hay Disp y -0 si no la hay
    m=(TamSerie-4)/nvar - 1 #Calculamos la cantidad de meses
    Letras<-seq((4+1),TamSerie,nvar) #Índices en los que está la variable ClasifABC
    Piezas<-seq((4+2),TamSerie,nvar) #Índices en los que está la variable Pzas
    Ventas<-seq((4+3),TamSerie,nvar) #Índices en los que está la variable Venta
    Utilidades<-seq((4+4),TamSerie,nvar) #Índices en los que está la variable Utilidad
    MUs<-seq((4+5),TamSerie,nvar) #Índices en los que está la variable Utilidad
    ParticipPiezas<-seq((4+6),TamSerie,nvar) #Índices en los que está la variable PartPzas
    ParticipVentas<-seq((4+7),TamSerie,nvar) #Índices en los que está la variable PartVenta
    ParticipUtilidades<-seq((4+8),TamSerie,nvar) #Índices en los que está la variable PartUtilidad
    if(PrecioLista==1)
    {
      Precios<-seq((4+9),TamSerie,nvar) #Índices en los que está la variable PrecioLista
    }
    
    #FORMATO COMPLET0
    #AGREGAMOS TÍTULOS
    writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    writeData(wb, CATEGORIAS[n], "MÁRGENES DE UTILIDAD %", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN PIEZAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    if(PrecioLista==1)
    {
      writeData(wb, CATEGORIAS[n], "PRECIOS DE LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+8*(m+1))
    }
    if( Disponibilidad==1 & PrecioLista==1)
    {
      # solo si hay Disponibilidad
      writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
                startRow = 1, startCol = SeccionGralFin[3])
    }
    
    #ESCRIBIMOS LAS TABLAS
    if(Disponibilidad==1 & PrecioLista==1)
    {
      writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    }
    if(Disponibilidad==0 & PrecioLista==1)
    {
      writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    }
    if(Disponibilidad==0 & PrecioLista==0)
    {
      writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipPiezas,ParticipVentas,ParticipUtilidades)]), startRow = 2, startCol = SeccionGralInicio[3])
    }
  
    #AGREGAMOS ESTILOS
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:SeccionGralFin[3]))
    addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[3]:SeccionGralFin[3], gridExpand = TRUE)
    
    setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:SeccionGralFin[3]), widths = 17)
    addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:SeccionGralFin[3]))
    
    # 
    # #FORMATO IRVING
    # #ESCRIBIMOS LOS TÍTULOS DE LAS SUBSECCIONES DE LA SECCIÓN
    # writeData(wb, CATEGORIAS[n], "CLASIFICACIONES", startRow = 1, startCol = SeccionGralInicio[3]+5)
    # writeData(wb, CATEGORIAS[n], "PIEZAS VENDIDAS", startRow = 1, startCol = SeccionGralInicio[3]+5+(m+1))
    # writeData(wb, CATEGORIAS[n], "VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+2*(m+1))
    # writeData(wb, CATEGORIAS[n], "UTILIDADES", startRow = 1, startCol = SeccionGralInicio[3]+5+3*(m+1))
    # writeData(wb, CATEGORIAS[n], "MÁRGENES DE UTILIDAD", startRow = 1, startCol = SeccionGralInicio[3]+5+4*(m+1))
    # writeData(wb, CATEGORIAS[n], "PARTICIPACIÓN VENTAS", startRow = 1, startCol = SeccionGralInicio[3]+5+5*(m+1))
    # if(PrecioLista==1)
    # {
    #   writeData(wb, CATEGORIAS[n], "PRECIOS LISTA", startRow = 1, startCol = SeccionGralInicio[3]+5+6*(m+1))
    # }
    # if(Disponibilidad==1 & PrecioLista==1)
    # {
    #   # solo si hay Disponibilidad
    #   writeData(wb, CATEGORIAS[n], names(ListaTablasCategoria[[n]][["ResultadosMes"]])[ncol(ListaTablasCategoria[[n]][["ResultadosMes"]])], 
    #             startRow = 1, startCol = SeccionGralInicio[3]+5+7*(m+1))
    # }
    # 
    # 
    # 
    # #ESCRIBIMOS LAS TABLAS
    # if(Disponibilidad==1 & PrecioLista==1)
    # {
    #   writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios,(TamSerie+1))]), startRow = 2, startCol = SeccionGralInicio[3])
    # }
    # if(Disponibilidad==0 & PrecioLista==1)
    # {
    #   writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas,Precios)]), startRow = 2, startCol = SeccionGralInicio[3])
    #   
    # }
    # if(Disponibilidad==0 & PrecioLista==0)
    # {
    #   writeData(wb, CATEGORIAS[n], cbind(ClasifTotal=rep(""), ListaTablasCategoria[[n]][["SerieResultadosMensuales"]][,c(1:4,Letras,Piezas,Ventas,Utilidades,MUs,ParticipVentas)]), startRow = 2, startCol = SeccionGralInicio[3])
    # }
    # 
    # 
    # 
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad))))# +1 si hay disp, +0 si no la hay
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad))))# +1 si hay disp, +0 si no la hay
    # addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:(nfilTabSerie+2), cols = SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad)), gridExpand = TRUE)# +1 si hay disp, +0 si no la hay
    # 
    # setColWidths(wb,CATEGORIAS[n], cols = c(SeccionGralInicio[3]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1)+(Disponibilidad))), widths = 17)# +1 si hay disp, +0 si no la hay
    # addFilter(wb, CATEGORIAS[n], rows = 2, cols = c(SeccionGralInicio[1]:(SeccionGralInicio[3]+5+(6+PrecioLista)*(m+1))))
    # 
    
    
    
    
    
    
  }
  #deleteData()
  #sheetVisibility()
  
  cat("GUARDANDO\n")
  saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)
  #do.call(paste0, as.list(c("a","b","c","d")))
  #return(wb)
}













#PARA EXPORTAR A EXCEL (DESACTIVADA)
ExportarTablasExcel<-function(Resumen,ListaTablasCategoria, NombreArchivo)
{
    EstiloMesTop<-createStyle(fontSize = 20,fgFill = "#a5c6fa",border = "TopBottom")
    EstiloAcumuladoTop<-createStyle(fontSize = 20,fgFill = "#3cc962",border = "TopBottom")
    EstiloResultadosMensualesTop<-createStyle(fontSize = 20,fgFill = "#e03442",border = "TopBottom")
    
    EstiloPorPzas<-createStyle(fgFill = "#f5f5f5",border = "TopBottom", halign = "left")
    EstiloPorVenta<-createStyle(fgFill = "#ebebeb",border = "TopBottom", halign = "left")
    EstiloPorUtilidad<-createStyle(fgFill = "#dedede",border = "TopBottom", halign = "left")
    EstiloResultado<-createStyle(fontColour = "#ffffff" ,fgFill = "#737373",border = "TopBottom", halign = "left")
    EstiloResultadosMensuales<-createStyle(fontColour = "#ffffff" ,fgFill = "#737373",border = "TopBottom", halign = "left")
    
    EstiloMesBody<-createStyle(fgFill = "#edf3fc",halign = "left", numFmt = "#,##0;-#,##0")
    EstiloAcumuladoBody<-createStyle(fgFill = "#e1fae8",halign = "left", numFmt = "#,##0;-#,##0")
    EstiloResultadosMensualesBody<-createStyle(fgFill = "#ffe8ea",halign = "left", numFmt = "#,##0;-#,##0")
    
    
    wb<-createWorkbook()
    
        
        addWorksheet(wb, "RESUMEN", gridLines = FALSE)
        
        #"MES SELECCIONADO"
        writeData(wb, "RESUMEN", Resumen[[1]], startRow = 1, startCol = 1)
        writeData(wb, "RESUMEN", cbind(PorPzas=rep(""), Resumen[[2]],
                                           PorVenta=rep(""), Resumen[[3]],
                                           PorUtilidad=rep(""), Resumen[[4]],
                                           ClasifTotal=rep(""), Resumen[[5]]), startRow = 2, startCol = 1)
        addStyle(wb, "RESUMEN", style = EstiloMesTop, rows = 1, cols = c(1:34))
        addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(1:10))
        addStyle(wb, "RESUMEN", style = EstiloPorVenta, rows = 2, cols = c(11:20))
        addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(21:30))
        addStyle(wb, "RESUMEN", style = EstiloResultado, rows = 2, cols = c(31:34))
        addStyle(wb, "RESUMEN", style = EstiloMesBody, rows = 3:300, cols = 1:34, gridExpand = TRUE)
        
        #"ACUMULADO"
        writeData(wb, "RESUMEN", "ACUMULADO", startRow = 1, startCol = 35)
        writeData(wb, "RESUMEN", cbind(PorPzas=rep(""), Resumen[[6]],
                                           PorVenta=rep(""), Resumen[[7]],
                                           PorUtilidad=rep(""), Resumen[[8]],
                                           ClasifTotal=rep(""), Resumen[[9]]), startRow = 2, startCol = 35)
        addStyle(wb, "RESUMEN", style = EstiloAcumuladoTop, rows = 1, cols = c(35:68))
        addStyle(wb, "RESUMEN", style = EstiloPorPzas, rows = 2, cols = c(35:44))
        addStyle(wb, "RESUMEN", style = EstiloPorVenta, rows = 2, cols = c(45:54))
        addStyle(wb, "RESUMEN", style = EstiloPorUtilidad, rows = 2, cols = c(55:64))
        addStyle(wb, "RESUMEN", style = EstiloResultado, rows = 2, cols = c(65:68))
        addStyle(wb, "RESUMEN", style = EstiloAcumuladoBody, rows = 3:300, cols = 35:68, gridExpand = TRUE)
        
        
        #"TODOS LOS MESES"
        writeData(wb, "RESUMEN", "RESULTADOS MENSUALES", startRow = 1, startCol = 69)
        writeData(wb, "RESUMEN", cbind(ClasifABC=rep(""), Resumen[[10]]), startRow = 2, startCol = 69)
        addStyle(wb, "RESUMEN", style = EstiloResultadosMensualesTop, rows = 1, cols = c(69:84))
        addStyle(wb, "RESUMEN", style = EstiloResultadosMensuales, rows = 2, cols = c(69:84))
        addStyle(wb, "RESUMEN", style = EstiloResultadosMensualesBody, rows = 3:300, cols = 69:84, gridExpand = TRUE)
        
        
        setColWidths(wb,"RESUMEN", cols = c(1,10,11,20,21,30,31,34,35,44,45,54,55,64,65,68,69), widths = 15)
        
        
        CATEGORIAS<-Resumen[[11]]
        for(n in c(1:length(CATEGORIAS)))
        {
            addWorksheet(wb, CATEGORIAS[n], gridLines = FALSE)
            
            #"MES SELECCIONADO"
            writeData(wb, CATEGORIAS[n], ListaTablasCategoria[[n]][[1]], startRow = 1, startCol = 1)
            writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][[2]],
                                               PorVenta=rep(""), ListaTablasCategoria[[n]][[3]],
                                               PorUtilidad=rep(""), ListaTablasCategoria[[n]][[4]],
                                               ClasifTotal=rep(""), ListaTablasCategoria[[n]][[5]]), startRow = 2, startCol = 1)
            addStyle(wb, CATEGORIAS[n], style = EstiloMesTop, rows = 1, cols = c(1:34))
            addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(1:10))
            addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(11:20))
            addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(21:30))
            addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(31:34))
            addStyle(wb, CATEGORIAS[n], style = EstiloMesBody, rows = 3:300, cols = 1:34, gridExpand = TRUE)
            
            #"ACUMULADO"
            writeData(wb, CATEGORIAS[n], "ACUMULADO", startRow = 1, startCol = 35)
            writeData(wb, CATEGORIAS[n], cbind(PorPzas=rep(""), ListaTablasCategoria[[n]][[6]],
                                               PorVenta=rep(""), ListaTablasCategoria[[n]][[7]],
                                               PorUtilidad=rep(""), ListaTablasCategoria[[n]][[8]],
                                               ClasifTotal=rep(""), ListaTablasCategoria[[n]][[9]]), startRow = 2, startCol = 35)
            addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoTop, rows = 1, cols = c(35:68))
            addStyle(wb, CATEGORIAS[n], style = EstiloPorPzas, rows = 2, cols = c(35:44))
            addStyle(wb, CATEGORIAS[n], style = EstiloPorVenta, rows = 2, cols = c(45:54))
            addStyle(wb, CATEGORIAS[n], style = EstiloPorUtilidad, rows = 2, cols = c(55:64))
            addStyle(wb, CATEGORIAS[n], style = EstiloResultado, rows = 2, cols = c(65:68))
            addStyle(wb, CATEGORIAS[n], style = EstiloAcumuladoBody, rows = 3:300, cols = 35:68, gridExpand = TRUE)
            
            
            #"TODOS LOS MESES"
            writeData(wb, CATEGORIAS[n], "RESULTADOS MENSUALES", startRow = 1, startCol = 69)
            writeData(wb, CATEGORIAS[n], cbind(ClasifABC=rep(""), ListaTablasCategoria[[n]][[10]]), startRow = 2, startCol = 69)
            addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesTop, rows = 1, cols = c(69:84))
            addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensuales, rows = 2, cols = c(69:84))
            addStyle(wb, CATEGORIAS[n], style = EstiloResultadosMensualesBody, rows = 3:300, cols = 69:84, gridExpand = TRUE)
            
            
            setColWidths(wb,CATEGORIAS[n], cols = c(1,10,11,20,21,30,31,34,35,44,45,54,55,64,65,68,69), widths = 15)
            
            
        }
        saveWorkbook(wb, file = paste0(NombreArchivo, ".xlsx"), overwrite = TRUE)

    
    #do.call(paste0, as.list(c("a","b","c","d")))
}

rm(
la,
p1,
p2,
p3,
pa,
pe,
pii,
po,
pu
)
rm(ExportarTablasExcelSimple)

pii<-data.frame(col3="c",col2="b", col1="a")
po<-data.frame("col2","col1","col3")
po<-data.frame(col2=c(),col1=c(),col3=c())

po
rbind(pa,pe,pii,po,pu)
pu<-data.frame(c(),c(),c())
pu
names(pu)<-c("col1","col2","col3")
class(pu)
pu<-data.frame("Categoria", NA, NA)
pu
names(pu)<-c("col1","col2","col3")
x
xxx
class(xxx)
rm(renglon)
 

list(c(1:10))
as_vector(as.list(c(1:10)))

pii<-data.frame(col3=c(3,"c"),col2=c(2,"b"), col1=c(1,"a"))
pii
tr_pii<-transpose(pii)
tr_pii
tr_pii1<-tr_pii[[1]]
tr_pii1
class(tr_pii1)
tr_pii2<-as_vector(tr_pii1)
tr_pii2
class(as_vector(tr_pii2))
tr_pii2["col3"]
#lo anterior es eq a hacer lo sig
a<-c(1:3)
names(a)<-c("col1","col2","col3")
a
#a las listas se les puede asignar un nombre
a<-list(objeto0="fecha" ,objeto1=c(1:10),objeto2=c(11:20), objeto3=c("a","b","c"))
a
a["objeto0"]

#conclusión para separar DF en renglones y hacer una lista 
pii<-data.frame(col3=c(3,"c"),col2=c(2,"b"), col1=c(1,"a"))
pii
tr_pii<-transpose(pii)
tr_pii
class(tr_pii)
class(tr_pii[[1]])
renglones<-lapply(transpose(pii),as_vector)#renglones como vector
renglones
renglones<-lapply(transpose(pii),as.data.frame)#renglones como data.frame de un renglón
renglones
#Ctrl + Shift + C
# 
# asssasa
# sasasas
# 
# asasas
# 
# 
# asasas
# asasas

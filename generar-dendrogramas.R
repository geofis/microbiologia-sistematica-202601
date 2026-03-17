# ============================================================
# ANÁLISIS DE AGRUPAMIENTO PARA SUDOKUS EN EXCEL SEMIESTRUCTURADO
# ============================================================

# Paquetes
library(readxl)
library(vegan)
library(dendextend)

# ------------------------------------------------------------
# Función auxiliar: normalizar texto
# ------------------------------------------------------------
normalizar_texto <- function(x) {
  x <- as.character(x)
  x <- trimws(x)
  x <- gsub("\\s+", " ", x)
  toupper(x)
}

# ------------------------------------------------------------
# Función auxiliar: localizar el encabezado del sudoku
# Busca la celda cuyo contenido coincide con el código
# y verifica que a su derecha estén A, B, C, ..., I
# ------------------------------------------------------------
localizar_sudoku <- function(hoja, codigo) {
  hoja_chr <- as.data.frame(lapply(hoja, as.character), stringsAsFactors = FALSE)
  mat_chr  <- as.matrix(hoja_chr)
  
  mat_norm <- apply(mat_chr, c(1, 2), normalizar_texto)
  codigo_norm <- normalizar_texto(codigo)
  
  pos <- which(mat_norm == codigo_norm, arr.ind = TRUE)
  
  if (nrow(pos) == 0) {
    stop(paste0("No se encontró el código del sudoku: '", codigo, "'"))
  }
  
  encabezado_esperado <- LETTERS[1:9]
  
  for (i in seq_len(nrow(pos))) {
    fila <- pos[i, 1]
    col  <- pos[i, 2]
    
    if ((col + 9) <= ncol(mat_norm)) {
      encabezado <- mat_norm[fila, (col + 1):(col + 9)]
      
      if (all(encabezado == encabezado_esperado)) {
        return(list(fila = fila, col = col))
      }
    }
  }
  
  stop(
    paste0(
      "Se encontró el texto '", codigo,
      "', pero no se identificó un encabezado válido A:I a su derecha."
    )
  )
}

# ------------------------------------------------------------
# Función auxiliar: extraer el bloque 9x9 del sudoku
# ------------------------------------------------------------
extraer_sudoku <- function(hoja, codigo) {
  loc <- localizar_sudoku(hoja, codigo)
  
  fila_ini <- loc$fila + 1
  fila_fin <- loc$fila + 9
  col_ini  <- loc$col + 1
  col_fin  <- loc$col + 9
  
  if (fila_fin > nrow(hoja) || col_fin > ncol(hoja)) {
    stop("El bloque 9x9 del sudoku excede los límites de la hoja.")
  }
  
  bloque <- hoja[fila_ini:fila_fin, col_ini:col_fin, drop = FALSE]
  
  bloque_num <- as.data.frame(
    lapply(bloque, function(x) suppressWarnings(as.numeric(as.character(x))))
  )
  
  colnames(bloque_num) <- LETTERS[1:9]
  rownames(bloque_num) <- 1:9
  
  if (any(is.na(as.matrix(bloque_num)))) {
    stop(
      paste0(
        "El bloque extraído para '", codigo,
        "' contiene valores no numéricos o celdas vacías."
      )
    )
  }
  
  return(as.matrix(bloque_num))
}

# ------------------------------------------------------------
# Función principal
# archivo        : ruta al xlsx
# codigo         : código del sudoku, por ejemplo "SK  (I)"
# hoja           : hoja del Excel
# k              : número de grupos
# distancia      : "euclidean" o "bray"
# metodo         : método de hclust
# graficar       : TRUE/FALSE
# motor_grafico  : "base" o "dendextend"
# ------------------------------------------------------------
analizar_sudoku <- function(archivo,
                            codigo,
                            hoja = 1,
                            k = 2,
                            distancia = c("euclidean", "bray"),
                            metodo = "average",
                            graficar = TRUE,
                            motor_grafico = c("base", "dendextend")) {
  
  distancia <- match.arg(distancia)
  motor_grafico <- match.arg(motor_grafico)
  
  # Leer hoja completa sin asumir estructura tabular
  datos <- read_excel(
    path = archivo,
    sheet = hoja,
    col_names = FALSE
  )
  
  # Extraer sudoku
  mat <- extraer_sudoku(datos, codigo)
  
  # Matriz de distancias
  if (distancia == "euclidean") {
    dist_mat <- dist(mat, method = "euclidean")
  } else if (distancia == "bray") {
    dist_mat <- vegdist(mat, method = "bray")
  }
  
  # Clustering jerárquico
  hc <- hclust(dist_mat, method = metodo)
  
  # Crear objeto dendrograma
  obj_dend <- as.dendrogram(hc)
  
  # Asignación de grupos
  grupos <- cutree(hc, k = k)
  
  # Graficar
  if (graficar) {
    
    if (motor_grafico == "base") {
      
      plot(
        hc,
        main = paste0("Dendrograma - ", codigo),
        xlab = "",
        sub = paste0("Distancia: ", distancia, " | Método: ", metodo),
        cex = 0.9
      )
      rect.hclust(hc, k = k)
      
    } else if (motor_grafico == "dendextend") {
      
      dend_col <- color_branches(obj_dend, k = k)
      dend_col <- set(dend_col, "labels_cex", 0.9)
      
      plot(
        dend_col,
        main = paste0("Dendrograma - ", codigo),
        ylab = "Altura"
      )
      
      rect.dendrogram(dend_col, k = k, border = 2:(k + 1))
    }
  }
  
  return(list(
    codigo = codigo,
    matriz = mat,
    distancia = dist_mat,
    hclust = hc,
    dendrograma = obj_dend,
    grupos = grupos
  ))
}
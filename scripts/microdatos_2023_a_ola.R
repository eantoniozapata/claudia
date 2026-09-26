# Convierte los microdatos de Latinobarómetro 2023 (Rdata) en porcentajes
# ponderados por país, con las mismas categorías que la herramienta online.
#
# La serie histórica de la herramienta online no trae 2023 desagregado por país.
# Este script reproduce el cálculo de la herramienta: porcentaje = suma de wt de
# cada respuesta / suma de wt del país, redondeado a un decimal. Verificado:
# el total regional coincide con el de la herramienta (Democracia 48,0 %;
# Grupos poderosos 72,1 %).
#
# Entrada: data/raw/lb2023/Latinobarometro_2023_Esp_Rdata_v1_0.rdata
#          (descomprimir F00017014-Latinobarometro_2023_Rdata_v1_0.zip)
# Salida : data/online/latinobarometro_2023_microdatos.csv
# Uso    : Rscript scripts/microdatos_2023_a_ola.R   (desde la raíz del proyecto)

ENTRADA <- "data/raw/lb2023/Latinobarometro_2023_Esp_Rdata_v1_0.rdata"
SALIDA <- "data/online/latinobarometro_2023_microdatos.csv"

e <- new.env()
load(ENTRADA, envir = e)
d <- get(grep("^Latinobarometro_2023", ls(e), value = TRUE)[1], envir = e)
stopifnot(all(d$numinves %in% c(23, 2023)))  # el archivo codifica 2023 como 23

# Nombres de país como en la herramienta online (idenpa = ISO 3166 numérico)
paises <- c("32" = "Argentina", "68" = "Bolivia", "76" = "Brasil", "152" = "Chile", "170" = "Colombia",
            "188" = "Costa Rica", "214" = "Rep. Dominicana", "218" = "Ecuador", "222" = "El Salvador",
            "320" = "Guatemala", "340" = "Honduras", "484" = "México", "558" = "Nicaragua",
            "591" = "Panamá", "600" = "Paraguay", "604" = "Perú", "724" = "España", "858" = "Uruguay",
            "862" = "Venezuela")

# Categorías de la herramienta online a partir de los códigos de los microdatos
preguntas <- list(
  "Apoyo democracia" = list(
    var = "P10STGBS",
    codigos = list("Democracia" = 1, "Gobierno Autoritario" = 2, "Da lo mismo" = 3,
                   "No sabe" = -1, "No contesta" = -2)),
  "Grupos poderosos" = list(
    var = "P12ST",
    # -5 = No sabe / No contesta; 0 no tiene etiqueta en el archivo (82 casos). La herramienta
    # online los cuenta como NS/NR: solo así su total regional (3,5 %) coincide.
    codigos = list("Grupos poderosos en su propio beneficio" = 1, "Para el bien de todo el pueblo" = 2,
                   "No sabe; no responde" = c(-1, -2, -5, 0)))
)

filas <- list()
for (variable in names(preguntas)) {
  p <- preguntas[[variable]]
  x <- d[[p$var]]
  stopifnot(all(x %in% unlist(p$codigos)))  # ningún código queda sin categoría
  for (cod in names(paises)) {
    en_pais <- d$idenpa == as.numeric(cod)
    if (!any(en_pais)) next
    for (cat in names(p$codigos)) {
      pct <- sum(d$wt[en_pais & x %in% p$codigos[[cat]]]) / sum(d$wt[en_pais])
      filas[[length(filas) + 1]] <- data.frame(
        variable = variable, pais = paises[[cod]], anio = 2023L, categoria = cat,
        porcentaje = round(100 * pct, 1), n = sum(en_pais))
    }
  }
  total <- tapply(d$wt, sapply(x, function(v) names(p$codigos)[sapply(p$codigos, function(c) v %in% c)]), sum)
  cat(variable, "- total regional ponderado (%):\n")
  print(round(100 * total / sum(d$wt), 1))
}
res <- do.call(rbind, filas)
write.csv(res, SALIDA, row.names = FALSE, fileEncoding = "UTF-8")
cat(sprintf("%s: %d filas, %d países\n", SALIDA, nrow(res), length(unique(res$pais))))

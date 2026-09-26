# =============================================================================
# Apoyo a la democracia y desigualdad en América Latina
# Estadística descriptiva y regresiones MCO sobre el panel país-año
#
# Entrada : excel/latinobarometro_panel.xlsx (libro armado con la herramienta
#           online de Latinobarómetro, Banco Mundial, CEPAL y Transparency Int.)
# Salida  : carpeta resultados/ (tablas en Excel y CSV, gráficos en PNG)
#
# El script lee SOLO las hojas con datos de origen (Datos, Gini, Puente_CPI,
# Corrupcion, Controles) y reconstruye todas las variables en R, así no depende
# de que Excel haya recalculado las fórmulas del libro.
#
# Uso: abrir en RStudio con el directorio de trabajo en la raíz del proyecto
#      (o cambiar RUTA_EXCEL) y ejecutar todo (Ctrl + Shift + Enter).
# =============================================================================

# ---- 0. Paquetes y rutas -----------------------------------------------------
paquetes <- c("readxl", "dplyr", "tidyr", "ggplot2", "sandwich", "lmtest", "car", "writexl")
faltan <- paquetes[!vapply(paquetes, requireNamespace, logical(1), quietly = TRUE)]
if (length(faltan) > 0) install.packages(faltan)
invisible(lapply(paquetes, library, character.only = TRUE))

RUTA_EXCEL <- "excel/latinobarometro_panel.xlsx"
SALIDA <- "resultados"
dir.create(SALIDA, showWarnings = FALSE)
dir.create(file.path(SALIDA, "graficos"), showWarnings = FALSE)
theme_set(theme_minimal(base_size = 11))

guardar_grafico <- function(g, nombre, ancho = 8, alto = 5) {
  ggsave(file.path(SALIDA, "graficos", nombre), g, width = ancho, height = alto, dpi = 150, bg = "white")
}

# ---- 1. Lectura de los datos de origen ---------------------------------------
# Se renombran las columnas por posición para evitar problemas con tildes.

# 1.1 Latinobarómetro (porcentajes ponderados de la herramienta online)
datos <- read_excel(RUTA_EXCEL, sheet = "Datos")
names(datos) <- c("variable", "pais", "anio", "categoria", "prop", "n")

# 1.2 Gini (Banco Mundial) y datos para interpolar
gini <- read_excel(RUTA_EXCEL, sheet = "Gini", range = cell_cols("A:I"))
names(gini) <- c("pais", "anio", "gini", "fuente", "clave", "anio_ant", "gini_ant", "anio_sig", "gini_sig")

# 1.3 Puente CPI 2011 (método antiguo, 0-10) -> 2012 (método nuevo, 0-100)
puente <- read_excel(RUTA_EXCEL, sheet = "Puente_CPI", range = cell_cols("A:D"))
names(puente) <- c("pais", "iso3", "cpi2011", "cpi2012")

# 1.4 CPI por país-año
corrupcion <- read_excel(RUTA_EXCEL, sheet = "Corrupcion", range = cell_cols("A:D"))
names(corrupcion) <- c("pais", "anio", "cpi_datahub", "cpi_hdx")

# 1.5 Controles
controles <- read_excel(RUTA_EXCEL, sheet = "Controles", range = cell_cols("A:E"))
names(controles) <- c("pais", "anio", "homicidios", "crec_pib", "gasto_social")

# ---- 2. Construcción del panel -----------------------------------------------

# 2.1 Porcentajes por categoría (una columna por categoría)
cat_ancho <- datos %>%
  mutate(col = paste(variable, categoria, sep = "|")) %>%
  select(pais, anio, col, prop) %>%
  pivot_wider(names_from = col, values_from = prop)

g <- function(v, c) cat_ancho[[paste(v, c, sep = "|")]]
panel <- cat_ancho %>%
  transmute(
    pais, anio,
    # Criterio A: % sobre el total (incluye NS/NR)
    apoyo_A  = 100 * g("Apoyo democracia", "Democracia"),
    grupos_A = 100 * g("Grupos poderosos", "Grupos poderosos en su propio beneficio"),
    # Criterio B: % sobre respuestas válidas (excluye NS/NR)
    apoyo_B  = 100 * g("Apoyo democracia", "Democracia") /
      (g("Apoyo democracia", "Democracia") + g("Apoyo democracia", "Gobierno Autoritario") +
         g("Apoyo democracia", "Da lo mismo")),
    grupos_B = 100 * g("Grupos poderosos", "Grupos poderosos en su propio beneficio") /
      (g("Grupos poderosos", "Grupos poderosos en su propio beneficio") +
         g("Grupos poderosos", "Para el bien de todo el pueblo"))
  )

# 2.2 Gini observado e interpolado (lineal entre años observados, sin extrapolar)
gini <- gini %>%
  mutate(gini_interp = ifelse(!is.na(gini), gini,
                              gini_ant + (gini_sig - gini_ant) * (anio - anio_ant) / (anio_sig - anio_ant))) %>%
  select(pais, anio, gini, gini_interp)

# 2.3 CPI: versión 0-100 desde 2012 y versión empalmada por regresión
m_puente <- lm(cpi2012 ~ cpi2011, data = puente)
a <- coef(m_puente)[1]; b <- coef(m_puente)[2]
cat(sprintf("Empalme CPI: CPI2012 = %.3f + %.3f x CPI2011  (R2 = %.3f, N = %d países)\n",
            a, b, summary(m_puente)$r.squared, nobs(m_puente)))
corrupcion <- corrupcion %>%
  mutate(cpi     = ifelse(anio >= 2012, cpi_hdx, NA),
         cpi_reg = ifelse(anio <= 2011, a + b * cpi_datahub, cpi_hdx)) %>%
  select(pais, anio, cpi, cpi_reg)

# 2.4 Unión de todo
panel <- panel %>%
  left_join(gini, by = c("pais", "anio")) %>%
  left_join(corrupcion, by = c("pais", "anio")) %>%
  left_join(controles, by = c("pais", "anio")) %>%
  arrange(pais, anio)

# Control de calidad contra los valores revisados en el Excel
arg24 <- filter(panel, pais == "Argentina", anio == 2024)
stopifnot(abs(arg24$apoyo_A - 74.6) < 0.05, abs(arg24$apoyo_B - 77.466) < 0.01)
cat(sprintf("Panel: %d filas país-año, %d países, años %d-%d\n",
            nrow(panel), n_distinct(panel$pais), min(panel$anio), max(panel$anio)))
write.csv(panel, file.path(SALIDA, "panel_construido.csv"), row.names = FALSE, fileEncoding = "UTF-8")

etiquetas <- c(
  apoyo_B = "Apoyo a la democracia (%, B)", apoyo_A = "Apoyo a la democracia (%, A)",
  grupos_B = "Gobierno de grupos poderosos (%, B)", grupos_A = "Gobierno de grupos poderosos (%, A)",
  gini = "Gini (0-100)", gini_interp = "Gini interpolado (0-100)",
  cpi = "CPI 0-100 (desde 2012)", cpi_reg = "CPI empalmado por regresión (0-100)",
  homicidios = "Homicidios por 100.000 hab.", crec_pib = "Crecimiento del PIB (%)",
  gasto_social = "Gasto social (% PIB, gob. central)"
)
vars_desc <- names(etiquetas)

# ---- 3. Estadística descriptiva ----------------------------------------------

# 3.1 Resumen de cada variable
descriptiva <- bind_rows(lapply(vars_desc, function(v) {
  x <- panel[[v]]
  data.frame(Variable = etiquetas[[v]], N = sum(!is.na(x)),
             Media = mean(x, na.rm = TRUE), DE = sd(x, na.rm = TRUE),
             Min = min(x, na.rm = TRUE), P25 = quantile(x, .25, na.rm = TRUE),
             Mediana = median(x, na.rm = TRUE), P75 = quantile(x, .75, na.rm = TRUE),
             Max = max(x, na.rm = TRUE),
             Anios = paste(range(panel$anio[!is.na(x)]), collapse = "-"), row.names = NULL)
})) %>% mutate(across(where(is.numeric) & !N, ~ round(.x, 2)))
cat("\n=== Estadística descriptiva (todas las filas país-año con dato) ===\n")
print(descriptiva, row.names = FALSE)

# 3.2 Cobertura: número de países con dato por año
cobertura <- panel %>%
  group_by(anio) %>%
  summarise(across(all_of(vars_desc), ~ sum(!is.na(.x))), .groups = "drop")

# 3.3 Promedios por país y por año
por_pais <- panel %>%
  group_by(pais) %>%
  summarise(olas = sum(!is.na(apoyo_B)),
            across(c(apoyo_B, grupos_B, gini, cpi, homicidios, crec_pib, gasto_social),
                   ~ ifelse(all(is.na(.x)), NA, round(mean(.x, na.rm = TRUE), 1))), .groups = "drop") %>%
  arrange(desc(apoyo_B))
por_anio <- panel %>%
  group_by(anio) %>%
  summarise(paises = sum(!is.na(apoyo_B)),
            across(c(apoyo_B, grupos_B, gini, cpi), ~ ifelse(all(is.na(.x)), NA, round(mean(.x, na.rm = TRUE), 1))), .groups = "drop")
cat("\n=== Promedio por país (ordenado por apoyo a la democracia, criterio B) ===\n")
print(as.data.frame(por_pais), row.names = FALSE)

# 3.4 Correlaciones (pares completos)
vars_cor <- c("apoyo_B", "grupos_B", "gini", "cpi", "cpi_reg", "homicidios", "crec_pib", "gasto_social")
correl <- round(cor(panel[vars_cor], use = "pairwise.complete.obs"), 2)
cat("\n=== Correlaciones (pares completos) ===\n")
print(correl)

# 3.5 Gráficos descriptivos
largo <- panel %>%
  select(pais, anio, all_of(c("apoyo_B", "grupos_B", "gini", "cpi", "homicidios", "crec_pib", "gasto_social"))) %>%
  pivot_longer(-c(pais, anio), names_to = "variable", values_to = "valor") %>%
  filter(!is.na(valor)) %>%
  mutate(variable = etiquetas[variable])

guardar_grafico(
  ggplot(largo, aes(valor)) + geom_histogram(bins = 20, fill = "#2a6f97", colour = "white") +
    facet_wrap(~ variable, scales = "free") + labs(title = "Distribución de las variables", x = NULL, y = "Frecuencia"),
  "01_histogramas.png", 10, 7)

guardar_grafico(
  ggplot(filter(panel, !is.na(apoyo_B)), aes(anio, apoyo_B)) + geom_line(colour = "#2a6f97") +
    geom_point(size = .8, colour = "#2a6f97") + facet_wrap(~ pais) +
    labs(title = "Apoyo a la democracia por país (criterio B)", x = NULL, y = "%"),
  "02_apoyo_por_pais.png", 11, 8)

guardar_grafico(
  panel %>% filter(pais != "España", !is.na(apoyo_B)) %>%
    group_by(anio) %>% summarise(apoyo = mean(apoyo_B), grupos = mean(grupos_B, na.rm = TRUE)) %>%
    pivot_longer(-anio) %>% filter(!is.nan(value)) %>%
    mutate(name = ifelse(name == "apoyo", "Apoyo a la democracia", "Gobierno de grupos poderosos")) %>%
    ggplot(aes(anio, value, colour = name)) + geom_line(linewidth = 1) + geom_point() +
    labs(title = "Promedio simple de América Latina (criterio B)", x = NULL, y = "%", colour = NULL) +
    theme(legend.position = "bottom"),
  "03_promedio_regional.png")

guardar_grafico(
  ggplot(filter(panel, !is.na(apoyo_B)), aes(reorder(pais, apoyo_B, median), apoyo_B)) +
    geom_boxplot(fill = "#a9d6e5") + coord_flip() +
    labs(title = "Apoyo a la democracia por país, todas las olas (criterio B)", x = NULL, y = "%"),
  "04_boxplot_pais.png", 8, 6)

dispersion <- panel %>%
  select(pais, anio, apoyo_B, grupos_B, gini, cpi, homicidios, crec_pib, gasto_social) %>%
  pivot_longer(-c(pais, anio, apoyo_B), names_to = "variable", values_to = "x") %>%
  filter(!is.na(x), !is.na(apoyo_B)) %>%
  mutate(variable = etiquetas[variable])
guardar_grafico(
  ggplot(dispersion, aes(x, apoyo_B)) + geom_point(alpha = .5, colour = "#2a6f97") +
    geom_smooth(method = "lm", formula = y ~ x, colour = "#c1121f") +
    facet_wrap(~ variable, scales = "free_x") +
    labs(title = "Apoyo a la democracia (B) frente a cada variable explicativa", x = NULL, y = "Apoyo (%)"),
  "05_dispersion.png", 10, 7)

cor_largo <- as.data.frame(as.table(correl)) %>%
  mutate(across(c(Var1, Var2), ~ factor(etiquetas[as.character(.x)], levels = etiquetas[vars_cor])))
guardar_grafico(
  ggplot(cor_largo, aes(Var1, Var2, fill = Freq)) + geom_tile() + geom_text(aes(label = Freq), size = 3) +
    scale_fill_gradient2(low = "#c1121f", mid = "white", high = "#2a6f97", limits = c(-1, 1)) +
    labs(title = "Matriz de correlaciones", x = NULL, y = NULL, fill = "r") +
    theme(axis.text.x = element_text(angle = 40, hjust = 1)),
  "06_correlaciones.png", 9, 7)

# ---- 4. Regresiones MCO -------------------------------------------------------
# Errores estándar agrupados por país (robustos a heterocedasticidad y a la
# correlación de los errores de un mismo país en el tiempo). Con ~18 países son
# pocos grupos: interprete la significancia con cautela.

ee_cluster <- function(m) vcovCL(m, cluster = m$pais, type = "HC1")

estrellas <- function(p) ifelse(p < .01, "***", ifelse(p < .05, "**", ifelse(p < .1, "*", "")))

# Tabla de regresiones: coeficiente con estrellas y error estándar entre paréntesis
tabla_modelos <- function(modelos, etiquetas_x) {
  terminos <- unique(unlist(lapply(modelos, function(m) names(coef(m)))))
  terminos <- terminos[!grepl("^factor\\(", terminos)]
  # Orden fijo: el de etiquetas_x, luego cualquier otro término y al final la constante
  terminos <- c(intersect(names(etiquetas_x), terminos),
                setdiff(terminos, c(names(etiquetas_x), "(Intercept)")), "(Intercept)")
  filas <- list()
  for (t in terminos) {
    coefs <- ses <- character(length(modelos))
    for (k in seq_along(modelos)) {
      ct <- coeftest(modelos[[k]], vcov. = ee_cluster(modelos[[k]]))
      if (t %in% rownames(ct)) {
        coefs[k] <- sprintf("%.3f%s", ct[t, 1], estrellas(ct[t, 4]))
        ses[k] <- sprintf("(%.3f)", ct[t, 2])
      }
    }
    nombre <- if (t == "(Intercept)") "Constante" else if (t %in% names(etiquetas_x)) etiquetas_x[[t]] else t
    filas[[length(filas) + 1]] <- c(nombre, coefs)
    filas[[length(filas) + 1]] <- c("", ses)
  }
  ef <- function(patron) sapply(modelos, function(m) ifelse(any(grepl(patron, names(coef(m)))), "Sí", "No"))
  pie <- list(
    c("Efectos fijos de año", ef("factor\\(anio\\)")),
    c("Efectos fijos de país", ef("factor\\(pais\\)")),
    c("Observaciones (país-año)", sapply(modelos, nobs)),
    c("Países", sapply(modelos, function(m) n_distinct(m$pais))),
    c("R2", sapply(modelos, function(m) sprintf("%.3f", summary(m)$r.squared))),
    c("R2 ajustado", sapply(modelos, function(m) sprintf("%.3f", summary(m)$adj.r.squared)))
  )
  tabla <- as.data.frame(do.call(rbind, c(filas, pie)), stringsAsFactors = FALSE)
  names(tabla) <- c("Variable", names(modelos))
  tabla
}

# Cada modelo guarda el país de sus filas para agrupar los errores. Los datos
# llegan ya filtrados a filas completas, así que lm no descarta ninguna.
ajustar <- function(formula, datos) {
  m <- lm(formula, data = datos)
  stopifnot(nobs(m) == nrow(datos))
  m$pais <- datos$pais
  m
}

# 4.1 Modelos principales: criterio B, Gini observado, CPI 0-100 (2012 en adelante)
# Muestra común = filas completas en el modelo más grande, para que los cambios en
# los coeficientes entre columnas no se deban a cambios de muestra.
vars_principal <- c("apoyo_B", "gini", "grupos_B", "cpi", "homicidios", "crec_pib", "gasto_social")
muestra <- panel %>% filter(if_all(all_of(vars_principal), ~ !is.na(.x)))
cat(sprintf("\nMuestra de los modelos principales: %d observaciones, %d países, años %s\n",
            nrow(muestra), n_distinct(muestra$pais), paste(sort(unique(muestra$anio)), collapse = ", ")))

principales <- list(
  "(1)" = ajustar(apoyo_B ~ gini, muestra),
  "(2)" = ajustar(apoyo_B ~ gini + grupos_B, muestra),
  "(3)" = ajustar(apoyo_B ~ gini + grupos_B + cpi, muestra),
  "(4)" = ajustar(apoyo_B ~ gini + grupos_B + cpi + homicidios + crec_pib + gasto_social, muestra),
  "(5)" = ajustar(apoyo_B ~ gini + grupos_B + cpi + homicidios + crec_pib + gasto_social + factor(anio), muestra)
)
etq_x <- c(gini = "Gini", gini_interp = "Gini interpolado", grupos_B = "Grupos poderosos (%)",
           grupos_A = "Grupos poderosos (%, criterio A)", cpi = "CPI 0-100", cpi_reg = "CPI empalmado (regresión)",
           homicidios = "Homicidios por 100.000", crec_pib = "Crecimiento del PIB (%)",
           gasto_social = "Gasto social (% PIB)")
tabla_principal <- tabla_modelos(principales, etq_x)
cat("\n=== Tabla 1. MCO: apoyo a la democracia (criterio B, %) ===\n")
cat("Errores estándar agrupados por país entre paréntesis. *** p<0,01; ** p<0,05; * p<0,1\n")
print(tabla_principal, row.names = FALSE)

# 4.2 Robustez
# (R1) criterio A; (R2) Gini interpolado; (R3) CPI empalmado por regresión con
# panel largo desde 2004 (sin gasto social, que empieza en 2010); (R4) = R3 con
# Gini interpolado; (R5) efectos fijos de país y año (usa solo variación dentro
# de cada país).
m_rob <- function(vars, formula) ajustar(formula, panel %>% filter(if_all(all_of(vars), ~ !is.na(.x))))
controles_x <- "homicidios + crec_pib + gasto_social"
robustez <- list(
  "(R1) Criterio A" = m_rob(c("apoyo_A", "gini", "grupos_A", "cpi", "homicidios", "crec_pib", "gasto_social"),
                            as.formula(paste("apoyo_A ~ gini + grupos_A + cpi +", controles_x, "+ factor(anio)"))),
  "(R2) Gini interp." = m_rob(c("apoyo_B", "gini_interp", "grupos_B", "cpi", "homicidios", "crec_pib", "gasto_social"),
                              as.formula(paste("apoyo_B ~ gini_interp + grupos_B + cpi +", controles_x, "+ factor(anio)"))),
  "(R3) CPI regresión" = m_rob(c("apoyo_B", "gini", "grupos_B", "cpi_reg", "homicidios", "crec_pib"),
                               apoyo_B ~ gini + grupos_B + cpi_reg + homicidios + crec_pib + factor(anio)),
  "(R4) CPI reg. + Gini interp." = m_rob(c("apoyo_B", "gini_interp", "grupos_B", "cpi_reg", "homicidios", "crec_pib"),
                                         apoyo_B ~ gini_interp + grupos_B + cpi_reg + homicidios + crec_pib + factor(anio)),
  "(R5) EF país y año" = m_rob(c("apoyo_B", "gini_interp", "grupos_B", "cpi_reg", "homicidios", "crec_pib"),
                               apoyo_B ~ gini_interp + grupos_B + cpi_reg + homicidios + crec_pib +
                                 factor(anio) + factor(pais))
)
tabla_robustez <- tabla_modelos(robustez, etq_x)
cat("\n=== Tabla 2. Robustez ===\n")
print(tabla_robustez, row.names = FALSE)

# ---- 5. Diagnóstico del modelo (4) --------------------------------------------
m4 <- principales[["(4)"]]
vif_m4 <- car::vif(m4)
bp_m4 <- lmtest::bptest(m4)
reset_m4 <- lmtest::resettest(m4, power = 2:3)
cook <- cooks.distance(m4)
influyentes <- muestra %>% mutate(cook = cook) %>% arrange(desc(cook)) %>%
  select(pais, anio, apoyo_B, cook) %>% filter(cook > 4 / nrow(muestra))
diagnostico <- data.frame(
  Prueba = c(paste("VIF", names(vif_m4)), "Breusch-Pagan (heterocedasticidad)", "RESET de Ramsey (forma funcional)"),
  Valor = round(c(vif_m4, bp_m4$statistic, reset_m4$statistic), 3),
  p_valor = c(rep(NA, length(vif_m4)), round(bp_m4$p.value, 4), round(reset_m4$p.value, 4))
)
cat("\n=== Diagnóstico del modelo (4) ===\n")
print(diagnostico, row.names = FALSE)
cat("VIF > 5 sugiere multicolinealidad; p < 0,05 en Breusch-Pagan indica heterocedasticidad\n",
    "(los errores agrupados ya la corrigen); p < 0,05 en RESET sugiere mala especificación.\n")
cat("Observaciones influyentes (distancia de Cook > 4/n):\n")
print(as.data.frame(influyentes), row.names = FALSE)

diag_df <- data.frame(ajustado = fitted(m4), residuo = resid(m4))
guardar_grafico(
  ggplot(diag_df, aes(ajustado, residuo)) + geom_point(colour = "#2a6f97") +
    geom_hline(yintercept = 0, linetype = 2) + geom_smooth(method = "loess", formula = y ~ x, se = FALSE, colour = "#c1121f") +
    labs(title = "Modelo (4): residuos frente a valores ajustados", x = "Valor ajustado", y = "Residuo"),
  "07_residuos_ajustados.png", 7, 5)
guardar_grafico(
  ggplot(diag_df, aes(sample = residuo)) + stat_qq(colour = "#2a6f97") + stat_qq_line() +
    labs(title = "Modelo (4): gráfico Q-Q de los residuos", x = "Cuantiles teóricos", y = "Cuantiles de los residuos"),
  "08_qq_residuos.png", 6, 5)

# ---- 6. Exportar ----------------------------------------------------------------
write_xlsx(list(
  "Descriptiva" = descriptiva,
  "Cobertura por anio" = cobertura,
  "Promedio por pais" = por_pais,
  "Promedio por anio" = por_anio,
  "Correlaciones" = cbind(Variable = rownames(correl), as.data.frame(correl)),
  "Tabla 1 MCO" = tabla_principal,
  "Tabla 2 Robustez" = tabla_robustez,
  "Diagnostico M4" = diagnostico,
  "Influyentes M4" = influyentes,
  "Panel" = panel
), file.path(SALIDA, "resultados_analisis.xlsx"))
cat(sprintf("\nListo. Resultados en '%s/' (resultados_analisis.xlsx y graficos/).\n", SALIDA))

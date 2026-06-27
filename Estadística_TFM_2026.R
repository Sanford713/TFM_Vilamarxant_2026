################################################################################
# Cambios temporales en comunidades hiporreicas del río Turia
# Archivo revisado para: TFM_David(2).xlsx
# Hoja usada: primera hoja del libro Excel, en este archivo llamada ORIGINAL.
#
# Estructura comprobada en el Excel adjunto:
#   A = Code
#   B = Date
#   C = Hour
#   D = Volume
#   E = Probe
#   F:I = parámetros físico-químicos del agua de río
#   J:M = parámetros físico-químicos del agua subterránea
#   N = Mesh size
#   O:W = grupos/taxones contabilizados en este archivo
#          Amphipoda, Cyclopoida, Harpacticoida, Diplostraca, Podocopida,
#          Isopoda, Acariformes, Insects, Bathynellacea
#
# Si en otra versión del Excel existen columnas X:Y, el script las incorporará
# automáticamente como columnas de taxones, porque toma O:min(Y, última columna).
#
# Análisis incluidos:
#   1) Matriz de densidades estandarizada por esfuerzo: individuos/litro.
#   2) Transformación de Hellinger.
#   3) Año como variable principal y DOY como covariable estacional circular:
#      sin_DOY y cos_DOY.
#   4) Mesh size como variable explicativa adicional.
#   5) PERMDISP para Year y Mesh_size.
#   6) PERMANOVA principal con sumas de cuadrados marginales.
#   7) PERMANOVA secundaria con permutaciones restringidas dentro de Year.
#   8) PERMANOVA exploratoria extendida con físico-química, si hay datos completos.
#   9) NMDS 2D y gráficos exploratorios.
#
# Notas de limpieza específicas para este Excel:
#   - Varias celdas decimales fueron autoformateadas por Excel como fechas
#     con formato d.m. Ejemplos: 18.2, 21.5 y 8.11. El script las repara.
#   - Algunas celdas de pH pueden aparecer como enteros por problema de separador
#     decimal, por ejemplo 8.115 como 8115. El script las repara en columnas pH.
#   - Conteos como 50+ se convierten a 50, interpretándolo como mínimo observado.
################################################################################

# ------------------------------------------------------------------------------
# 0. Paquetes
# ------------------------------------------------------------------------------
# Instalar si hace falta:
# install.packages(c(
#   "readxl", "dplyr", "tidyr", "stringr", "lubridate", "readr",
#   "vegan", "permute", "ggplot2", "tibble", "purrr"
# ))

library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(lubridate)
library(readr)
library(vegan)
library(permute)
library(ggplot2)
library(tibble)
library(purrr)

# ------------------------------------------------------------------------------
# 1. Parámetros de entrada/salida
# ------------------------------------------------------------------------------
input_file <- "TFM_David.xlsx"  # Cambiar por la ruta local si es necesario
sheet_id   <- 1                    # Primera hoja del libro Excel
n_perm     <- 999
set.seed(123)

output_dir <- "resultados_Turia_PERMANOVA_NMDS"
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

# ------------------------------------------------------------------------------
# 2. Funciones auxiliares de limpieza
# ------------------------------------------------------------------------------
na_strings <- c("", "NA", "Na", "na", "N/A", "n/a", "NULL", "null", "-", "--")

# Convierte una fecha de Excel que procede de un decimal mal interpretado.
# Ejemplo: si Excel interpretó 8.11 como 8 de noviembre, se recupera 8.11;
# si interpretó 18.2 como 18 de febrero, se recupera 18.2.
date_to_decimal_d_m <- function(dt) {
  m <- month(dt)
  day(dt) + m / if_else(m < 10, 10, 100)
}

# Repara números leídos desde Excel.
# expected:
#   "generic"      = número normal
#   "date_decimal" = números que pueden haber sido convertidos a fecha d.m
#   "pH"           = pH, con reparación de fechas d.m y de enteros tipo 8115
#   "count"        = conteos; 50+ -> 50
#   "mesh"         = tamaño de malla; se deja como número entero/factorizable
parse_excel_numeric <- function(x, expected = "generic") {
  # readxl con col_types = "text" suele devolver caracteres, pero esta función
  # también tolera vectores numéricos, POSIXct o Date.
  if (inherits(x, "Date") || inherits(x, "POSIXt")) {
    out <- date_to_decimal_d_m(as.Date(x))
  } else {
    x_chr <- as.character(x)
    x_chr <- str_squish(x_chr)
    x_chr[x_chr %in% na_strings] <- NA_character_
    x_chr <- str_replace_all(x_chr, ",", ".")

    # Conteos como 50+ se interpretan como 50.
    x_chr <- str_replace(x_chr, "\\+$", "")

    out <- suppressWarnings(parse_number(x_chr, locale = locale(decimal_mark = ".")))

    # Detecta cadenas de fecha. readxl puede devolver las celdas d.m como
    # "2023-05-14" o "2023-05-14 UTC" si fuerza col_types = "text".
    date_token <- str_extract(
      x_chr,
      "^\\d{4}[-/]\\d{1,2}[-/]\\d{1,2}|^\\d{1,2}[-/]\\d{1,2}[-/]\\d{2,4}"
    )
    has_date_token <- !is.na(date_token)

    if (any(has_date_token, na.rm = TRUE)) {
      dt <- suppressWarnings(as.Date(parse_date_time(
        date_token[has_date_token],
        orders = c("ymd", "dmy", "Ymd", "dmY"),
        tz = "UTC"
      )))
      out[has_date_token] <- date_to_decimal_d_m(dt)
    }

    # Detecta seriales de Excel en columnas donde el valor real debería ser
    # pequeño/decimal. Ejemplo: 45067 corresponde a 21/05/2023 -> 21.5.
    if (expected %in% c("date_decimal", "pH")) {
      idx_serial <- !is.na(out) & out > 30000 & out < 60000
      if (any(idx_serial)) {
        dt <- as.Date(out[idx_serial], origin = "1899-12-30")
        out[idx_serial] <- date_to_decimal_d_m(dt)
      }
    }
  }

  # Reparación específica para pH: 8115 debe ser 8.115, no 8115.
  if (expected == "pH") {
    idx_bad_ph <- !is.na(out) & out > 14 & out < 30000
    while (any(idx_bad_ph)) {
      out[idx_bad_ph] <- out[idx_bad_ph] / 10
      idx_bad_ph <- !is.na(out) & out > 14 & out < 30000
    }
  }

  as.numeric(out)
}

parse_excel_date <- function(x) {
  if (inherits(x, "Date") || inherits(x, "POSIXt")) {
    return(as.Date(x))
  }

  x_chr <- as.character(x)
  x_chr <- str_squish(x_chr)
  x_chr[x_chr %in% na_strings] <- NA_character_

  dt <- suppressWarnings(as.Date(parse_date_time(
    x_chr,
    orders = c(
      "ymd", "dmy", "dmY", "Ymd",
      "d/m/Y", "d/m/y", "Y-m-d", "Y/m/d",
      "ymd HMS", "dmy HMS"
    ),
    tz = "UTC"
  )))

  # Si la fecha viene como serial de Excel, por ejemplo 44951.
  idx_serial <- is.na(dt) & str_detect(x_chr, "^\\d+(\\.\\d+)?$")
  if (any(idx_serial, na.rm = TRUE)) {
    serial <- suppressWarnings(as.numeric(x_chr[idx_serial]))
    dt[idx_serial] <- as.Date(serial, origin = "1899-12-30")
  }

  dt
}

parse_excel_time <- function(x) {
  if (inherits(x, "POSIXt")) return(format(x, "%H:%M"))

  x_chr <- as.character(x)
  x_chr <- str_squish(x_chr)
  x_chr[x_chr %in% na_strings] <- NA_character_

  # Si ya tiene formato HH:MM o HH:MM:SS.
  out <- str_extract(x_chr, "\\d{1,2}:\\d{2}(:\\d{2})?")
  out <- str_replace(out, "^(\\d{1,2}:\\d{2}):\\d{2}$", "\\1")

  # Si viene como fracción de día de Excel, por ejemplo 0.4930556.
  idx <- is.na(out) & str_detect(x_chr, "^0?\\.\\d+$")
  if (any(idx, na.rm = TRUE)) {
    frac <- suppressWarnings(as.numeric(x_chr[idx]))
    secs <- round(frac * 24 * 60 * 60)
    hh <- floor(secs / 3600)
    mm <- floor((secs %% 3600) / 60)
    out[idx] <- sprintf("%02d:%02d", hh, mm)
  }

  # Conserva la cadena original si no se ha podido interpretar.
  out[is.na(out) & !is.na(x_chr)] <- x_chr[is.na(out) & !is.na(x_chr)]
  out
}

safe_scale <- function(x) {
  x <- as.numeric(x)
  s <- sd(x, na.rm = TRUE)
  if (all(is.na(x)) || is.na(s) || s == 0) {
    return(rep(NA_real_, length(x)))
  }
  as.numeric(scale(x))
}

write_adonis <- function(x, file) {
  out <- as.data.frame(x) %>% rownames_to_column(var = "term")
  write_csv(out, file.path(output_dir, file))
  invisible(out)
}

run_permdisp <- function(dist_object, group, group_name, prefix) {
  group <- droplevels(factor(group))
  tab <- table(group)
  write_csv(
    tibble(group = names(tab), n = as.integer(tab)),
    file.path(output_dir, paste0(prefix, "_group_sizes.csv"))
  )

  if (nlevels(group) < 2 || any(tab < 2)) {
    msg <- paste0(
      "PERMDISP para ", group_name,
      " no ejecutado: se necesitan al menos dos grupos y al menos dos muestras por grupo."
    )
    warning(msg)
    writeLines(msg, file.path(output_dir, paste0(prefix, "_NO_EJECUTADO.txt")))
    return(invisible(NULL))
  }

  bd <- betadisper(dist_object, group = group, type = "centroid")
  bd_anova <- anova(bd)
  bd_perm <- permutest(bd, permutations = n_perm)

  capture.output(bd_anova, file = file.path(output_dir, paste0(prefix, "_anova.txt")))
  capture.output(bd_perm,  file = file.path(output_dir, paste0(prefix, "_permutation_test.txt")))

  invisible(list(model = bd, anova = bd_anova, permutest = bd_perm))
}

# ------------------------------------------------------------------------------
# 3. Lectura de datos
# ------------------------------------------------------------------------------
# Se lee todo como texto para evitar que readxl imponga tipos incoherentes en
# columnas mezcladas por autoformato de Excel. Después se tipifica explícitamente.
raw <- read_excel(
  input_file,
  sheet = sheet_id,
  col_types = "text",
  .name_repair = "unique"
)

stopifnot(ncol(raw) >= 15)

# Posiciones según el diseño del Excel.
col_code   <- names(raw)[1]
col_date   <- names(raw)[2]
col_hour   <- names(raw)[3]
col_volume <- names(raw)[4]
col_probe  <- names(raw)[5]
col_mesh   <- names(raw)[14]

river_cols  <- names(raw)[6:9]
ground_cols <- names(raw)[10:13]

# Taxones: O:Y, pero en este archivo concreto solo existen O:W.
last_taxon_position <- min(25, ncol(raw))
taxa_positions <- 15:last_taxon_position
taxa_cols <- names(raw)[taxa_positions]

message("Hoja leída: ", sheet_id)
message("Número de filas de datos: ", nrow(raw))
message("Número de columnas: ", ncol(raw))
message("Columnas de taxones usadas: ", paste(taxa_cols, collapse = ", "))

# ------------------------------------------------------------------------------
# 4. Limpieza y construcción de metadatos
# ------------------------------------------------------------------------------
dat <- raw %>%
  transmute(
    Code = .data[[col_code]],
    Date = parse_excel_date(.data[[col_date]]),
    Hour = parse_excel_time(.data[[col_hour]]),
    Volume_L = parse_excel_numeric(.data[[col_volume]], expected = "date_decimal"),
    Probe = .data[[col_probe]],

    River_T_C       = parse_excel_numeric(.data[[river_cols[1]]], expected = "date_decimal"),
    River_Cond_uScm = parse_excel_numeric(.data[[river_cols[2]]], expected = "generic"),
    River_O2_pct    = parse_excel_numeric(.data[[river_cols[3]]], expected = "generic"),
    River_pH        = parse_excel_numeric(.data[[river_cols[4]]], expected = "pH"),

    Ground_T_C       = parse_excel_numeric(.data[[ground_cols[1]]], expected = "date_decimal"),
    Ground_Cond_uScm = parse_excel_numeric(.data[[ground_cols[2]]], expected = "generic"),
    Ground_O2_pct    = parse_excel_numeric(.data[[ground_cols[3]]], expected = "generic"),
    Ground_pH        = parse_excel_numeric(.data[[ground_cols[4]]], expected = "pH"),

    Mesh_size = parse_excel_numeric(.data[[col_mesh]], expected = "mesh")
  ) %>%
  bind_cols(
    raw %>%
      select(all_of(taxa_cols)) %>%
      mutate(across(everything(), ~ parse_excel_numeric(.x, expected = "count"))) %>%
      mutate(across(everything(), ~ replace_na(.x, 0)))
  ) %>%
  mutate(
    Year = factor(year(Date)),
    DOY = yday(Date),
    days_in_year = if_else(leap_year(Date), 366, 365),
    sin_DOY = sin(2 * pi * (DOY - 1) / days_in_year),
    cos_DOY = cos(2 * pi * (DOY - 1) / days_in_year),
    Mesh_size = factor(Mesh_size)
  )

# Guarda una tabla limpia para revisar que las conversiones fueron correctas.
write_csv(dat, file.path(output_dir, "00_datos_limpios_completos.csv"))

# Comprobaciones de rango para detectar errores de importación.
range_checks <- tibble(
  variable = c("Volume_L", "River_T_C", "Ground_T_C", "River_pH", "Ground_pH", "Mesh_size"),
  min = c(
    min(dat$Volume_L, na.rm = TRUE),
    min(dat$River_T_C, na.rm = TRUE),
    min(dat$Ground_T_C, na.rm = TRUE),
    min(dat$River_pH, na.rm = TRUE),
    min(dat$Ground_pH, na.rm = TRUE),
    min(as.numeric(as.character(dat$Mesh_size)), na.rm = TRUE)
  ),
  max = c(
    max(dat$Volume_L, na.rm = TRUE),
    max(dat$River_T_C, na.rm = TRUE),
    max(dat$Ground_T_C, na.rm = TRUE),
    max(dat$River_pH, na.rm = TRUE),
    max(dat$Ground_pH, na.rm = TRUE),
    max(as.numeric(as.character(dat$Mesh_size)), na.rm = TRUE)
  )
)
write_csv(range_checks, file.path(output_dir, "00_revision_rangos_variables.csv"))

# Filtra muestras sin información esencial.
valid_rows <- dat %>%
  mutate(.row_id = row_number()) %>%
  filter(
    !is.na(Code),
    !is.na(Date),
    !is.na(Year),
    !is.na(Volume_L),
    Volume_L > 0,
    !is.na(Mesh_size)
  ) %>%
  pull(.row_id)

if (length(valid_rows) < nrow(dat)) {
  warning("Se excluyen ", nrow(dat) - length(valid_rows),
          " filas por Code, fecha, volumen o mesh size ausente/no válido.")
}

dat <- dat[valid_rows, , drop = FALSE]

# ------------------------------------------------------------------------------
# 5. Matriz de densidades y transformación de Hellinger
# ------------------------------------------------------------------------------
comm_counts <- dat %>%
  select(all_of(taxa_cols)) %>%
  as.data.frame()

rownames(comm_counts) <- dat$Code

# Elimina columnas de taxones que sean completamente cero, si existieran.
non_zero_taxa <- colSums(comm_counts, na.rm = TRUE) > 0
if (any(!non_zero_taxa)) {
  warning("Se excluyen taxones/grupos con suma total 0: ",
          paste(names(comm_counts)[!non_zero_taxa], collapse = ", "))
}
comm_counts <- comm_counts[, non_zero_taxa, drop = FALSE]

# Densidad = individuos/litro muestreado.
comm_density <- sweep(as.matrix(comm_counts), 1, dat$Volume_L, FUN = "/")
rownames(comm_density) <- dat$Code

# Elimina muestras completamente vacías.
non_empty <- rowSums(comm_density, na.rm = TRUE) > 0
if (any(!non_empty)) {
  warning("Se excluyen ", sum(!non_empty),
          " muestras con suma total de densidad = 0.")
}

comm_density <- comm_density[non_empty, , drop = FALSE]
dat <- dat[non_empty, , drop = FALSE]

# Transformación de Hellinger: sqrt(abundancia relativa por muestra).
comm_hel <- decostand(comm_density, method = "hellinger")

write_csv(as.data.frame(comm_counts) %>% rownames_to_column("Code"),
          file.path(output_dir, "01_conteos_limpios.csv"))
write_csv(as.data.frame(comm_density) %>% rownames_to_column("Code"),
          file.path(output_dir, "02_densidades_individuos_por_litro.csv"))
write_csv(as.data.frame(comm_hel) %>% rownames_to_column("Code"),
          file.path(output_dir, "03_matriz_hellinger.csv"))

meta <- dat %>%
  select(
    Code, Date, Hour, Volume_L, Probe, Mesh_size, Year, DOY, sin_DOY, cos_DOY,
    River_T_C, River_Cond_uScm, River_O2_pct, River_pH,
    Ground_T_C, Ground_Cond_uScm, Ground_O2_pct, Ground_pH
  ) %>%
  mutate(
    Year = droplevels(Year),
    Mesh_size = droplevels(Mesh_size)
  )

write_csv(meta, file.path(output_dir, "00_metadata_limpios.csv"))

# Distancia euclídea sobre Hellinger.
dist_hel <- dist(comm_hel, method = "euclidean")

# ------------------------------------------------------------------------------
# 6. PERMDISP: homogeneidad de dispersiones multivariantes
# ------------------------------------------------------------------------------
permdisp_year <- run_permdisp(
  dist_object = dist_hel,
  group = meta$Year,
  group_name = "Year",
  prefix = "04_PERMDISP_Year"
)

permdisp_mesh <- run_permdisp(
  dist_object = dist_hel,
  group = meta$Mesh_size,
  group_name = "Mesh_size",
  prefix = "05_PERMDISP_Mesh_size"
)

# ------------------------------------------------------------------------------
# 7. PERMANOVA principal: Year + estacionalidad circular + Mesh_size
# ------------------------------------------------------------------------------
# Nota estadística importante:
# Si Year es el predictor principal, no se debe bloquear la permutación dentro de
# Year para probar Year, porque entonces las muestras no pueden permutarse entre
# anualidades. Por tanto, este modelo principal usa permutaciones no restringidas
# y sumas de cuadrados marginales para evaluar la contribución independiente de
# cada predictor.

perm_unrestricted <- how(nperm = n_perm)

adonis_core <- adonis2(
  dist_hel ~ Year + sin_DOY + cos_DOY + Mesh_size,
  data = meta,
  permutations = perm_unrestricted,
  by = "margin"
)

write_adonis(adonis_core, "06_PERMANOVA_principal_Year_DOY_Mesh.csv")
capture.output(adonis_core,
               file = file.path(output_dir, "06_PERMANOVA_principal_Year_DOY_Mesh.txt"))

# ------------------------------------------------------------------------------
# 8. PERMANOVA secundaria con permutaciones restringidas dentro de Year
# ------------------------------------------------------------------------------
# Este análisis evalúa Mesh_size y la señal estacional dentro de cada año. Year se
# omite de la fórmula porque se usa como bloque de permutación.

perm_within_year <- how(nperm = n_perm, blocks = meta$Year)

adonis_within_year <- adonis2(
  dist_hel ~ sin_DOY + cos_DOY + Mesh_size,
  data = meta,
  permutations = perm_within_year,
  by = "margin"
)

write_adonis(adonis_within_year, "07_PERMANOVA_bloqueada_dentro_de_Year.csv")
capture.output(adonis_within_year,
               file = file.path(output_dir, "07_PERMANOVA_bloqueada_dentro_de_Year.txt"))

# ------------------------------------------------------------------------------
# 9. PERMANOVA exploratoria extendida con físico-química del agua
# ------------------------------------------------------------------------------
# Se ejecuta solo con muestras completas para las variables ambientales incluidas.
# Esta parte es exploratoria: la disponibilidad de O2 y otros parámetros puede
# reducir bastante el tamaño muestral.

env_vars_all <- c(
  "River_T_C", "River_Cond_uScm", "River_O2_pct", "River_pH",
  "Ground_T_C", "Ground_Cond_uScm", "Ground_O2_pct", "Ground_pH"
)

# Conserva variables ambientales con al menos 10 datos y varianza no nula.
env_vars <- env_vars_all[vapply(meta[env_vars_all], function(z) {
  z <- as.numeric(z)
  sum(!is.na(z)) >= 10 && !is.na(sd(z, na.rm = TRUE)) && sd(z, na.rm = TRUE) > 0
}, logical(1))]

meta_env <- meta %>%
  mutate(across(all_of(env_vars), safe_scale))

rhs_terms_env <- c("Year", "sin_DOY", "cos_DOY", "Mesh_size", env_vars)
complete_env <- complete.cases(meta_env[, rhs_terms_env])

min_n_env <- max(15, length(rhs_terms_env) + 5)

if (length(env_vars) > 0 &&
    sum(complete_env) >= min_n_env &&
    nlevels(droplevels(meta_env$Year[complete_env])) >= 2 &&
    nlevels(droplevels(meta_env$Mesh_size[complete_env])) >= 2) {

  dist_env <- dist(comm_hel[complete_env, , drop = FALSE], method = "euclidean")
  meta_env_cc <- meta_env[complete_env, , drop = FALSE] %>%
    mutate(
      Year = droplevels(Year),
      Mesh_size = droplevels(Mesh_size)
    )

  form_env <- as.formula(paste("dist_env ~", paste(rhs_terms_env, collapse = " + ")))

  adonis_env <- adonis2(
    form_env,
    data = meta_env_cc,
    permutations = how(nperm = n_perm),
    by = "margin"
  )

  write_adonis(adonis_env, "08_PERMANOVA_extendida_fisicoquimica.csv")
  capture.output(adonis_env,
                 file = file.path(output_dir, "08_PERMANOVA_extendida_fisicoquimica.txt"))

  write_csv(
    tibble(
      variable = env_vars,
      n_non_missing = vapply(meta[env_vars], function(z) sum(!is.na(z)), integer(1))
    ),
    file.path(output_dir, "08_variables_fisicoquimicas_incluidas.csv")
  )
} else {
  msg <- paste0(
    "No se ejecuta la PERMANOVA extendida con físico-química. ",
    "Variables ambientales útiles: ", length(env_vars), "; ",
    "muestras completas: ", sum(complete_env), "; ",
    "mínimo requerido por el script: ", min_n_env, "."
  )
  message(msg)
  writeLines(msg, con = file.path(output_dir, "08_PERMANOVA_extendida_fisicoquimica_NO_EJECUTADA.txt"))
}

# ------------------------------------------------------------------------------
# 10. NMDS exploratorio en 2D
# ------------------------------------------------------------------------------
# Se usa matriz Hellinger y distancia euclídea. autotransform = FALSE porque la
# transformación ya se ha aplicado.

nmds <- metaMDS(
  comm_hel,
  distance = "euclidean",
  k = 2,
  trymax = 200,
  autotransform = FALSE,
  trace = FALSE
)

nmds_scores <- as.data.frame(scores(nmds, display = "sites")) %>%
  rownames_to_column("Code") %>%
  left_join(meta, by = "Code")

write_csv(nmds_scores, file.path(output_dir, "09_NMDS_scores.csv"))
writeLines(paste("NMDS stress =", round(nmds$stress, 4)),
           con = file.path(output_dir, "09_NMDS_stress.txt"))

year_centroids <- nmds_scores %>%
  group_by(Year) %>%
  summarise(
    NMDS1_centroid = mean(NMDS1, na.rm = TRUE),
    NMDS2_centroid = mean(NMDS2, na.rm = TRUE),
    n = n(),
    .groups = "drop"
  )

p_nmds <- ggplot(nmds_scores, aes(x = NMDS1, y = NMDS2)) +
  geom_point(aes(colour = Year, shape = Mesh_size), size = 3, alpha = 0.85) +
  geom_point(
    data = year_centroids,
    aes(x = NMDS1_centroid, y = NMDS2_centroid, colour = Year),
    inherit.aes = FALSE,
    size = 5,
    shape = 4,
    stroke = 1.2
  ) +
  labs(
    title = "NMDS de comunidades hiporreicas del río Turia",
    subtitle = paste0("Hellinger + distancia euclídea; stress = ", round(nmds$stress, 4)),
    x = "NMDS1",
    y = "NMDS2",
    colour = "Año",
    shape = "Mesh size"
  ) +
  theme_bw(base_size = 12)

ggsave(file.path(output_dir, "10_NMDS_Year_Mesh.png"), p_nmds,
       width = 8, height = 6, dpi = 300)
ggsave(file.path(output_dir, "10_NMDS_Year_Mesh.pdf"), p_nmds,
       width = 8, height = 6)

p_nmds_doy <- ggplot(nmds_scores, aes(x = NMDS1, y = NMDS2)) +
  geom_point(aes(colour = DOY, shape = Mesh_size), size = 3, alpha = 0.85) +
  scale_colour_viridis_c(option = "D") +
  labs(
    title = "NMDS con gradiente estacional",
    subtitle = paste0("DOY = día del año; stress = ", round(nmds$stress, 4)),
    x = "NMDS1",
    y = "NMDS2",
    colour = "DOY",
    shape = "Mesh size"
  ) +
  theme_bw(base_size = 12)

ggsave(file.path(output_dir, "11_NMDS_DOY_Mesh.png"), p_nmds_doy,
       width = 8, height = 6, dpi = 300)
ggsave(file.path(output_dir, "11_NMDS_DOY_Mesh.pdf"), p_nmds_doy,
       width = 8, height = 6)

png(file.path(output_dir, "12_NMDS_stressplot.png"), width = 1800, height = 1400, res = 250)
stressplot(nmds, main = paste0("NMDS stressplot; stress = ", round(nmds$stress, 4)))
dev.off()

# ------------------------------------------------------------------------------
# 11. Resumen final en consola
# ------------------------------------------------------------------------------
cat("\nAnálisis completado.\n")
cat("Muestras analizadas:", nrow(meta), "\n")
cat("Taxones/grupos analizados:", ncol(comm_hel), "\n")
cat("Años:", paste(levels(meta$Year), collapse = ", "), "\n")
cat("Mesh sizes:", paste(levels(meta$Mesh_size), collapse = ", "), "\n")
cat("NMDS stress:", round(nmds$stress, 4), "\n")
cat("Resultados guardados en:", normalizePath(output_dir), "\n")

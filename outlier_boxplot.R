# =============================================================================
# outlier_boxplot.R
# Payments Statistics — Outlier Detection Visualisation
#
# Chart: Boxplot of total payment value (EUR bn) by instrument
#        Germany (DE) | Domestic transactions | 2019-Q1 to 2024-Q4
#        Outliers highlighted as red points above/below whiskers
# =============================================================================

library(readxl)
library(ggplot2)
library(dplyr)
library(scales)

# --- Paths ---
COUNTRY     <- "DE"
INPUT_FILE  <- "/mnt/c/Users/qendresa.krasniqi/Downloads/payments_statistics_datasets.xlsx"
OUTPUT_FILE <- "/mnt/c/Users/qendresa.krasniqi/Downloads/final_summary/germany_outlier_boxplot.jpg"

# =============================================================================
# STEP 1 — LOAD
# Read directly from raw source file — same source of truth as pipeline.py
# =============================================================================

cat("\n[LOAD] Reading raw payments data...\n")
df <- read_excel(INPUT_FILE, sheet = "payments_transactions_quarterly")
cat(sprintf("  Loaded: %d rows x %d columns\n", nrow(df), ncol(df)))

# =============================================================================
# STEP 2 — PREPARE
# Filter to Germany, remove nulls
# Aggregate by instrument and quarter — one value per quarter per instrument
# Convert to EUR bn for readability
# =============================================================================

cat(sprintf("[PREPARE] Filtering to %s...\n", COUNTRY))

df_prep <- df %>%
  filter(
    !is.na(total_value_eur_mn),
    reporting_country == COUNTRY,
    transaction_type  == "domestic"
  ) %>%
  group_by(payment_instrument, quarter) %>%
  summarise(
    total_value_eur_bn = round(sum(total_value_eur_mn, na.rm = TRUE) / 1000, 2),
    .groups = "drop"
  ) %>%
  mutate(
    instrument_label = recode(payment_instrument,
      "credit_transfer" = "Credit Transfer",
      "card_payment"    = "Card Payment",
      "direct_debit"    = "Direct Debit",
      "e-money"         = "E-Money",
      "cheque"          = "Cheque"
    ),
    # Reorder instruments by median value descending
    instrument_label = reorder(instrument_label, total_value_eur_bn, FUN = median)
  )

cat(sprintf("  Rows after filter : %d\n", nrow(df_prep)))

# =============================================================================
# STEP 3 — DETECT OUTLIERS
# Use IQR method — same logic as pipeline.py Step 4
# Points beyond 1.5 x IQR from Q1/Q3 are flagged as outliers
# =============================================================================

cat("[OUTLIERS] Flagging outliers using IQR method...\n")

df_prep <- df_prep %>%
  group_by(instrument_label) %>%
  mutate(
    Q1      = quantile(total_value_eur_bn, 0.25),
    Q3      = quantile(total_value_eur_bn, 0.75),
    IQR     = Q3 - Q1,
    is_outlier = total_value_eur_bn < (Q1 - 1.5 * IQR) |
                 total_value_eur_bn > (Q3 + 1.5 * IQR)
  ) %>%
  ungroup()

n_outliers <- sum(df_prep$is_outlier)
cat(sprintf("  Outliers flagged  : %d rows\n", n_outliers))

df_outliers <- df_prep %>% filter(is_outlier)

# =============================================================================
# STEP 4 — PLOT
# Boxplot per instrument — outliers shown as red points
# Clean minimal style consistent with time series chart
# =============================================================================

cat("[PLOT] Building boxplot...\n")

instrument_colours <- c(
  "Credit Transfer" = "#003366",
  "Card Payment"    = "#FF6600",
  "Direct Debit"    = "#009966",
  "E-Money"         = "#CC0000",
  "Cheque"          = "#6600CC"
)

p <- ggplot(df_prep, aes(
    x    = instrument_label,
    y    = total_value_eur_bn,
    fill = instrument_label
  )) +

  # Boxplot — hide default outlier points (we plot our own)
  geom_boxplot(
    outlier.shape = NA,
    alpha         = 0.6,
    width         = 0.5,
    colour        = "#333333",
    linewidth     = 0.5
  ) +

  # All data points as jittered dots for transparency
  geom_jitter(
    width   = 0.15,
    size    = 1.8,
    alpha   = 0.5,
    colour  = "#555555"
  ) +

  # Outliers highlighted in red
  geom_point(
    data   = df_outliers,
    colour = "#CC0000",
    size   = 3,
    shape  = 18        # diamond shape
  ) +

  # Label outlier count if any exist
  {if (n_outliers > 0)
    annotate("text",
      x     = 0.6,
      y     = max(df_prep$total_value_eur_bn) * 0.98,
      label = sprintf("Red points = outliers (IQR method)\nn = %d flagged", n_outliers),
      hjust = 0,
      size  = 3,
      colour = "#CC0000"
    )
  } +

  scale_y_continuous(
    labels = label_comma(suffix = " bn"),
    expand = expansion(mult = c(0.05, 0.1))
  ) +
  scale_fill_manual(values = instrument_colours) +

  labs(
    title    = "Germany — Payment Value Distribution by Instrument",
    subtitle = "Domestic transactions  |  Quarterly observations, 2019-Q1 to 2024-Q4  |  EUR bn",
    x        = NULL,
    y        = "Total Value (EUR bn)",
    caption  = "Source: Synthetic data — Euro Area payments statistics framework  |  Country: DE\nOutliers identified using IQR method (beyond 1.5 × IQR from Q1/Q3)"
  ) +

  theme_minimal(base_size = 11) +
  theme(
    plot.title       = element_text(face = "bold", colour = "#003366", size = 13),
    plot.subtitle    = element_text(colour = "#555555", size = 9.5),
    plot.caption     = element_text(colour = "#999999", size = 8, hjust = 0),
    plot.margin      = margin(10, 20, 10, 10),
    axis.text.x      = element_text(size = 10, face = "bold"),
    axis.text.y      = element_text(size = 9),
    legend.position  = "none",
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(colour = "#EEEEEE", linewidth = 0.4),
    plot.background  = element_rect(fill = "white", colour = NA),
    panel.background = element_rect(fill = "white", colour = NA)
  )

# =============================================================================
# STEP 5 — SAVE
# =============================================================================

cat(sprintf("[SAVE] Writing chart to: %s\n", OUTPUT_FILE))
dir.create(dirname(OUTPUT_FILE), showWarnings = FALSE, recursive = TRUE)

ggsave(
  filename = OUTPUT_FILE,
  plot     = p,
  width    = 11,
  height   = 7,
  dpi      = 150,
  device   = "jpeg",
  quality  = 95,
  bg       = "white"
)

cat("[DONE] Chart saved successfully.\n\n")
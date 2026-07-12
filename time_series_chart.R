# =============================================================================
# charts.R
# Payments Statistics — Quarterly Time Series Visualisation
#
# Chart: Total payment value (EUR bn) by payment instrument
#        Germany (DE) | Domestic transactions | 2019-Q1 to 2024-Q4
# =============================================================================

library(readxl)
library(ggplot2)
library(dplyr)
library(scales)
library(ggrepel)

# --- Paths ---
COUNTRY     <- "DE"
INPUT_FILE  <- "/mnt/c/Users/qendresa.krasniqi/Downloads/payments_statistics_datasets.xlsx"
OUTPUT_FILE <- "/mnt/c/Users/qendresa.krasniqi/Downloads/final_summary/germany_payments_time_series.jpg"

# =============================================================================
# STEP 1 — LOAD
# Read directly from raw source file — same source of truth as pipeline.py
# =============================================================================

cat("\n[LOAD] Reading raw payments data...\n")
df <- read_excel(INPUT_FILE, sheet = "payments_transactions_quarterly")
cat(sprintf("  Loaded: %d rows x %d columns\n", nrow(df), ncol(df)))

# =============================================================================
# STEP 2 — PREPARE
# Filter to Germany domestic transactions only
# Aggregate total value (EUR bn) by instrument and quarter
# Remove nulls before aggregating — mirrors pipeline.py imputation logic
# =============================================================================

cat(sprintf("[PREPARE] Filtering to %s domestic transactions...\n", COUNTRY))

df_agg <- df %>%
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
    # Format quarter label: 2023Q1 -> 2023-Q1
    period = sub("(\\d{4})(Q\\d)", "\\1-\\2", quarter),
    # Order quarters chronologically
    period = factor(period, levels = unique(period[order(quarter)])),
    # Clean instrument labels
    instrument_label = recode(payment_instrument,
      "credit_transfer" = "Credit Transfer",
      "card_payment"    = "Card Payment",
      "direct_debit"    = "Direct Debit",
      "e-money"         = "E-Money",
      "cheque"          = "Cheque"
    )
  )

# Label only last point per instrument for end-of-line labels
df_labels <- df_agg %>%
  group_by(instrument_label) %>%
  filter(quarter == max(quarter)) %>%
  ungroup()

cat(sprintf("  Rows after filter : %d\n", nrow(df_agg)))
cat(sprintf("  Instruments       : %s\n", paste(unique(df_agg$instrument_label), collapse = ", ")))
cat(sprintf("  Period            : %s to %s\n", min(df_agg$quarter), max(df_agg$quarter)))

# =============================================================================
# STEP 3 — PLOT
# Clean minimal line chart — one line per payment instrument
# All 24 quarter labels shown rotated 90 degrees
# End-of-line labels replace legend for cleaner layout
# =============================================================================

cat("[PLOT] Building chart...\n")

instrument_colours <- c(
  "Credit Transfer" = "#003366",
  "Card Payment"    = "#FF6600",
  "Direct Debit"    = "#009966",
  "E-Money"         = "#CC0000",
  "Cheque"          = "#6600CC"
)

quarter_levels <- levels(df_agg$period)

p <- ggplot(df_agg, aes(
    x      = period,
    y      = total_value_eur_bn,
    colour = instrument_label,
    group  = instrument_label
  )) +

  geom_line(linewidth = 0.9) +
  geom_point(size = 1.5, alpha = 0.8) +

  # End-of-line instrument labels — no connector lines
  geom_text_repel(
    data          = df_labels,
    aes(label     = instrument_label),
    nudge_x       = 1.5,
    direction     = "y",
    hjust         = 0,
    size          = 3.2,
    segment.color = NA,
    show.legend   = FALSE
  ) +

  # Show all 24 quarter labels
  scale_x_discrete(
    breaks = quarter_levels,
    labels = quarter_levels
  ) +
  scale_y_continuous(
    labels = label_comma(suffix = " bn"),
    expand = expansion(mult = c(0.05, 0.08))
  ) +
  scale_colour_manual(values = instrument_colours) +

  labs(
    title    = "Germany — Payments Total Value by Instrument",
    subtitle = "Domestic transactions  |  Quarterly series, 2019-Q1 to 2024-Q4  |  EUR bn",
    x        = NULL,
    y        = "Total Value (EUR bn)",
    caption  = "Source: Synthetic data — Euro Area payments statistics framework  |  Country: DE"
  ) +

  theme_minimal(base_size = 11) +
  theme(
    plot.title       = element_text(face = "bold", colour = "#003366", size = 13),
    plot.subtitle    = element_text(colour = "#555555", size = 9.5),
    plot.caption     = element_text(colour = "#999999", size = 8, hjust = 0),
    plot.margin      = margin(10, 90, 10, 10),
    axis.text.x      = element_text(angle = 90, hjust = 1, vjust = 0.5, size = 7.5),
    axis.text.y      = element_text(size = 9),
    legend.position  = "none",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(colour = "#EEEEEE", linewidth = 0.4),
    plot.background  = element_rect(fill = "white", colour = NA),
    panel.background = element_rect(fill = "white", colour = NA)
  )

# =============================================================================
# STEP 4 — SAVE
# =============================================================================

cat(sprintf("[SAVE] Writing chart to: %s\n", OUTPUT_FILE))
dir.create(dirname(OUTPUT_FILE), showWarnings = FALSE, recursive = TRUE)

ggsave(
  filename = OUTPUT_FILE,
  plot     = p,
  width    = 13,
  height   = 6.5,
  dpi      = 150,
  device   = "jpeg",
  quality  = 95,
  bg       = "white"
)

cat("[DONE] Chart saved successfully.\n\n")
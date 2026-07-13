# =============================================================================
# charts.R
# Payments Statistics — Quarterly Time Series Visualisation
#
# Chart: Total payment value (EUR bn) by instrument over time
#        2019-Q1 to 2024-Q4 across all euro area countries
# =============================================================================

library(readxl)
library(ggplot2)
library(dplyr)
library(scales)
library(ggrepel)

# --- Paths ---
INPUT_FILE  <- "C:/Users/qendresa.krasniqi/Downloads/payments_statistics_datasets.xlsx"
OUTPUT_FILE <- "C:/Users/qendresa.krasniqi/Downloads/final_summary/payments_time_series.png"

# =============================================================================
# STEP 1 — LOAD
# Read directly from raw source file — same source of truth as pipeline.py
# =============================================================================

cat("\n[LOAD] Reading raw payments data...\n")
df <- read_excel(INPUT_FILE, sheet = "payments_transactions_quarterly")
cat(sprintf("  Loaded: %d rows x %d columns\n", nrow(df), ncol(df)))

# =============================================================================
# STEP 2 — PREPARE
# Aggregate total value (EUR bn) by instrument and quarter
# Sort quarters chronologically and format labels as 2019-Q1 etc.
# Remove nulls before aggregating (mirrors pipeline.py imputation logic)
# =============================================================================

cat("[PREPARE] Aggregating by instrument and quarter...\n")

df_agg <- df %>%
  filter(!is.na(total_value_eur_mn)) %>%
  group_by(payment_instrument, quarter) %>%
  summarise(
    total_value_eur_bn = round(sum(total_value_eur_mn, na.rm = TRUE) / 1000, 2),
    .groups = "drop"
  ) %>%
  mutate(
    # Format quarter label: 2023Q1 -> 2019-Q1
    period = sub("(\\d{4})(Q\\d)", "\\1-\\2", quarter),
    # Order quarters chronologically
    period = factor(period, levels = unique(period[order(quarter)])),
    # Clean instrument labels for legend
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

cat(sprintf("  Instruments: %s\n",
    paste(unique(df_agg$instrument_label), collapse = ", ")))
cat(sprintf("  Quarters: %s to %s\n",
    min(df_agg$quarter), max(df_agg$quarter)))

# =============================================================================
# STEP 3 — PLOT
# Clean minimal line chart — one line per payment instrument
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

# Show every 4th quarter on x-axis to avoid crowding
quarter_levels <- levels(df_agg$period)
x_breaks <- quarter_levels[seq(1, length(quarter_levels), by = 4)]

p <- ggplot(df_agg, aes(
    x      = period,
    y      = total_value_eur_bn,
    colour = instrument_label,
    group  = instrument_label
  )) +

  geom_line(linewidth = 0.9) +
  geom_point(size = 1.5, alpha = 0.8) +

  # End-of-line labels
  geom_text_repel(
    data          = df_labels,
    aes(label     = instrument_label),
    nudge_x       = 1.5,
    direction     = "y",
    hjust         = 0,
    size          = 3.2,
    segment.size  = 0.3,
    segment.color = "grey60",
    show.legend   = FALSE
  ) +

  scale_x_discrete(breaks = x_breaks) +
  scale_y_continuous(
    labels = label_comma(suffix = " bn"),
    expand = expansion(mult = c(0.02, 0.05))
  ) +
  scale_colour_manual(values = instrument_colours) +

  labs(
    title    = "Euro Area Payments — Total Value by Instrument",
    subtitle = "Quarterly series, 2019-Q1 to 2024-Q4  |  EUR bn  |  All euro area countries",
    x        = NULL,
    y        = "Total Value (EUR bn)",
    caption  = "Source: Synthetic data — Euro Area payments statistics framework"
  ) +

  theme_minimal(base_size = 11) +
  theme(
    plot.title       = element_text(face = "bold", colour = "#003366", size = 13),
    plot.subtitle    = element_text(colour = "#555555", size = 9.5),
    plot.caption     = element_text(colour = "#999999", size = 8, hjust = 0),
    plot.margin      = margin(10, 80, 10, 10),
    axis.text.x      = element_text(angle = 45, hjust = 1, size = 8.5),
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

# Create output directory if it doesn't exist
dir.create(dirname(OUTPUT_FILE), showWarnings = FALSE, recursive = TRUE)

ggsave(
  filename = OUTPUT_FILE,
  plot     = p,
  width    = 13,
  height   = 6.5,
  dpi      = 150,
  bg       = "white"
)

cat("[DONE] Chart saved successfully.\n\n")
# Updated: Tue Jul 14 00:20:27 CEST 2026

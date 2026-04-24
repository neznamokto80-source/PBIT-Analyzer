import logging
import os


def setup_logging(level: str | None = None) -> None:
    log_level = (level or os.getenv("PBI_ANALYZER_LOG_LEVEL", "INFO")).upper()
    numeric_level = getattr(logging, log_level, logging.INFO)

    logging.basicConfig(
        level=numeric_level,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )

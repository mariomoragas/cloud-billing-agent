from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class ParserResult:
    dataframe: pd.DataFrame
    data_quality: pd.DataFrame = field(default_factory=pd.DataFrame)


@dataclass
class ProcessResult:
    output_path: Path
    presentation_path: Path | None
    raw_df: pd.DataFrame
    service_summary_df: pd.DataFrame
    region_summary_df: pd.DataFrame
    oci_mapping_df: pd.DataFrame
    data_quality_df: pd.DataFrame = field(default_factory=pd.DataFrame)

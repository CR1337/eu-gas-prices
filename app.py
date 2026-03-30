import streamlit as st
import os
import json
import pandas as pd
import datetime as dt
from xlsx_downloader import XlsxDownloader
from data_extractor import RecentDataExtractor, AllDataExtractor
from typing import Optional, Tuple


OUTPUT_FILE_SUFFIX: str = "_by_price_super_prices_europe_"
OUTPUT_FILE_TYPE_SUFFIX: str = ".csv"
CSV_SEP: str = ";"
DECIMAL_SEP: str = ","


def generate_filename(all_: bool, since: Optional[dt.datetime] = None) -> str:
    now = dt.datetime.now()
    prefix = f"{now.day:02d}-{now.month:02d}-{now.year}_{now.hour:02d}{now.minute:02d}{now.second:02d}"
    filename = os.path.join(
        f"{prefix}{OUTPUT_FILE_SUFFIX}"
        + (
            "recent"
            if not all_
            else (
                "all"
                if not since
                else f"since_{since.strftime('%Y-%m-%d')}"  # type: ignore
            )
        )
        + OUTPUT_FILE_TYPE_SUFFIX,
    )
    return filename


def prepare_data(
    since: dt.datetime, last_df: pd.DataFrame
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    recent_workbook = XlsxDownloader().download(False)
    neighbors_workbook = XlsxDownloader().download(True)

    recent_df = RecentDataExtractor().extract(recent_workbook, last_df)
    neighbors_df = AllDataExtractor().extract(neighbors_workbook, since)

    with open("de_neighbor_countries.json", "r", encoding="utf-8") as f:
        neighbor_names = json.load(f)
    keep_colunmns = ["Tag"] + neighbor_names
    neighbors_df = neighbors_df.loc[:, keep_colunmns]
    assert isinstance(neighbors_df, pd.DataFrame)

    return recent_df, neighbors_df


def render_app():
    if not st.session_state.get("state"):
        st.session_state["state"] = "upload"
        st.session_state["recent_df"] = None
        st.session_state["neighbors_df"] = None
        st.session_state["since"] = dt.datetime.now() - dt.timedelta(356 * 10)

    match st.session_state["state"]:
        case "upload":
            old_file = st.file_uploader(
                label="Datei von letzter Woche hier hochladen",
                type="csv",
                max_upload_size=1,
            )
            if old_file is not None:
                with st.spinner("Bitte warten...", show_time=True):
                    try:
                        last_df = pd.read_csv(
                            old_file, sep=CSV_SEP, decimal=DECIMAL_SEP
                        )
                        recent_df, neighbors_df = prepare_data(
                            st.session_state["since"], last_df
                        )
                    except Exception:
                        st.error(
                            "Fehler beim Lesen der hochgeladenen Datei!\nBitte valide Datei hochladen.",
                            icon="🚨",
                        )
                        raise
                    else:
                        st.session_state["recent_df"] = recent_df
                        st.session_state["neighbors_df"] = neighbors_df
                        st.session_state["state"] = "download"
                        st.rerun()

        case "download":
            recent_df: pd.DataFrame = st.session_state["recent_df"]
            neighbors_df: pd.DataFrame = st.session_state["neighbors_df"]

            recent_csv = recent_df.to_csv(
                index=False, sep=CSV_SEP, decimal=DECIMAL_SEP
            ).encode("utf-8")
            neighbors_csv = neighbors_df.to_csv(
                index=False, sep=CSV_SEP, decimal=DECIMAL_SEP
            ).encode("utf-8")

            recent_filename = generate_filename(False, st.session_state["since"])
            neighbors_filename = generate_filename(True, st.session_state["since"])

            st.download_button(
                label="Daten der letzen Woche herunterladen",
                data=recent_csv,
                file_name=recent_filename,
                mime="text/csv",
                icon=":material/download:",
            )

            st.download_button(
                label="Daten der Nachbarstaaten der letzen 10 Jahre herunterladen",
                data=neighbors_csv,
                file_name=neighbors_filename,
                mime="text/csv",
                icon=":material/download:",
            )


if __name__ == "__main__":
    render_app()

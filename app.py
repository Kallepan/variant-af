import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

import os

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm


INPUT_PATH = "input"
OUTPUT_PATH = "results"
REPORT_FILE = "altaf_report.docx"
REPORT_TEMPLATE = "template.docx"
ALTAF_CUTOFF = 0.9
ALTAF_KEY = "AltAF"
ROUND_VALUE = 4

def read_data(file_path):
    df = pd.read_csv(file_path)
    
    n_before_filter = df[ALTAF_KEY].count()
    df[ALTAF_KEY] = pd.to_numeric(df[ALTAF_KEY], errors="coerce")
    df = df.loc[df[ALTAF_KEY] < ALTAF_CUTOFF]

    n_after_filter = df[ALTAF_KEY].count()

    return df, n_before_filter - n_after_filter

def calc_stats(df: pd.DataFrame, n_filtered: int) -> dict:
    count = df[ALTAF_KEY].count()
    min = df[ALTAF_KEY].min()
    max = df[ALTAF_KEY].max()

    mean = df[ALTAF_KEY].mean()
    median = df[ALTAF_KEY].median()

    std = df[ALTAF_KEY].std()

    confidence_lower = round(mean - 1.96 * std/np.sqrt(count), ROUND_VALUE)
    confidence_upper = round(mean + 1.96 * std/np.sqrt(count), ROUND_VALUE)

    return {
        "min": round(min, ROUND_VALUE),
        "max": round(max, ROUND_VALUE),
        "mean": round(mean, ROUND_VALUE),
        "median": round(median, ROUND_VALUE),
        "count": round(count, ROUND_VALUE),
        "std": round(std, ROUND_VALUE),
        "confidence_interval_95%": f"{confidence_lower} - {confidence_upper}",
        f"entfernte_varianten: > {ALTAF_CUTOFF}": n_filtered,
    }

def create_histogram(df: pd.DataFrame, file_name: str):

    plt.hist(df[ALTAF_KEY], bins=100)
    plt.title(f"{file_name[:-4].replace('_','>')} - ALTAF Histogram")
    plt.xlabel(ALTAF_KEY)
    plt.ylabel(ALTAF_KEY)
    plt.savefig(os.path.join(OUTPUT_PATH, file_name))
    plt.close()

def main():
    if not os.path.exists(OUTPUT_PATH):
        os.makedir(OUTPUT_PATH)

    if not os.path.exists(INPUT_PATH):
        raise Exception("Input path does not exist")
    
    # Read the data
    input_files = [file for file in os.listdir(INPUT_PATH) if file.endswith(".csv")]

    variants_data = []
    template = DocxTemplate(REPORT_TEMPLATE)
    for file in input_files:

        df, n_filtered = read_data(os.path.join(INPUT_PATH, file))
        stats = calc_stats(df, n_filtered)
        image_file_name = file.replace(".csv", ".png")
        create_histogram(df, image_file_name)

        variants_data.append({
            "name": file[:-4].replace("_", ">"),
            "stats": stats,
            "figure": InlineImage(template, os.path.join(OUTPUT_PATH, image_file_name), width=Cm(16))
        })
    
    template.render({"variants": variants_data})
    template.save(os.path.join(OUTPUT_PATH, REPORT_FILE))

if __name__ == "__main__":
    main()
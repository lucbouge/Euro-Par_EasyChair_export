import pandas as pd
import os, os.path
import re
from typing import NamedTuple
import matplotlib.pyplot as plt
import numpy as np


dirname = "/Users/bouge/share/Recherche/EuroPar/Conferences/2022/Selection"
filename = "data_2022-04-26.xlsx"

################################################################################


def main():
    df_dict = make_df_dict()
    # export_reviews(df_dict)
    export_reviews_length(df_dict)


################################################################################


class Review(NamedTuple):
    submission_number: str
    review_number: str
    version_number: str


stats_format = """
Number of reviews: {size} 
Min length: {min} 
Max length: {max}  
Mean: {mean:.2f}
Std: {std:.2f}
number with length < mean - 0.5*std: {left_tail}
"""


def export_reviews_length(df_dict):
    review_to_length_dict = dict()
    df = df_dict["Reviews"]
    for (index, row) in df.iterrows():
        submission_number = row["submission #"]
        review_number = row["number"]
        version_number = row["version"]
        text = row["text"].strip()
        text = re.sub(r"\W+", "", text)  # Keep word characters only

        key = Review(
            submission_number=submission_number,
            review_number=review_number,
            version_number=version_number,
        )
        assert key not in review_to_length_dict
        review_to_length_dict[key] = len(text)
    ##
    values = tuple(review_to_length_dict.values())
    size = len(values)
    min = np.min(values)
    max = np.max(values)
    mean = np.mean(values)
    std = np.std(values)
    left_tail = len(tuple(x for x in values if x <= mean - 0.5 * std))
    print(stats_format.format_map(locals()))
    plt.hist(values, bins=100)
    plt.show()


################################################################################


def export_reviews(df_dict):
    reviews_groups = make_reviews_groups(df_dict)
    comments_groups = make_comments_groups(df_dict)
    ##
    reviews_dirname = os.path.join(dirname, "Reviews")
    os.makedirs(reviews_dirname, exist_ok=True)
    for (submission, reviews_df) in reviews_groups:
        reviews_filename = f"submission_{int(submission):03d}.txt"
        reviews_path = os.path.join(reviews_dirname, reviews_filename)
        print(reviews_path)
        comments_df = None
        if submission in comments_groups.groups:
            comments_df = comments_groups.get_group(submission)
        reviews = make_reviews(reviews_df=reviews_df, comments_df=comments_df)
        with open(reviews_path, "w") as cout:
            print(reviews, file=cout)


################################################################################

reviews_format = """
=====================================================================
Submission number: {submission_number}
Reviewer number: {number}
Version number: {version}

{scores}
Total score: {total_score}
{text}
"""

comments_format = """
=====================================================================
Comment on submission number: {submission_number}
Date: {date} {time}

{comment}
"""


def make_reviews(*, reviews_df=None, comments_df=None):
    assert reviews_df is not None
    # assert comments_df is not None
    ##
    reviews = ""
    for (index, row) in reviews_df.iterrows():
        submission_number = row["submission #"]
        number = row["number"]
        version = row["version"]
        text = row["text"]
        text = re.sub(
            r"^\(OVERALL EVALUATION\)\s+", "\n==> General evaluation\n\n", text
        )
        text = re.sub(
            r"\(CONFIDENTIAL REMARKS FOR THE PROGRAM COMMITTEE\)\s+",
            "\n==> Confidential remarks\n\n",
            text,
        )
        total_score = row["total score"]
        scores = row["scores"]
        date = row["date"]
        time = row["time"]
        assert pd.isna(row["attachment?"])
        reviews += reviews_format.format_map(locals())
    if comments_df is not None:
        for (index, row) in comments_df.iterrows():
            comment = row["comment"]
            reviews += comments_format.format_map(locals())
    return reviews


################################################################################


def make_hidden_columns(df):
    hidden_columns = (
        column for column in df.columns if "member" in column or "reviewer" in column
    )
    return list(hidden_columns)


def make_reviews_groups(df_dict):
    df = df_dict["Reviews"]
    hidden_columns = make_hidden_columns(df)
    df = df.drop(columns=hidden_columns)
    print(df.info())
    groups = df.groupby(by="submission #", axis="index")
    return groups


def make_comments_groups(df_dict):
    df = df_dict["Comments"]
    hidden_columns = make_hidden_columns(df)
    df = df.drop(columns=hidden_columns)
    print(df.info())
    ##
    groups = df.groupby(by="submission #", axis="index")
    return groups


################################################################################


def make_df_dict():
    path = os.path.join(dirname, filename)
    df_dict = pd.read_excel(path, engine="openpyxl", dtype="string", sheet_name=None)
    ##
    for (sheetname, df) in df_dict.items():
        print("=" * 30, sheetname)
        for column in df.columns:
            df[column] = df[column].str.strip()
    return df_dict


################################################################################

main()

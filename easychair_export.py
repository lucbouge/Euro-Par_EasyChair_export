import pandas as pd
import os, os.path
import re


dirname = "/Users/bouge/share/Recherche/EuroPar/Conferences/2022/Selection"
filename = "data_2022-04-26.xlsx"

################################################################################


def main():
    path = os.path.join(dirname, filename)
    df_dict = pd.read_excel(path, engine="openpyxl", dtype="string", sheet_name=None)
    ##
    for (sheetname, df) in df_dict.items():
        print("=" * 30, sheetname)
        for column in df.columns:
            df[column] = df[column].str.strip()
        # print(df.info())
    ##
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
Date: {date} {time}

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
    ##
    reviews = list()
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
        review_text = reviews_format.format_map(locals())
        reviews.append(review_text)
    ##
    if comments_df is not None:
        for (index, row) in comments_df.iterrows():
            comment = row["comment"]
            comment_text = comments_format.format_map(locals())
            reviews.append(comment_text)
    ##
    return "\n".join(reviews)


################################################################################


def make_hidden_columns(df):
    hidden_columns = (
        column for column in df.columns if "member" in column or "reviewer" in column
    )
    return hidden_columns


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
    df = df.drop(columns=list(hidden_columns))
    print(df.info())
    ##
    groups = df.groupby(by="submission #", axis="index")
    return groups


################################################################################

main()

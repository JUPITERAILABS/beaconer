import pandas as pd
# from fuzzywuzzy import fuzz
# from fuzzywuzzy import process
from rapidfuzz import fuzz
from rapidfuzz import process
import os
import pytesseract
from pdf2image import convert_from_path
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS, cross_origin

from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

from datetime import datetime

import json
import re
import PyPDF2


app = Flask(__name__)
app.config['CORS_HEADERS'] = 'Content-Type'

@app.route("/compare_questionnaires", methods=["GET", "POST"])
@cross_origin()
def compare_questionnaires():
    print("in this function")
    data = request.files['SIGlight']
    data.save(os.path.join('./temp/', data.filename))
    file_path = os.path.join('./temp/', data.filename)
    # fileobj = data.read()

    # Set up base template
    base_file = "/home/ps/Desktop/Questionnaires/Questionnaires/Beaconer SIGLight Report -AN.xlsx"
    base_questions = pd.read_excel(base_file)
    base_questions = base_questions['Unnamed: 4'][2:]
    base_questions = base_questions.apply(lambda row: re.sub("^\d+\.\d+", "", row, count=1))

    # Set up file to compare
    # new_file = file_path
    try:
        new_questions = pd.read_excel(file_path)
    except:
        return "Error in reading Excel file"
    file = pd.ExcelFile(file_path)
    sheets = file.sheet_names

    if len(sheets) > 1:
        if "Assessment details" in sheets:
            sheets.remove("Assessment details")
            df = pd.concat(pd.read_excel(file_path, sheet_name=sheets))
            new_questions = pd.DataFrame({"Question": df['Question'].to_list(), "Answer": df['Correct answers'].to_list()})
            new_questions['Question'] = new_questions['Question'].apply(lambda row: re.sub("^\d+\.\d+|^[A-Z]{2}\.\d+\.", "", row, count=1))

        else:
            for i in ['Glossary', 'Copyright', 'Terms of Use', 'Cover Page', 'Business Information', 'Documentation']:
                sheets.remove(i)
            df = pd.concat(pd.read_excel(file_path, sheet_name=sheets, header=None))
            # new code
            start_row = 0
            start_column = 0

            # new_questions = pd.read_excel(file_path, header=None)

            for index, row in df.iterrows():
                if 'Question' in str(row):
                    start_row = index
                    start_column = row[row.astype(str).str.contains('Question', case=False, na=False)].index.tolist()[0]
                    # print(start_row)
                    # print(start_column)
                    break
                elif 'Questionnaire'in str(row):
                    continue
            questions = df.iloc[start_row+1:, start_column]
            questions = questions.str.replace(r"^\d+\.\d+|^[A-Z]{2}\.\d+\.", '', regex=True)
            print(questions)

            for index, row in df.iterrows():
                if 'Response' in str(row):
                    start_row = index
                    start_column = row[row.astype(str).str.contains('Response', case=False, na=False)].index.tolist()[0]
                    break

            answer = df.iloc[start_row+1:, start_column]

            # new_questions = new_questions[['Unnamed: 0', 'Unnamed: 2']][7:]
            new_questions = pd.DataFrame({"Question": questions, "Answer": answer})

            # old code
            # new_questions = pd.DataFrame({"Question": df['Question/Request'].to_list(), "Answer": df['Response'].to_list()})
            # new_questions['Question'] = new_questions['Question'].apply(lambda row: re.sub("^\d+\.\d+|^[A-Z]{2}\.\d+\.", "", row, count=1))


        # print(sheets)
        # df = pd.concat(pd.read_excel(file_path, sheet_name=sheets))

        # new_questions = pd.DataFrame({"Question": df['Question'].to_list(), "Answer": df['Actual answers'].to_list()})
        # new_questions['Question'] = new_questions['Question'].apply(lambda row: re.sub("^\d+\.\d+|^[A-Z]{2}\.\d+\.", "", row, count=1))
        # print(new_questions)

    else:
        start_row = 0
        start_column = 0

        new_questions = pd.read_excel(file_path, header=None)

        for index, row in new_questions.iterrows():
            if 'Question' in str(row):
                start_row = index
                start_column = row[row.astype(str).str.contains('Question', case=False, na=False)].index.tolist()[0]
                # print(start_row)
                # print(start_column)
                break
            elif 'Questionnaire'in str(row):
                continue
        questions = new_questions.iloc[start_row+1:, start_column]
        questions = questions.str.replace(r"^\d+\.\d+|^[A-Z]{2}\.\d+\.", '', regex=True)
        print(questions)

        for index, row in new_questions.iterrows():
            if 'Response' in str(row):
                start_row = index
                start_column = row[row.astype(str).str.contains('Response', case=False, na=False)].index.tolist()[0]
                break

        answer = new_questions.iloc[start_row+1:, start_column]

        # new_questions = new_questions[['Unnamed: 0', 'Unnamed: 2']][7:]
        new_questions = pd.DataFrame({"Question": questions, "Answer": answer})
        new_questions['Question'] = new_questions['Question'].apply(lambda row: re.sub("^\d+\.\d+", "", row, count=1))
        print(new_questions)


    dict = {'Question': [],
            'Answer': []}

    for row in base_questions:
        temp = process.extractOne(row, new_questions['Question'])
        if temp[1] > 85:
            dict['Question'].append(row)
            if type(new_questions['Answer'][temp[2]]) is not float:
                dict['Answer'].append(new_questions["Answer"][temp[2]])
            else:
                dict['Answer'].append("Not Answered")
        else:
            dict['Question'].append(row)
            dict['Answer'].append("No")
    print("function end")
    df = pd.DataFrame(dict)
    df.to_excel("output/test2.xlsx")
    return json.dumps(dict, indent=4)

def convert_pdf_to_text(pdf_file):
    # pdf_file = open("soc_2/Venus-SOC2-Type2 AN.pdf", "rb")
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    num_pages = len(pdf_reader.pages)
    txt_file = open("temp_pdf_text.txt", "w")
    for i in range(num_pages):
        page_text = pdf_reader.pages[i].extract_text()
        txt_file.write(page_text)
        # print(page_text)
    # pdf_file.close()
    txt_file.close()
    return "temp_pdf_text.txt"

stopwords= ['i', 'me', 'my', 'myself', 'we', 'our', 'ours', 'ourselves', 'you', "you're", "you've",\
            "you'll", "you'd", 'your', 'yours', 'yourself', 'yourselves', 'he', 'him', 'his', 'himself', \
            'she', "she's", 'her', 'hers', 'herself', 'it', "it's", 'its', 'itself', 'they', 'them', 'their',\
            'theirs', 'themselves', 'what', 'which', 'who', 'whom', 'this', 'that', "that'll", 'these', 'those', \
            'am', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'having', 'do', 'does', \
            'did', 'doing', 'a', 'an', 'the', 'and', 'but', 'if', 'or', 'because', 'as', 'until', 'while', 'of', \
            'at', 'by', 'for', 'with', 'about', 'against', 'between', 'into', 'through', 'during', 'before', 'after',\
            'above', 'below', 'to', 'from', 'up', 'down', 'in', 'out', 'on', 'off', 'over', 'under', 'again', 'further',\
            'then', 'once', 'here', 'there', 'when', 'where', 'why', 'how', 'all', 'any', 'both', 'each', 'few', 'more',\
            'most', 'other', 'some', 'such', 'only', 'own', 'same', 'so', 'than', 'too', 'very', \
            's', 't', 'can', 'will', 'just', 'don', "don't", 'should', "should've", 'now', 'd', 'll', 'm', 'o', 're', \
            've', 'y', 'ain', 'aren', "aren't", 'couldn', "couldn't", 'didn', "didn't", 'doesn', "doesn't", 'hadn',\
            "hadn't", 'hasn', "hasn't", 'haven', "haven't", 'isn', "isn't", 'ma', 'mightn', "mightn't", 'mustn',\
            "mustn't", 'needn', "needn't", 'shan', "shan't", 'shouldn', "shouldn't", 'wasn', "wasn't", 'weren', "weren't", \
            'won', "won't", 'wouldn', "wouldn't"]

def clean_pdf(pdf_txt_path):
    with open(pdf_txt_path, 'r', encoding='utf-8') as file:
        texts = file.readlines()

    cleaned_lines = []
    for text in texts:
        text = text.strip()

        if not text:
            continue

        text = re.sub(r'\b\w\b', '', text)
        text = re.sub(r"\\", " ", text)
        text = re.sub(r"\/", " ", text)
        text = re.sub(r"[\n\t\-]", " ", text)
        text = re.sub(r"won't", "will not", text)
        text = re.sub(r"won’t", "will not", text)
        text = re.sub(r"can\'t", "can not", text)
        text = re.sub(r"can\’t", "can not", text)
        text = re.sub(r"n\'t", " not", text)
        text = re.sub(r"n\’t", " not", text)
        text = re.sub(r"\'re", " are", text)
        text = re.sub(r"\’re", " are", text)
        text = re.sub(r"\'s", " is", text)
        text = re.sub(r"\’s", " is", text)
        text = re.sub(r"\'d", " would", text)
        text = re.sub(r"\’d", " would", text)
        text = re.sub(r"\'ll", " will", text)
        text = re.sub(r"\’ll", " will", text)
        text = re.sub(r"\'t", " not", text)
        text = re.sub(r"\’t", " not", text)
        text = re.sub(r"\'ve", " have", text)
        text = re.sub(r"\’ve", " have", text)
        text = re.sub(r"\’m", " am", text)
        text = re.sub(r"\'m", " am", text)
        text = re.sub(r'[^A-Za-z_]', " ", text)
        text = re.sub(r"\s+", " ", text)
        text = ' '.join(e for e in text.split() if e.lower() not in stopwords)
        cleaned_lines.append(text.lower())

    cleaned_text = "\n".join(cleaned_lines)
    if os.path.exists(pdf_txt_path):
        os.remove(pdf_txt_path)
    return cleaned_text



@app.route("/compare_pdfs", methods=['POST', 'GET'])
@cross_origin()
def compare_pdfs():
    # Load the BERT model
    data = request.files['SOC2']
    # print(data.read())
    file_name = data.filename
    pdf_path = os.path.join("./temp/", file_name)
    data.save(pdf_path)
    output_text_file = os.path.join("output/", str(datetime.now()))
    
    pdf_txt_path = convert_pdf_to_text(data)
    text = clean_pdf(pdf_txt_path)
    # images = convert_from_path(pdf_path)
    # text = ""
    # for image in images:
        # text += pytesseract.image_to_string(image, lang='eng')
    # with open(output_text_file, 'w') as file:
        # file.write(text)

    model = SentenceTransformer('bert-base-nli-mean-tokens')

    # Load the Excel sheet for sheet1
    sheet1 = pd.read_excel('/home/ps/github/beaconers/questionnaires/new reports/template.xlsx', header=None)

    # with open('/content/Atlassian-Platform-SOC2-Type-2_30-Sep-2021-1.txt', 'r') as file:
    #     texts_sheet2 = file.readlines()
    #texts_sheet2 = [preprocess(line) for line in texts_sh2]
    # texts_sheet2 = preprocess(texts_sheet2)

    texts_sheet1 = sheet1.values.flatten().tolist()
    texts_sheet2 = text.split('\n')

    print("TEXTS")
    # print(texts_sheet1)
    # print(texts_sheet2)

    embeddings_sheet1 = model.encode(texts_sheet1)
    embeddings_sheet2 = model.encode(texts_sheet2)

    matching_questions = []
    unmatched_questions = []

    for i, embedding_sheet1 in enumerate(embeddings_sheet1):
        max_similarity = 0.0
        max_similarity_index = -1
        for j, embedding_sheet2 in enumerate(embeddings_sheet2):
            similarity = cosine_similarity([embedding_sheet1], [embedding_sheet2])[0][0]
            if similarity > max_similarity:
                max_similarity = similarity
                max_similarity_index = j

        if max_similarity_index != -1:
            if max_similarity > 0.75:
                matching_questions.append((sheet1.iloc[i, 0], texts_sheet2[max_similarity_index], max_similarity, 'Yes'))
            else:
                matching_questions.append((sheet1.iloc[i, 0], texts_sheet2[max_similarity_index], max_similarity, 'No'))

            # matching_questions.append((sheet1.iloc[i, 0], texts_sheet2[max_similarity_index], max_similarity, 'Yes'))
        # else:
        #     unmatched_questions.append(sheet1.iloc[i, 0])

    matching_df = pd.DataFrame(matching_questions, columns=['Question 1', 'Question 2', 'Similarity', 'Matching'])


    # unmatched_df = pd.DataFrame({'Question 1': unmatched_questions, 'Question 2': '', 'Similarity': '', 'Matching': 'No'})
    # updated_df = matching_df.append(unmatched_df, ignore_index=True)
    updated_df = {'Question': matching_df['Question 1'].to_list(), 'Answer': matching_df['Matching'].to_list(), 'Similarity': matching_df['Similarity'].to_list()}

    excel_file = file_name.split(".")[0] + ".xlsx"
    matching_df.to_excel(excel_file, index=False)
    return json.dumps(updated_df, indent=4)

    # print("Unmatched Questions:")
    # for question in unmatched_questions:
    #     print(question)


# compare_questionnaires("/home/ps/github/beaconers/questionnaires/Snyk-Third Party Security & Privacy Questionnaire(4).xlsx")
# convert_pdf_to_text("/home/ps/github/beaconers/sco_2/CHEQ Paradome Security and Compliance.pdf", "output/test2.txt")

# driver function
if __name__ == '__main__':
    app.run(host="0.0.0.0", debug=True)

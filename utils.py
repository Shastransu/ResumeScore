import spacy
from pyresparser import ResumeParser
spacy.load("en_core_web_sm")
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
from docx import Document
import pandas as pd
from experience import Experience_extracter
#import re

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


#function for parsing images where PDF can't be parsed
def ImageParsing(TotalFileName,skillsRequired=None):
    pages = convert_from_path(TotalFileName, 500)
    page_counter = 1
    text=""
    for page in pages:
        resume_image = "Resume_page" + str(page_counter) + ".jpg"
        page.save(resume_image, 'JPEG')
        im = Image.open(resume_image)
        text = text.join(pytesseract.image_to_string(im))
        page_counter += 1

    #Converting text to docx file
    document = Document()
    document.add_paragraph(text)
    document.save('resume_doc.docx')
    # name = re.findall("(^[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*)|(^[A-Z][A-Za-z]*\s+[A-Z][A-Za-z]*)", text)
    #print(name)
    # custom regex works only for phone number
    extracted_data=ResumeParser("resume_doc.docx",custom_regex=r"(\+\d{2})?(\s)?(\d{3})(.)?(\d{3})(.)?(\d{4})",skills_file=skillsRequired).get_extracted_data()

    return extracted_data

#Extracting data from PDF extractable resume
def Resume_Parser(TotalFilename,skillsRequired=None):
    extracted_data = ResumeParser(TotalFilename,custom_regex=r"(\+\d{2})?(\s)?(\d{3})(.)?(\d{3})(.)?(\d{4})",skills_file=skillsRequired).get_extracted_data()

    if extracted_data["name"] is None:
        extracted_data=ImageParsing(TotalFilename,skillsRequired)
        # return extracted_data, skillsParsed

    skillsParsed = []
    extracted_skill_lower = [x.lower() for x in extracted_data["skills"]]
    # # print(extracted_skill_lower)
    # if skillsRequired == None:
    #     skillsParsed = extracted_skill_lower
    # else:
    #     for eachSkill in skillsRequired:
    #         if eachSkill in extracted_skill_lower:
    #             skillsParsed.append(eachSkill)
    return extracted_data,extracted_skill_lower


def RatingCalculator(extracted_data,RatingSkillsDict):
    RatingMarks = 0
    TotalRating = len(RatingSkillsDict)
    #prefered_Skils=[]

    # for skill in extracted_data["skills"]:
    #     if skill.lower() in RatingSkillsDict.keys():
    #         RatingMarks += 1
    #     for key,value in RatingSkillsDict.items():
    #         if skill.lower() in RatingSkillsDict.keys():
    #             if value==2:
    #                 if key not in prefered_Skils:
    #                     prefered_Skils.append(key)
    #TotalRating = len(RatingSkillsDict)
    prefered_skills = []
    main_skills = []
    for key, value in RatingSkillsDict.items():
        if value == 2:
            main_skills.append(key)
    for skill in extracted_data["skills"]:
        if skill.lower() in RatingSkillsDict.keys():
            RatingMarks += 1
        if skill.lower() in main_skills:
            prefered_skills.append(skill)

    Rating = round(RatingMarks / TotalRating * 10, 2)

    if Rating<3:
        MatchType="Not a good Match"
    elif Rating>3 and Rating<6:
        MatchType = "Partially Match"
    elif Rating> 6 and Rating<9:
        MatchType= "Good Match"
    else:
        MatchType = "Perfect Match"

    return Rating, MatchType,prefered_skills

def RatingSkillsCounter(df, RatingSkills):
    RateSkillCounter = {}

    for RateSkill in RatingSkills:
        RateSkillCounter[RateSkill] = 0
        for skillList in df['Skills']:
            # skillList = skillList.lower()
            # skillList = skillList.replace(r'[', '')
            # skillList = skillList.replace(r']', '')
            # skillList = skillList.replace(r"'", '')
            # skillList = skillList.split(", ")
            # print(skillList)
            if RateSkill in skillList:
                RateSkillCounter[RateSkill] += 1

    df_RateSkillCounter = pd.DataFrame.from_dict(RateSkillCounter, orient='index',columns=['No_Of_Resume_With_Skills'])
    return df_RateSkillCounter

def ExperienceCalculator(TotalFilename):
    exp=Experience_extracter(TotalFilename)
    #print(exp,type(exp))
    total_experience=float(exp[0])
    type_=exp[1]
    if total_experience is None or total_experience <1 or type_ == "months":
        expType="Fresher"
        total_experience=total_experience/12
    elif total_experience <3 and type_ =="years":
        expType="Junior"
    elif total_experience <6 and type_ == "years":
        expType="Mid-Senior"
    elif total_experience >= 6 and type_ == "years":
        expType= "Leadership"
    else:
        expType="Fresher"
    return total_experience, expType

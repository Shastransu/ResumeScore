import os
import pandas as pd
from utils import Resume_Parser,RatingSkillsCounter
from utils import RatingCalculator,ExperienceCalculator
import Constants
import pandas as pd
import re

def Resume_extract(RatingSkillsDict):
    data_frame = pd.DataFrame(columns=['Name', 'PhoneNumber', 'Email', 'Skills', 'Prefered Skills', 'Experience','Experience Type', 'Rating'])

    # Traversing through the files in the folder
    filenum = 1
    for file in filenames:
        TotalFilename = folderName + "\\" + file
        skillFile=r"C:\Users\Shastransu\Desktop\Skills List.csv"
        extracted_data,skills = Resume_Parser(TotalFilename,skillsRequired=skillFile)
        Rating,MatchType,prefered_skills = RatingCalculator(extracted_data, RatingSkillsDict)
        total_exp ,expType = ExperienceCalculator(TotalFilename)
        #creating dataframe
        data_frame = data_frame.append({'Name': extracted_data["name"], 'PhoneNumber': extracted_data["mobile_number"],
                          'Email': extracted_data["email"], 'Skills': skills, 'Prefered Skills':prefered_skills,
                          'Experience': total_exp,'Experience Type': expType, 'Rating': Rating, 'Profile Match Type': MatchType}, ignore_index=True)
        print("Sl. no."+str(filenum)+"\t FileName: "+file+"\t Status : Information Extraction Complete")
        filenum+=1
    df_RatingSkills = RatingSkillsCounter(data_frame, RatingSkills)
    print("\n No. of Resumes with Required skills in the folder")
    print(df_RatingSkills)

    #convert dataframe to excel
    ExportExcel="output.xlsx"
    writer = pd.ExcelWriter(ExportExcel, engine='openpyxl')
    data_frame.to_excel(writer,sheet_name="Resume Summary")
    df_RatingSkills.to_excel(writer,sheet_name="RatingSkills")
    writer.save()
    writer.close()


if __name__=="__main__":
    folderName=input("Enter the path for resumes:")
    # folderName = r"C:\Users\Shastransu\PycharmProjects\Resume\Resumes"
    cwd = os.path.abspath(folderName)
    filenames = os.listdir(cwd)

    RatingSkillsNum = int(input("Enter the number of Rating skills you need."))
    RatingSkills=[]
    print("Enter the skills Required")
    for each in range(RatingSkillsNum):
        skill=input()
        RatingSkills.append(skill.lower())

    RatingSkillsDict = {i: 1 for i in RatingSkills}
    extra_preference = input("Do you want to give extra preference to any of the above given skills(y/n):")


    if extra_preference.lower() == "y":
        var = int(input("Number of skills"))
        for index in range(var):
            imp_skill = input("enter the skill: ")
            if imp_skill in RatingSkills:
                RatingSkillsDict[imp_skill] = 2

    else:
        pass

    # skills_file=open(r"C:\Users\Shastransu\PycharmProjects\Resume\skillsRequired","r")
    # RatingSkills=skills_file.read().split(",")
    print("Analyzing Resumes on the basis of following skills: ", RatingSkills)

    # print(RatingSkills)
    Resume_extract(RatingSkillsDict)


